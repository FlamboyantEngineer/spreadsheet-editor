import type ExcelJS from 'exceljs';
import { resolveUnit } from '../calculation/units.js';
import type { CellFormat } from '../model/review.js';
import { excelColor } from './colors.js';

const LOCALE_IDS: Record<string, string> = {
  'en-US': '409',
  'de-DE': '407',
  'en-GB': '809',
  'fr-FR': '40C',
  'es-ES': 'C0A',
  'it-IT': '410',
  'ja-JP': '411',
  'ko-KR': '412',
  'zh-CN': '804',
  'zh-TW': '404',
  'pt-BR': '416',
  'nl-NL': '413',
  'sv-SE': '41D',
  'da-DK': '406',
  'fi-FI': '40B',
  'nb-NO': '414',
};
const SYMBOL_CURRENCIES: Record<string, string> = { $: 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY' };

/** The locale and currency of a `[$CUR-LCID]` tag; null when the locale ID is not supported. */
function formatTag(code: string): { locale?: string; currency?: string } | null {
  const tag = /\[\$([A-Z]{3}|[$€£¥]?)-([0-9a-f]+)\]/i.exec(code);
  if (!tag) return {};
  const id = tag[2].replace(/^0+/, '').toUpperCase();
  const locale = Object.keys(LOCALE_IDS).find((key) => LOCALE_IDS[key] === id);
  if (!locale) return null;
  const currency = tag[1] ? (SYMBOL_CURRENCIES[tag[1]] ?? tag[1].toUpperCase()) : undefined;
  return { locale, ...(currency ? { currency } : {}) };
}

/** A quoted suffix is a unit format only when the unit registry recognizes it. */
function unitFormat(suffix: string | undefined, decimals: number): CellFormat | undefined {
  if (!suffix) return undefined;
  try {
    resolveUnit(suffix);
    return { kind: 'unit', unit: suffix, inputUnit: suffix, decimals };
  } catch {
    return undefined;
  }
}

/** The currency a body's symbol names: euro, pound or yen, else dollar. */
const symbolCurrency = (body: string) => SYMBOL_CURRENCIES[/[$€£¥]/.exec(body)?.[0] ?? '$'];

/** A currency format when the tag names a currency or the body shows a currency symbol. */
function currencyFormat(
  body: string,
  { currency, decimals }: { currency: string | undefined; decimals: number },
): CellFormat | undefined {
  if (!currency && !/[$€£¥]/.test(body)) return undefined;
  return {
    kind: 'currency',
    decimals,
    currency: currency ?? symbolCurrency(body),
  };
}

interface FormatSection {
  pattern: string;
  literal: string | undefined;
  hasModifiers: boolean;
  elapsed: boolean;
}

/** The positive/default section. Quoted and escaped text and Excel controls are not field tokens. */
function formatSection(code: string): FormatSection {
  const section: FormatSection = {
    pattern: '',
    literal: undefined,
    hasModifiers: false,
    elapsed: false,
  };
  for (const match of code.matchAll(/"[^"]*"|[\\_*][\s\S]|\[[^\]]*\]|[\s\S]/g)) {
    const token = match[0];
    if (token === ';') {
      section.hasModifiers = true;
      break;
    }
    if (token.startsWith('"')) {
      if (section.literal !== undefined) section.hasModifiers = true;
      section.literal = token.slice(1, -1).trim();
    } else if (token.length > 1) {
      section.hasModifiers = true;
      section.elapsed ||= /^\[(?:hh?|mm?|ss?)\]$/i.test(token);
    } else section.pattern += token;
  }
  return section;
}

/** A month field by itself is a date, and alongside hours or seconds it represents minutes. */
function calendarFormat(body: string): CellFormat | undefined {
  const fields = body.replace(/am\/pm|a\/p/gi, 'h');
  if (!/^[ymdhs/:.,\s-]+$/i.test(fields) || !/[ymdhs]/i.test(fields)) return undefined;
  const time = /[hs]/i.test(fields),
    date = /[dy]/i.test(fields);
  return { kind: time ? (date ? 'datetime' : 'time') : 'date' };
}

/** The format a code's fields describe, before its locale tag is applied. */
function bodyFormat(
  body: string,
  { currency, literal }: { currency?: string; literal?: string },
): CellFormat | null {
  if (body === 'General' || body === '@') return { kind: 'general' };
  if (body === '# ?/?') return { kind: 'fraction' };
  const decimals = Math.min(100, /\.([0#]+)/.exec(body)?.[1].length ?? 0);
  const calendar = calendarFormat(body);
  if (calendar) return calendar;
  if (body.includes('%')) return { kind: 'percent', decimals };
  const priced = currencyFormat(body + (literal ?? ''), { currency, decimals });
  if (priced) return priced;
  if (/[eE][+-]/.test(body)) return { kind: 'scientific', decimals };
  return /^[#0.,()\-+\s]+$/.test(body)
    ? (unitFormat(literal, decimals) ?? { kind: 'number', decimals })
    : null;
}

export function importNumberFormat(code: string | undefined): {
  format: CellFormat | null;
  lossy: boolean;
} {
  if (!code || code === 'General' || code === '@')
    return { format: { kind: 'general' }, lossy: false };
  const metadata = formatTag(code);
  if (!metadata) return { format: null, lossy: true };
  const section = formatSection(code.replace(/\[\$[^\]]*\]/g, ''));
  const format = section.elapsed
    ? null
    : bodyFormat(section.pattern.trim(), { ...metadata, literal: section.literal });
  return {
    format: format && { ...format, ...metadata },
    lossy: section.hasModifiers || Boolean(section.literal && format?.kind !== 'unit'),
  };
}
/** Excel only has a finite locale-ID vocabulary; unsupported locale metadata is reported, not invented. */
export function numberFormatLoss(format?: CellFormat): string | null {
  if (format?.locale && !Object.hasOwn(LOCALE_IDS, format.locale))
    return `XLSX cannot retain locale ${format.locale}; its default number-format locale is used.`;
  if (format?.currency && format.kind !== 'currency')
    return 'XLSX only retains currency metadata on a currency number format.';
  if (format && (format.inputUnit || format.expectedDimension || format.quantityKind))
    return 'XLSX does not retain input-unit, expected-dimension or quantity-kind metadata.';
  return null;
}
const FIXED_CODES: Partial<Record<CellFormat['kind'], string>> = {
  general: 'General',
  fraction: '# ?/?',
  date: 'mmm d, yyyy',
  time: 'hh:mm:ss',
  datetime: 'mmm d, yyyy hh:mm:ss',
};

function localePrefix(format: CellFormat): string {
  const locale = LOCALE_IDS[format.locale ?? 'en-US'];
  if (format.kind === 'currency') return `[$${format.currency ?? 'USD'}-${locale ?? '409'}]`;
  return format.locale && locale ? `[$-${locale}]` : '';
}

function numberSuffix(format: CellFormat): string {
  if (format.kind === 'percent') return '%';
  if (format.kind === 'scientific') return 'E+00';
  return format.kind === 'unit' ? ` "${format.unit}"` : '';
}

export function exportNumberFormat(format?: CellFormat): string {
  if (!format) return 'General';
  const prefix = localePrefix(format);
  const fixed = FIXED_CODES[format.kind];
  if (fixed) return prefix + fixed;
  const grouped = format.kind === 'number' || format.kind === 'currency';
  const places = format.decimals ?? (grouped ? 2 : 0),
    decimal = places ? '.' + '0'.repeat(places) : '';
  return `${prefix}${grouped ? '#,##' : ''}0${decimal}${numberSuffix(format)}`;
}

type Color = (value: Parameters<typeof excelColor>[0]) => string | undefined;

function importFont(values: Record<string, string>, font: Partial<ExcelJS.Font>, color: Color) {
  if (font.bold) values['font-weight'] = 'bold';
  if (font.italic) values['font-style'] = 'italic';
  if (font.strike) values['text-decoration'] = 'line-through';
  if (font.underline) values['text-decoration'] = 'underline';
  if (font.name && font.name !== 'Calibri') values['font-family'] = font.name;
  if (font.size && font.size !== 11) values['font-size'] = `${(font.size * 4) / 3}px`;
  const foreground = color(font.color);
  if (foreground) values.color = foreground;
}

function importAlignment(values: Record<string, string>, alignment: Partial<ExcelJS.Alignment>) {
  if (alignment.horizontal) values['text-align'] = alignment.horizontal;
  if (alignment.vertical) values['vertical-align'] = alignment.vertical;
  if (alignment.wrapText) values['white-space'] = 'pre-wrap';
  if (alignment.indent) values['text-indent'] = `${alignment.indent}em`;
}

const SIDES = ['top', 'right', 'bottom', 'left'] as const;

/** Map an Excel cell style to the model's style properties; the caller serializes them. */
export function importCellStyle(
  style: Partial<ExcelJS.Style>,
  theme: readonly string[] = [],
): Record<string, string> {
  const color: Color = (value) => excelColor(value, theme);
  const values: Record<string, string> = {};
  importFont(values, style.font ?? {}, color);
  const fill = style.fill;
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const background = color(fill.fgColor);
    if (background) values['background-color'] = background;
  }
  importAlignment(values, style.alignment ?? {});
  for (const side of SIDES) {
    const border = style.border?.[side];
    if (border?.style) values[`border-${side}`] = `1px solid ${color(border.color) ?? '#000000'}`;
  }
  return values;
}

const argb = (color: string) => ({ argb: 'FF' + color.replace('#', '') });

function exportFont(properties: Record<string, string>): Partial<ExcelJS.Font> {
  const font: Partial<ExcelJS.Font> = {};
  if (properties['font-weight'] === 'bold') font.bold = true;
  if (properties['font-style'] === 'italic') font.italic = true;
  if (properties['text-decoration'] === 'line-through') font.strike = true;
  if (properties['text-decoration'] === 'underline') font.underline = true;
  if (properties['font-family']) font.name = properties['font-family'];
  if (properties['font-size']) font.size = (Number.parseFloat(properties['font-size']) * 3) / 4;
  if (properties.color) font.color = argb(properties.color);
  return font;
}

function exportAlignment(properties: Record<string, string>): Partial<ExcelJS.Alignment> {
  const alignment: Partial<ExcelJS.Alignment> = {
    horizontal: (properties['text-align'] ?? 'center') as ExcelJS.Alignment['horizontal'],
  };
  if (properties['vertical-align'])
    alignment.vertical = properties['vertical-align'] as ExcelJS.Alignment['vertical'];
  if (properties['white-space'] === 'pre-wrap') alignment.wrapText = true;
  return alignment;
}

export function exportCellStyle(properties: Record<string, string>): Partial<ExcelJS.Style> {
  const border: Partial<ExcelJS.Borders> = {};
  for (const side of SIDES) {
    const value = properties[`border-${side}`];
    if (!value) continue;
    const match = /#[a-f0-9]{6}/i.exec(value);
    border[side] = { style: 'thin', color: { argb: 'FF' + (match?.[0].slice(1) ?? '000000') } };
  }
  const background = properties['background-color'];
  return {
    font: exportFont(properties),
    alignment: exportAlignment(properties),
    border,
    ...(background
      ? { fill: { type: 'pattern', pattern: 'solid', fgColor: argb(background) } as ExcelJS.Fill }
      : {}),
  };
}
