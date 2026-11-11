import { excelColor } from './colors.js';
import type ExcelJS from 'exceljs';
import type { CellFormat } from '../model/review.js';
import { resolveUnit } from '../calculation/units.js';

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
type LocalizedFormat = CellFormat & { locale?: string; currency?: string };
export function importNumberFormat(code: string): CellFormat | null {
  if (!code || code === 'General' || code === '@') return { kind: 'general' };
  const tag = /\[\$([A-Z]{3}|[$€£¥]?)-([0-9a-f]+)\]/i.exec(code);
  const locale =
    tag &&
    Object.keys(LOCALE_IDS).find(
      (key) => LOCALE_IDS[key] === tag[2].replace(/^0+/, '').toUpperCase(),
    );
  if (tag && !locale) return null;
  const currency =
    tag && tag[1]
      ? ({ $: 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY' }[tag[1]] ?? tag[1].toUpperCase())
      : undefined;
  const body = code.replace(/\[\$[^\]]*\]/g, '');
  const metadata = { ...(locale ? { locale } : {}), ...(currency ? { currency } : {}) };
  const result = (value: CellFormat): CellFormat => ({ ...value, ...metadata });
  if (body === 'General' || body === '@') return result({ kind: 'general' });
  if (body === '# ?/?') return result({ kind: 'fraction' });
  const suffix = /"([^"\n]+)"\s*$/.exec(body)?.[1].trim();
  const decimals = /\.([0#]+)/.exec(body)?.[1].length ?? 0;
  if (suffix) {
    try {
      resolveUnit(suffix);
      return result({ kind: 'unit', unit: suffix, inputUnit: suffix, decimals });
    } catch {
      /* A quoted suffix is only a unit when the registry recognizes it. */
    }
  }
  if (/[dy]/i.test(body)) return result({ kind: /[hs]/i.test(body) ? 'datetime' : 'date' });
  if (/[hs]/i.test(body)) return result({ kind: 'time' });
  if (body.includes('%')) return result({ kind: 'percent', decimals });
  if (currency || /[$€£¥]/.test(body))
    return result({
      kind: 'currency',
      decimals,
      ...(!currency
        ? {
            currency: body.includes('€')
              ? 'EUR'
              : body.includes('£')
                ? 'GBP'
                : body.includes('¥')
                  ? 'JPY'
                  : 'USD',
          }
        : {}),
    } as LocalizedFormat);
  if (/[eE][+-]/.test(body)) return result({ kind: 'scientific', decimals });
  if (/^[#0.,;()\-+\s]+$/.test(body)) return result({ kind: 'number', decimals });
  return null;
}
/** Excel only has a finite locale-ID vocabulary; unsupported locale metadata is reported, not invented. */
export function numberFormatLoss(format?: LocalizedFormat): string | null {
  if (format?.locale && !Object.hasOwn(LOCALE_IDS, format.locale))
    return `XLSX cannot retain locale ${format.locale}; its default number-format locale is used.`;
  if (format?.currency && format.kind !== 'currency')
    return 'XLSX only retains currency metadata on a currency number format.';
  if (format && (format.inputUnit || format.expectedDimension || format.quantityKind))
    return 'XLSX does not retain input-unit, expected-dimension or quantity-kind metadata.';
  return null;
}
export function exportNumberFormat(format?: LocalizedFormat): string {
  if (!format) return 'General';
  const locale = LOCALE_IDS[format.locale ?? 'en-US'];
  const prefix =
    format.kind === 'currency'
      ? `[$${format.currency ?? 'USD'}-${locale ?? '409'}]`
      : format.locale && locale
        ? `[$-${locale}]`
        : '';
  if (format.kind === 'general') return prefix + 'General';
  if (format.kind === 'fraction') return prefix + '# ?/?';
  if (format.kind === 'date') return prefix + 'mmm d, yyyy';
  if (format.kind === 'time') return prefix + 'hh:mm:ss';
  if (format.kind === 'datetime') return prefix + 'mmm d, yyyy hh:mm:ss';
  const places =
      format.decimals ?? (format.kind === 'number' || format.kind === 'currency' ? 2 : 0),
    decimal = places ? '.' + '0'.repeat(places) : '';
  return `${prefix}${format.kind === 'number' || format.kind === 'currency' ? '#,##' : ''}0${decimal}${format.kind === 'percent' ? '%' : format.kind === 'scientific' ? 'E+00' : format.kind === 'unit' ? ` "${format.unit}"` : ''}`;
}
export function importCellStyle(
  style: Partial<ExcelJS.Style>,
  theme: readonly string[] = [],
): string {
  const color = (value: Parameters<typeof excelColor>[0]) => excelColor(value, theme);
  const values: Record<string, string> = {},
    font = style.font;
  if (font?.bold) values['font-weight'] = 'bold';
  if (font?.italic) values['font-style'] = 'italic';
  if (font?.strike) values['text-decoration'] = 'line-through';
  if (font?.underline) values['text-decoration'] = 'underline';
  if (font?.name && font.name !== 'Calibri') values['font-family'] = font.name;
  if (font?.size && font.size !== 11) values['font-size'] = `${(font.size * 4) / 3}px`;
  const foreground = color(font?.color);
  if (foreground) values.color = foreground;
  const fill = style.fill;
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const background = color(fill.fgColor);
    if (background) values['background-color'] = background;
  }
  if (style.alignment?.horizontal) values['text-align'] = style.alignment.horizontal;
  if (style.alignment?.vertical) values['vertical-align'] = style.alignment.vertical;
  if (style.alignment?.wrapText) values['white-space'] = 'pre-wrap';
  if (style.alignment?.indent) values['text-indent'] = `${style.alignment.indent}em`;
  for (const side of ['top', 'right', 'bottom', 'left'] as const)
    if (style.border?.[side]?.style)
      values[`border-${side}`] = `1px solid ${color(style.border[side]?.color) ?? '#000000'}`;
  return Object.entries(values)
    .map(([key, value]) => `${key}:${value}`)
    .join(';');
}
export function exportCellStyle(properties: Record<string, string>): Partial<ExcelJS.Style> {
  const font: Partial<ExcelJS.Font> = {},
    alignment: Partial<ExcelJS.Alignment> = {},
    border: Partial<ExcelJS.Borders> = {};
  if (properties['font-weight'] === 'bold') font.bold = true;
  if (properties['font-style'] === 'italic') font.italic = true;
  if (properties['text-decoration'] === 'line-through') font.strike = true;
  if (properties['text-decoration'] === 'underline') font.underline = true;
  if (properties['font-family']) font.name = properties['font-family'];
  if (properties['font-size']) font.size = (Number.parseFloat(properties['font-size']) * 3) / 4;
  if (properties.color) font.color = { argb: 'FF' + properties.color.replace('#', '') };
  alignment.horizontal = (properties['text-align'] ?? 'center') as ExcelJS.Alignment['horizontal'];
  if (properties['vertical-align'])
    alignment.vertical = properties['vertical-align'] as ExcelJS.Alignment['vertical'];
  if (properties['white-space'] === 'pre-wrap') alignment.wrapText = true;
  for (const side of ['top', 'right', 'bottom', 'left'] as const)
    if (properties[`border-${side}`]) {
      const match = /#[a-f0-9]{6}/i.exec(properties[`border-${side}`]);
      border[side] = { style: 'thin', color: { argb: 'FF' + (match?.[0].slice(1) ?? '000000') } };
    }
  return {
    font,
    alignment,
    border,
    ...(properties['background-color']
      ? {
          fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF' + properties['background-color'].replace('#', '') },
          } as ExcelJS.Fill,
        }
      : {}),
  };
}
