import { excelColor } from './colors.js';
import type ExcelJS from 'exceljs';
import type { CellFormat } from '../model/review.js';
import { UNIT_PRESETS } from '../calculation/units.js';

export function importNumberFormat(code: string): CellFormat | null {
  if (!code || code === 'General' || code === '@') return { kind: 'general' };
  const suffix = /"([^"\n]+)"\s*$/.exec(code)?.[1].trim();
  const decimals = /\.([0#]+)/.exec(code)?.[1].length ?? 0;
  if (suffix && UNIT_PRESETS.includes(suffix))
    return { kind: 'unit', unit: suffix, inputUnit: suffix, decimals };
  if (/[dy]/i.test(code)) return { kind: /[hs]/i.test(code) ? 'datetime' : 'date' };
  if (/[hs]/i.test(code)) return { kind: 'time' };
  if (code.includes('%')) return { kind: 'percent', decimals };
  if (code.includes('$') || /\[\$USD/.test(code)) return { kind: 'currency', decimals };
  if (/[eE][+-]/.test(code)) return { kind: 'scientific', decimals };
  if (/^[#0.,;()\-+\s]+$/.test(code)) return { kind: 'number', decimals };
  return null;
}
export function exportNumberFormat(format?: CellFormat): string {
  if (!format || format.kind === 'general' || format.kind === 'fraction') return 'General';
  if (format.kind === 'date') return 'mmm d, yyyy';
  if (format.kind === 'time') return 'hh:mm:ss';
  if (format.kind === 'datetime') return 'mmm d, yyyy hh:mm:ss';
  const places =
      format.decimals ?? (format.kind === 'number' || format.kind === 'currency' ? 2 : 0),
    decimal = places ? '.' + '0'.repeat(places) : '';
  return `${format.kind === 'currency' ? '$' : ''}0${decimal}${format.kind === 'percent' ? '%' : format.kind === 'scientific' ? 'E+00' : format.kind === 'unit' ? ` "${format.unit}"` : ''}`;
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
