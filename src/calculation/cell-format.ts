import { formatCalendar } from './calendar.js';
import { quantityInUnit } from './units.js';
import { Decimal } from 'decimal.js';
import { formatCellValue, type DisplayOptions } from './format.js';
import { isNumeric, type CellValue } from './values.js';
import type { CellFormat } from '../model/review.js';

export function formatWorksheetCell(
  value: CellValue,
  format?: CellFormat,
  options?: DisplayOptions,
): string {
  if (!format || format.kind === 'general' || !isNumeric(value))
    return formatCellValue(value, options);
  if (format.kind === 'date' || format.kind === 'time' || format.kind === 'datetime') {
    try {
      return formatCalendar(value, format.kind);
    } catch {
      return '#VALUE!';
    }
  }
  if (format.kind === 'unit') {
    try {
      const displayed = quantityInUnit(value, format.unit!);
      return `${format.decimals === undefined ? formatCellValue(displayed, options) : formatWorksheetCell(displayed, { kind: 'number', decimals: format.decimals }, options)} ${format.unit}`;
    } catch {
      return '#VALUE!';
    }
  }
  if (format.kind === 'fraction')
    return formatCellValue(value, { ...options, notation: 'fraction' });
  const decimals =
    format.decimals ?? (format.kind === 'number' || format.kind === 'currency' ? 2 : 0);
  if (!Number.isInteger(decimals) || decimals < 0 || decimals > 100)
    throw new RangeError('Invalid decimal places.');
  const D = Decimal.clone({
    precision: Math.min(
      100_100,
      Math.max(value.numerator.length, value.denominator.length, decimals) + 110,
    ),
    rounding: Decimal.ROUND_HALF_EVEN,
  });
  let decimal = new D(value.numerator).div(value.denominator);
  if (format.kind === 'percent') decimal = decimal.mul(100);
  const text =
    format.kind === 'scientific' ? decimal.toExponential(decimals) : decimal.toFixed(decimals);
  return `${value.approximate && options?.approximationMarker !== false ? '≈' : ''}${format.kind === 'currency' ? '$' : ''}${text}${format.kind === 'percent' ? '%' : ''}`;
}
