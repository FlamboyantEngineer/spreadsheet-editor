import { Decimal } from 'decimal.js';
import { dimensionLabel } from './units.js';
import { type CellValue, isError, isNumeric, type NumericValue } from './values.js';

export interface DisplayOptions {
  significantDigits?: number;
  notation?: 'auto' | 'decimal' | 'fraction';
  approximationMarker?: boolean;
}

function plainText(value: Exclude<CellValue, NumericValue>): string {
  if (isError(value)) return value.code;
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return value;
}

function numberText(value: NumericValue, options: DisplayOptions): string {
  const digits = options.significantDigits ?? 32;
  if (!Number.isInteger(digits) || digits < 1 || digits > 10_000)
    throw new RangeError('Invalid display precision.');
  if (options.notation === 'fraction')
    return value.denominator === '1' ? value.numerator : `${value.numerator}/${value.denominator}`;
  const D = Decimal.clone({ precision: digits + 8, rounding: Decimal.ROUND_HALF_EVEN });
  const number = new D(value.numerator).div(value.denominator).toSignificantDigits(digits);
  return options.notation === 'decimal' ? number.toFixed() : number.toString();
}

export function formatCellValue(value: CellValue, options: DisplayOptions = {}): string {
  if (!isNumeric(value)) return plainText(value);
  const text = numberText(value, options);
  const rendered = value.approximate && options.approximationMarker !== false ? `≈${text}` : text;
  const unit = dimensionLabel(value);
  return unit ? `${rendered} ${unit}` : rendered;
}
