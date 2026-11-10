import { Decimal } from 'decimal.js';
import { dimensionLabel } from './units.js';
import { isError, isNumeric, type CellValue } from './values.js';

export interface DisplayOptions {
  significantDigits?: number;
  notation?: 'auto' | 'decimal' | 'fraction';
  approximationMarker?: boolean;
}

export function formatCellValue(value: CellValue, options: DisplayOptions = {}): string {
  if (isError(value)) return value.code;
  if (!isNumeric(value))
    return value == null ? '' : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value;
  const digits = options.significantDigits ?? 32;
  if (!Number.isInteger(digits) || digits < 1 || digits > 10_000)
    throw new RangeError('Invalid display precision.');
  const D = Decimal.clone({ precision: digits + 8, rounding: Decimal.ROUND_HALF_EVEN });
  const number = new D(value.numerator).div(value.denominator).toSignificantDigits(digits);
  const text =
    options.notation === 'fraction'
      ? value.denominator === '1'
        ? value.numerator
        : `${value.numerator}/${value.denominator}`
      : options.notation === 'decimal'
        ? number.toFixed()
        : number.toString();
  const rendered = value.approximate && options.approximationMarker !== false ? `≈${text}` : text;
  const unit = dimensionLabel(value);
  return unit ? `${rendered} ${unit}` : rendered;
}
