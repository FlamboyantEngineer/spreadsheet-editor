import { formatCellValue } from './format.js';
import type { Arithmetic, Value } from './function-call.js';
import {
  CalculationError,
  type CellValue,
  type ErrorValue,
  isError,
  isNumeric,
  type NumericValue,
  TEXT_COLLATOR,
} from './values.js';

/** How the built-in functions and operators read their operands. */
export interface Coercion {
  readonly zero: NumericValue;
  readonly one: NumericValue;
  scalar(value: Value): CellValue;
  checked(value: CellValue): Exclude<CellValue, ErrorValue>;
  numeric(value: Value): NumericValue;
  truth(value: Value): boolean;
  compare(a: CellValue, b: CellValue): number;
  text(value: Value): string;
  flattened(values: readonly Value[]): CellValue[];
}

const scalar = (value: Value): CellValue =>
  Array.isArray(value) ? (value[0]?.[0] ?? null) : value;
function checked(value: CellValue): Exclude<CellValue, ErrorValue> {
  if (isError(value)) throw new CalculationError(value.code);
  return value;
}
/** Text, booleans, then numbers; an empty cell takes the kind of what it is compared with. */
const rank = (value: CellValue) =>
  typeof value === 'boolean' ? 2 : typeof value === 'string' ? 1 : 0;
function blankAs(other: CellValue, zero: NumericValue): CellValue {
  if (typeof other === 'string') return '';
  return typeof other === 'boolean' ? false : zero;
}

export function createCoercion(math: Arithmetic): Coercion {
  const zero = math.parse('0'),
    one = math.parse('1');
  function numeric(value: Value): NumericValue {
    const item = checked(scalar(value));
    if (isNumeric(item)) return item;
    if (item === null) return zero;
    if (typeof item === 'boolean') return item ? one : zero;
    return math.parse(item);
  }
  function truth(value: Value): boolean {
    const item = checked(scalar(value));
    if (isNumeric(item)) return item.numerator !== '0';
    if (item === null) return false;
    if (typeof item === 'boolean') return item;
    const upper = item.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') return upper === 'TRUE';
    return numeric(item).numerator !== '0';
  }
  function compare(a: CellValue, b: CellValue): number {
    checked(a);
    checked(b);
    const av = a === null ? blankAs(b, zero) : a,
      bv = b === null ? blankAs(a, zero) : b;
    if (rank(av) !== rank(bv)) return rank(av) < rank(bv) ? -1 : 1;
    if (isNumeric(av) && isNumeric(bv)) return math.compare(av, bv);
    return TEXT_COLLATOR.compare(String(av).toUpperCase(), String(bv).toUpperCase());
  }
  return {
    zero,
    one,
    scalar,
    checked,
    numeric,
    truth,
    compare,
    text: (value) => formatCellValue(checked(scalar(value)), { approximationMarker: false }),
    flattened: (values) =>
      values.flatMap((value) => (Array.isArray(value) ? value.flat() : [value])),
  };
}
