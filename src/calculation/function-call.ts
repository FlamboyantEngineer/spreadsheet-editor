import { formatCellValue } from './format.js';
import type { createQuantityArithmetic } from './units.js';
import {
  CalculationError,
  type CellValue,
  isError,
  isNumeric,
  type NumericValue,
} from './values.js';

/** A formula value: one cell, or the rows of a range. */
export type Value = CellValue | CellValue[][];
export type Arithmetic = ReturnType<typeof createQuantityArithmetic>;

/** One call of a function whose arguments are evaluated before it runs. */
export interface FunctionCall {
  readonly name: string;
  readonly args: readonly Value[];
  readonly math: Arithmetic;
  readonly compare: (a: CellValue, b: CellValue) => number;
}
export type FunctionImpl = (call: FunctionCall) => Value;

export const scalar = (value: Value): CellValue =>
  Array.isArray(value) ? (value[0]?.[0] ?? null) : value;
export const flat = (value: Value): CellValue[] => (Array.isArray(value) ? value.flat() : [value]);
export const matrix = (value: Value): CellValue[][] => (Array.isArray(value) ? value : [[value]]);
export const fail = (
  code: ConstructorParameters<typeof CalculationError>[0] = '#VALUE!',
): never => {
  throw new CalculationError(code);
};
export function checked(value: CellValue): CellValue {
  if (isError(value)) throw new CalculationError(value.code);
  return value;
}

/** Fails with `#VALUE!` unless the call has between `min` and `max` arguments. */
export function arity(call: FunctionCall, min: number, max = min): void {
  if (call.args.length < min || call.args.length > max) fail();
}
export const zeroOf = (call: FunctionCall) => call.math.parse('0');
export const oneOf = (call: FunctionCall) => call.math.parse('1');

export function number(call: FunctionCall, value: Value): NumericValue {
  const item = checked(scalar(value));
  if (isNumeric(item)) return item;
  if (item === null || item === false) return zeroOf(call);
  if (item === true) return oneOf(call);
  return call.math.parse(String(item));
}
export function integer(
  call: FunctionCall,
  value: Value,
  [min, max]: readonly [number, number] = [-1000000, 1000000],
): number {
  return call.math.integer(number(call, value), min, max);
}
export const text = (value: Value) =>
  formatCellValue(checked(scalar(value)), { approximationMarker: false });
/** Truth of an argument; text other than TRUE is false. */
export function truth(value: Value): boolean {
  const item = checked(scalar(value));
  if (typeof item === 'boolean') return item;
  if (item === null) return false;
  if (isNumeric(item)) return item.numerator !== '0';
  return String(item).toUpperCase() === 'TRUE';
}
/** Whether an optional flag argument is present and non-zero. */
export const flagged = (call: FunctionCall, value: Value | undefined) =>
  value !== undefined && number(call, value).numerator !== '0';
