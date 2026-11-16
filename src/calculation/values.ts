import type { Dimension } from './quantity-core.js';
/**
 * Text order for calculation: one fixed collation, so a client and a server compare and sort text the
 * same way whatever their runtime locale.
 */
export const TEXT_COLLATOR = new Intl.Collator('en-US');

export interface NumericValue {
  readonly dimension?: Dimension;
  readonly quantityKind?: string;
  readonly temperature?: 'absolute' | 'difference';
  readonly provenance?: Record<string, string>;
  readonly kind: 'number';
  readonly numerator: string;
  readonly denominator: string;
  readonly approximate: boolean;
  /** Working significant digits for an approximation; absent for exact values. */
  readonly precision?: number;
}

/** Every error a calculated cell can hold. */
export const FORMULA_ERROR_CODES = Object.freeze([
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#ERROR!',
  '#LIMIT!',
  '#CYCLE!',
] as const);
export type FormulaErrorCode = (typeof FORMULA_ERROR_CODES)[number];
export interface ErrorValue {
  readonly kind: 'error';
  readonly code: FormulaErrorCode;
}
export type CellValue = NumericValue | ErrorValue | string | boolean | null;

export function isNumeric(value: CellValue): value is NumericValue {
  return typeof value === 'object' && value !== null && value.kind === 'number';
}
export function isError(value: CellValue): value is ErrorValue {
  return typeof value === 'object' && value !== null && value.kind === 'error';
}

export class CalculationError extends Error {
  constructor(readonly code: FormulaErrorCode) {
    super(code);
  }
}

export function errorValue(code: FormulaErrorCode): ErrorValue {
  return { kind: 'error', code };
}

/** Finite, portable JSON for host-owned captured calculation context. */
export type JsonValue =
  null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };

export function assertJsonValue(
  value: unknown,
  maxBytes = 16 * 1024 * 1024,
): asserts value is JsonValue {
  const seen = new Set<object>();
  let bytes = 0;
  function visit(item: unknown, depth: number) {
    if (depth > 32) throw new TypeError('Host JSON exceeds maximum depth 32.');
    if (item === null || typeof item === 'boolean') {
      bytes += 5;
      return;
    }
    if (typeof item === 'number' && Number.isFinite(item)) {
      bytes += 24;
      return;
    }
    if (typeof item === 'string') {
      bytes += new TextEncoder().encode(JSON.stringify(item)).length;
      return;
    }
    if (!item || typeof item !== 'object' || seen.has(item))
      throw new TypeError('Host context must be finite acyclic JSON.');
    const prototype = Object.getPrototypeOf(item);
    if (!Array.isArray(item) && prototype !== Object.prototype && prototype !== null)
      throw new TypeError('Host context requires plain JSON objects.');
    seen.add(item);
    bytes += 2;
    for (const [key, child] of Object.entries(item)) {
      bytes += key.length * 3 + 4;
      visit(child, depth + 1);
      if (bytes > maxBytes) throw new RangeError('Host JSON exceeds the byte limit.');
    }
    seen.delete(item);
  }
  visit(value, 0);
  if (bytes > maxBytes) throw new RangeError('Host JSON exceeds the byte limit.');
}

const isPlainRecord = (value: unknown): value is Record<string, unknown> =>
  Boolean(value) &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(value) as object | null);

function validMagnitude(cell: NumericValue, maxDigits: number): boolean {
  return (
    cell.kind === 'number' &&
    typeof cell.numerator === 'string' &&
    typeof cell.denominator === 'string' &&
    cell.numerator.length <= maxDigits + 1 &&
    cell.denominator.length <= maxDigits &&
    /^-?(?:0|[1-9]\d*)$/.test(cell.numerator) &&
    /^[1-9]\d*$/.test(cell.denominator)
  );
}
function validPrecision(cell: NumericValue): boolean {
  if (typeof cell.approximate !== 'boolean') return false;
  if (cell.precision === undefined) return !cell.approximate;
  return Number.isInteger(cell.precision) && cell.precision >= 16 && cell.precision <= 1000;
}
const validDimension = ({ dimension }: NumericValue) =>
  dimension === undefined ||
  (Array.isArray(dimension) &&
    dimension.length === 7 &&
    Array.from(dimension).every((n) => Number.isInteger(n) && Math.abs(n) <= 32));
const validKind = ({ quantityKind }: NumericValue) =>
  quantityKind === undefined ||
  (typeof quantityKind === 'string' && Boolean(quantityKind) && quantityKind.length <= 128);
const validTemperature = (cell: NumericValue) =>
  cell.temperature === undefined ||
  (['absolute', 'difference'].includes(cell.temperature) && isPureTemperature(cell));
const validRevision = ([key, revision]: [string, unknown]) =>
  Boolean(key) &&
  key.length <= 1024 &&
  typeof revision === 'string' &&
  Boolean(revision) &&
  revision.length <= 1024;
const validProvenance = ({ provenance }: NumericValue) =>
  provenance === undefined ||
  (isPlainRecord(provenance) &&
    Object.keys(provenance).length <= 128 &&
    Object.entries(provenance).every(validRevision));

export function assertNumericValue(
  value: unknown,
  maxDigits = 100_000,
): asserts value is NumericValue {
  if (!isPlainRecord(value)) throw new TypeError('Invalid numeric value.');
  const cell = value as unknown as NumericValue;
  if (
    !validMagnitude(cell, maxDigits) ||
    !validPrecision(cell) ||
    !validDimension(cell) ||
    !validKind(cell) ||
    !validTemperature(cell)
  )
    throw new TypeError('Invalid numeric value or quantity metadata.');
  if (!validProvenance(cell)) throw new TypeError('Invalid numeric provenance.');
}

export function isPureTemperature(value: NumericValue): boolean {
  return value.dimension?.every((n, index) => n === (index === 4 ? 1 : 0)) === true;
}

/** Missing temperature tags in version-1 values mean absolute, never a guessed interval. */
export function normalizeNumericValue(value: NumericValue): NumericValue {
  assertNumericValue(value);
  return {
    ...value,
    ...(value.dimension ? { dimension: [...value.dimension] as NumericValue['dimension'] } : {}),
    ...(value.provenance ? { provenance: { ...value.provenance } } : {}),
    ...(isPureTemperature(value) ? { temperature: value.temperature ?? 'absolute' } : {}),
  };
}

export function mergeNumericProvenance(
  ...values: NumericValue[]
): Record<string, string> | undefined {
  const provenance: Record<string, string> = Object.create(null) as Record<string, string>;
  for (const value of values)
    for (const [key, revision] of Object.entries(value.provenance ?? {})) {
      if (Object.hasOwn(provenance, key) && provenance[key] !== revision)
        throw new CalculationError('#VALUE!');
      provenance[key] = revision;
    }
  return Object.keys(provenance).length ? { ...provenance } : undefined;
}
