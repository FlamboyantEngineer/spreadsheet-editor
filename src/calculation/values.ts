export interface NumericValue {
  readonly dimension?: import('./units.js').Dimension;
  readonly quantityKind?: string;
  readonly kind: 'number';
  readonly numerator: string;
  readonly denominator: string;
  readonly approximate: boolean;
  /** Working significant digits for an approximation; absent for exact values. */
  readonly precision?: number;
}

export type FormulaErrorCode =
  '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#ERROR!' | '#LIMIT!' | '#CYCLE!';
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
