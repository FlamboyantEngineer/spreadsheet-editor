import type { NumericOptions } from './rational.js';
import { createQuantityArithmetic } from './units.js';
import { CalculationError, type CellValue, isNumeric, type NumericValue } from './values.js';

/** Selection statistics use the same physical-kind and exact-value rules as formulas. */
export function summarizeCellValues(
  values: readonly CellValue[],
  options: NumericOptions = {},
): { count: number; sum: NumericValue | null; average: NumericValue | null } {
  const numeric = values.filter(isNumeric),
    arithmetic = createQuantityArithmetic(options);
  const available = (calculate: () => NumericValue) => {
    try {
      return calculate();
    } catch (error) {
      if (!(error instanceof CalculationError)) throw error;
      return null;
    }
  };
  return {
    count: values.filter((value) => value !== null && value !== '').length,
    sum: available(() => numeric.reduce(arithmetic.add, arithmetic.parse('0'))),
    average: numeric.length ? available(() => arithmetic.average(numeric)) : null,
  };
}
