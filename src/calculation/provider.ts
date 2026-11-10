import type { Workbook } from '../model/workbook.js';
import {
  calculateWorkbook,
  type CalculationOptions,
  type CalculationResult,
  type FormulaHost,
} from './evaluator.js';

export interface CalculationProvider {
  calculate(
    workbook: Workbook,
    options?: CalculationOptions,
    signal?: AbortSignal,
  ): Promise<CalculationResult>;
}

export function createClientCalculator(host: FormulaHost = {}): CalculationProvider {
  return {
    async calculate(workbook, options, signal) {
      signal?.throwIfAborted();
      const result = calculateWorkbook(workbook, options, host);
      signal?.throwIfAborted();
      return result;
    },
  };
}
