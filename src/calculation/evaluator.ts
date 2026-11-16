import { assertWorkbook, type Workbook, type WorkbookLimits } from '../model/workbook.js';
import { errorValueOf } from './cell-evaluation.js';
import { EXTENDED_FUNCTIONS, FUNCTION_ALIASES } from './extended-functions.js';
import type { NumericOptions } from './rational.js';
import { assertJsonValue, type CellValue, type JsonValue } from './values.js';
import { WorkbookEvaluation } from './workbook-evaluation.js';

export { formulaOperandKey } from './cell-evaluation.js';

export interface CalculationOptions extends NumericOptions, WorkbookLimits {
  timeZone?: string;
  maxOperations?: number;
  maxDepth?: number;
}
export interface CalculatedWorksheet {
  id: string;
  values: CellValue[][];
}
export interface ResolvedFormulaOperand {
  formula: string;
  name: string;
  value: CellValue;
  hostContext?: JsonValue;
}
export interface CalculationContext {
  operands: Record<string, ResolvedFormulaOperand>;
  hostContext?: JsonValue;
}
export interface CalculationDependencies {
  cells: Record<string, string[]>;
  extents?: Record<string, string[]>;
  /**
   * `'formulas'`: `cells` has an entry for every formula cell and none for other cells, which read
   * nothing. Without it, a cell with no entry has unknown dependencies.
   */
  scope?: 'formulas';
}
export interface FormulaCellContext {
  sheetId: string;
  x: number;
  y: number;
  formula: string;
}
export interface CalculationResult {
  schemaVersion: 1;
  worksheets: CalculatedWorksheet[];
  context?: CalculationContext;
  dependencies?: CalculationDependencies;
}
export interface FormulaHost {
  /** One captured clock value per pass for TODAY and NOW. */
  now?(): Date;
  hostContext?: JsonValue;
  readName?(name: string, context?: FormulaCellContext): CellValue | undefined;
  /** Host functions are synchronous and share this calculation's execution budget. */
  callFunction?(
    name: string,
    args: readonly CellValue[],
    context: FormulaCellContext,
  ): CellValue | undefined;
}

export const SUPPORTED_FUNCTIONS = Object.freeze([
  ...EXTENDED_FUNCTIONS,
  ...Object.keys(FUNCTION_ALIASES),
  'SWITCH',
  'ACOSH',
  'ASINH',
  'ATANH',
  'COSH',
  'SINH',
  'TANH',
  'EXPM1',
  'LOG1P',
  'IF',
  'IFERROR',
  'IFNA',
  'ISNA',
  'ISERR',
  'IFS',
  'CHOOSE',
  'DATE',
  'TIME',
  'TODAY',
  'NOW',
  'YEAR',
  'MONTH',
  'DAY',
  'HOUR',
  'MINUTE',
  'SECOND',
  'DAYS',
  'AND',
  'OR',
  'NOT',
  'TRUE',
  'FALSE',
  'SUM',
  'PRODUCT',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'ABS',
  'SIGN',
  'SQRT',
  'POWER',
  'MOD',
  'ROUND',
  'ROUNDUP',
  'ROUNDDOWN',
  'INT',
  'TRUNC',
  'CEILING',
  'FLOOR',
  'QUOTIENT',
  'PI',
  'SIN',
  'COS',
  'TAN',
  'ASIN',
  'ACOS',
  'ATAN',
  'EXP',
  'LN',
  'LOG',
  'LOG10',
  'ROW',
  'COLUMN',
  'ROWS',
  'COLUMNS',
  'LEN',
  'LEFT',
  'RIGHT',
  'MID',
  'LOWER',
  'UPPER',
  'TRIM',
  'CONCAT',
  'CONCATENATE',
  'VALUE',
  'ISNUMBER',
  'ISTEXT',
  'ISBLANK',
  'ISERROR',
  'NA',
  'SUMSQ',
  'VAR.S',
  'VAR.P',
  'STDEV.S',
  'STDEV.P',
]);

/** Pure, deterministic workbook evaluation; this module imports no browser or UI code. */
export function calculateWorkbook(
  workbook: Workbook,
  options: CalculationOptions = {},
  host: FormulaHost = {},
): CalculationResult {
  assertWorkbook(workbook, options);
  const evaluation = new WorkbookEvaluation(workbook, options, host);
  if (host.hostContext !== undefined) assertJsonValue(host.hostContext);
  if (
    options.wordSizeBits !== undefined &&
    (typeof options.wordSizeBits !== 'string' ||
      !/^[1-9]\d{0,5}$/.test(options.wordSizeBits) ||
      BigInt(options.wordSizeBits) > 100000n)
  )
    throw new RangeError('Invalid wordSizeBits context.');
  const result: CalculationResult = {
    schemaVersion: 1,
    worksheets: workbook.worksheets.map((sheet, index) => ({
      id: sheet.id,
      values: sheet.data.map((row, y) =>
        row.map((_cell, x) => {
          try {
            return evaluation.ensureCell({ sheet: index, x, y });
          } catch (error) {
            return errorValueOf(error);
          }
        }),
      ),
    })),
  };
  result.dependencies = evaluation.dependencies();
  const { operands } = evaluation;
  if (Object.keys(operands).length || host.hostContext !== undefined)
    result.context = {
      operands,
      ...(host.hostContext === undefined ? {} : { hostContext: structuredClone(host.hostContext) }),
    };
  return result;
}
