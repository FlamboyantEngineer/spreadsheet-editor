import type { Workbook } from '../model/workbook.js';
import type {
  CalculationContext,
  CalculationDependencies,
  CalculationResult,
} from './evaluator.js';
import {
  assertJsonValue,
  assertNumericValue,
  type CellValue,
  FORMULA_ERROR_CODES,
} from './values.js';

const isRecord = (value: unknown): value is Record<string, unknown> =>
  Boolean(value) && typeof value === 'object' && !Array.isArray(value);

function validateCell(value: CellValue): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (!value || typeof value !== 'object') throw new TypeError('Invalid calculated cell.');
  if (value.kind === 'error' && (FORMULA_ERROR_CODES as readonly string[]).includes(value.code))
    return;
  assertNumericValue(value);
}

function assertWorksheets(result: CalculationResult, workbook: Workbook): void {
  if (
    result.schemaVersion !== 1 ||
    !Array.isArray(result.worksheets) ||
    result.worksheets.length !== workbook.worksheets.length
  )
    throw new TypeError('Invalid result worksheets.');
  workbook.worksheets.forEach((sheet, index) => {
    const output = result.worksheets[index];
    if (
      output?.id !== sheet.id ||
      !Array.isArray(output.values) ||
      output.values.length !== sheet.data.length
    )
      throw new TypeError('Result does not match the workbook.');
    output.values.forEach((row, y) => {
      if (!Array.isArray(row) || row.length !== sheet.data[y].length)
        throw new TypeError('Result does not match the row.');
      row.forEach(validateCell);
    });
  });
}

function assertOperand(operand: CalculationContext['operands'][string]): void {
  if (
    !operand ||
    typeof operand !== 'object' ||
    typeof operand.formula !== 'string' ||
    operand.formula.length > 200000 ||
    typeof operand.name !== 'string' ||
    operand.name.length > 256
  )
    throw new TypeError('Invalid captured operand.');
  validateCell(operand.value);
  if (operand.hostContext !== undefined) assertJsonValue(operand.hostContext);
}

function assertContext(context: CalculationContext): void {
  if (
    !isRecord(context) ||
    !isRecord(context.operands) ||
    Object.keys(context.operands).length > 100000
  )
    throw new TypeError('Invalid calculation context.');
  if (context.hostContext !== undefined) assertJsonValue(context.hostContext);
  Object.values(context.operands).forEach(assertOperand);
}

/** Checks a `[sheetId, x, y]` dependency key against the workbook's sheets and the grid. */
function cellKeyCheck(sheets: ReadonlySet<string>) {
  return (key: unknown) => {
    let tuple: unknown;
    try {
      tuple = typeof key === 'string' ? JSON.parse(key) : undefined;
    } catch {
      throw new TypeError('Invalid dependency cell key.');
    }
    const valid =
      Array.isArray(tuple) &&
      tuple.length === 3 &&
      typeof tuple[0] === 'string' &&
      sheets.has(tuple[0]) &&
      inRange(tuple[1], 16383) &&
      inRange(tuple[2], 1048575);
    if (!valid) throw new TypeError('Invalid dependency cell key.');
  };
}
const inRange = (value: unknown, max: number) =>
  Number.isSafeInteger(value) && (value as number) >= 0 && (value as number) <= max;

function assertExtents(
  extents: CalculationDependencies['extents'],
  { sheets, cellKey }: { sheets: ReadonlySet<string>; cellKey: (key: unknown) => void },
): void {
  if (!isRecord(extents)) throw new TypeError('Invalid extent dependencies.');
  for (const [key, list] of Object.entries(extents)) {
    cellKey(key);
    if (
      !Array.isArray(list) ||
      list.length > 200 ||
      list.some((sheet) => typeof sheet !== 'string' || !sheets.has(sheet))
    )
      throw new TypeError('Invalid extent dependency.');
  }
}

function assertDependencies(dependencies: CalculationDependencies, workbook: Workbook): void {
  if (
    !isRecord(dependencies) ||
    !isRecord(dependencies.cells) ||
    (dependencies.scope !== undefined && dependencies.scope !== 'formulas')
  )
    throw new TypeError('Invalid calculation dependencies.');
  const sheets = new Set(workbook.worksheets.map((sheet) => sheet.id));
  const cellKey = cellKeyCheck(sheets);
  if (Object.keys(dependencies.cells).length > 250000)
    throw new RangeError('Too many calculation dependencies.');
  for (const [key, reads] of Object.entries(dependencies.cells)) {
    cellKey(key);
    if (!Array.isArray(reads) || reads.length > 10000)
      throw new TypeError('Invalid dependency reads.');
    reads.forEach(cellKey);
  }
  if (dependencies.extents !== undefined) assertExtents(dependencies.extents, { sheets, cellKey });
}

/** Validates an untrusted calculation result against the workbook it claims to calculate. */
export function assertCalculationResult(
  value: unknown,
  workbook: Workbook,
): asserts value is CalculationResult {
  if (!value || typeof value !== 'object') throw new TypeError('Invalid calculation response.');
  const result = value as CalculationResult;
  assertWorksheets(result, workbook);
  if (result.context !== undefined) assertContext(result.context);
  if (result.dependencies !== undefined) assertDependencies(result.dependencies, workbook);
}
