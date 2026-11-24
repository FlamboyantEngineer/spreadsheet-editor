import type ExcelJS from 'exceljs';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';
import { formatCellValue } from '../calculation/format.js';
import { createArithmetic } from '../calculation/rational.js';
import { type CellValue, isError, isNumeric, type NumericValue } from '../calculation/values.js';
import type { FormulaNode } from '../model/formula-parser.js';
import type { ExcelHostFunction } from './contracts.js';

/** The error codes Excel has; other calculation errors are exported as their text. */
export const EXCEL_ERRORS = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];

const exact = createArithmetic({ maxIntegerDigits: 100000 });
const excelMinMagnitude = 2.22507385850721e-308;
const excelMaxStoredMagnitude = 9.99999999999999e307;

function exactExcelText(value: NumericValue): string {
  if (value.denominator === '1') return value.numerator;
  let denominator = BigInt(value.denominator);
  while (denominator % 2n === 0n) denominator /= 2n;
  while (denominator % 5n === 0n) denominator /= 5n;
  if (denominator === 1n) {
    const rendered = formatCellValue(value, {
      significantDigits: 10000,
      approximationMarker: false,
    });
    if (exact.compare(exact.parse(rendered), value) === 0) return rendered;
  }
  return `${value.numerator}/${value.denominator}`;
}

/** Native Excel numbers must round-trip exactly at Excel's 15-digit precision. */
function excelNumber(value: NumericValue, formulaCache: boolean) {
  const rounded = formatCellValue(value, { significantDigits: 15, approximationMarker: false });
  const native = Number(rounded),
    magnitude = Math.abs(native);
  const maximum = formulaCache ? Number.MAX_VALUE : excelMaxStoredMagnitude;
  if (
    (magnitude === 0 || magnitude >= excelMinMagnitude) &&
    magnitude <= maximum &&
    exact.compare(exact.parse(String(native)), value) === 0
  )
    return native;
  return exactExcelText(value);
}

export function excelValue(
  value: CellValue,
  formulaCache = false,
): string | number | boolean | ExcelJS.CellErrorValue | null {
  if (isError(value))
    return EXCEL_ERRORS.includes(value.code)
      ? { error: value.code as ExcelJS.CellErrorValue['error'] }
      : value.code;
  if (!isNumeric(value)) return value;
  if (value.dimension) return formatCellValue(value, { significantDigits: 10000 });
  return excelNumber(value, formulaCache);
}

export function numericTextWarning(value: CellValue, written: ReturnType<typeof excelValue>) {
  if (isNumeric(value) && typeof written === 'string')
    return {
      code: 'numeric-text',
      message: 'An exact value is stored as text because an Excel number would lose data.',
      disposition: 'stored-value' as const,
    };
  return undefined;
}

/** A captured operand as Excel formula text. */
export function operandText(value: CellValue): string {
  if (isNumeric(value))
    return value.denominator === '1'
      ? value.numerator
      : `(${value.numerator}/${value.denominator})`;
  if (typeof value === 'string') return '"' + value.replaceAll('"', '""') + '"';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return isError(value) ? value.code : '""';
}

/** Whether a formula calls a function that neither the catalog nor a declared host function has. */
export function unsupportedCall(
  node: FormulaNode,
  functions: ReadonlyMap<string, ExcelHostFunction>,
): boolean {
  if (node.kind === 'call')
    return (
      (!SUPPORTED_FUNCTIONS.includes(node.name) && !functions.has(node.name)) ||
      (!functions.has(node.name) && node.args.some((arg) => unsupportedCall(arg, functions)))
    );
  if (node.kind === 'binary')
    return unsupportedCall(node.left, functions) || unsupportedCall(node.right, functions);
  return node.kind === 'unary' ? unsupportedCall(node.operand, functions) : false;
}
