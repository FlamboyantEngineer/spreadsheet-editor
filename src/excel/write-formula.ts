import type ExcelJS from 'exceljs';
import type { CalculationResult } from '../calculation/evaluator.js';
import { type CellValue, isError } from '../calculation/values.js';
import { type FormulaNode, FormulaSyntaxError, parseFormula } from '../model/formula-parser.js';
import { capturedText, excelTextProblem } from './captured-text.js';
import type { ExcelHostFunction } from './contracts.js';
import { EXCEL_ERRORS, excelValue, numericTextWarning, unsupportedCall } from './export-values.js';
import { inlineHostOperands } from './inline-operands.js';
import type { ReportedWarning } from './warnings.js';

type FormulaWarning = Omit<ReportedWarning, 'location'>;
type Warn = (warning: FormulaWarning) => void;

interface FormulaCell {
  sheetId: string;
  x: number;
  y: number;
  value: CellValue;
}

interface FormulaExport {
  result: CalculationResult;
  functions: ReadonlyMap<string, ExcelHostFunction>;
  warn: Warn;
}

/** Direct rich results use the same projection as their inlined operand. */
function projectFormulaResult(
  node: FormulaNode,
  value: CellValue,
  { functions, warn }: FormulaExport,
) {
  if (
    typeof value !== 'string' ||
    node.kind !== 'call' ||
    functions.get(node.name)?.capturedText !== 'html-to-text'
  )
    return value;
  try {
    return capturedText(value, 'html-to-text', 32767);
  } catch (error) {
    if (!(error instanceof RangeError)) throw error;
    warn({
      code: 'rich-formula-value',
      message:
        'The rich formula result cannot be represented as Excel cell text. The authored formula is retained as text.',
      disposition: 'not-exported',
    });
    return undefined;
  }
}

/** A disconnected formula needs a stored result. Retain the authored text if that result cannot fit. */
function storedFormulaValue(
  { value, raw }: { value: CellValue; raw: string },
  losses: readonly FormulaWarning[],
  warn: Warn,
) {
  const stored = excelValue(value);
  if (typeof stored === 'string' && excelTextProblem(stored)) {
    for (const loss of losses) warn({ ...loss, disposition: 'not-exported' });
    warn({
      code: 'formula-stored-result',
      message:
        'The formula and its stored result cannot be exported. The authored formula is retained as text.',
      disposition: 'not-exported',
    });
    return raw;
  }
  losses.forEach(warn);
  const numericText = numericTextWarning(value, stored);
  if (numericText) warn(numericText);
  return stored;
}

function formulaWithCache(formula: string, value: CellValue, warn: Warn): ExcelJS.CellFormulaValue {
  const cache = excelValue(value, true);
  if (typeof cache === 'string' && excelTextProblem(cache)) {
    warn({
      code: 'formula-cache',
      message:
        'Excel cannot represent the whole cached result as text. The formula is retained without a cache.',
      disposition: 'not-exported',
    });
    return { formula };
  }
  const numericText = numericTextWarning(value, cache);
  if (numericText) warn(numericText);
  return { formula, result: cache ?? undefined };
}

function parseExportFormula(formula: string, warn: Warn) {
  try {
    return parseFormula(formula);
  } catch (error) {
    if (!(error instanceof FormulaSyntaxError)) throw error;
    const limited = error.token === '#LIMIT!';
    warn({
      code: limited ? 'formula-parser-limit' : 'malformed-formula',
      message: limited
        ? 'The formula exceeds the parser limits. The authored formula is retained as text.'
        : 'The formula cannot be parsed. The authored formula is retained as text.',
      disposition: 'not-exported',
    });
    return undefined;
  }
}

/** Writes an Excel formula and its captured result, or a reported fallback, without calculation. */
export function exportFormula(
  raw: string,
  target: FormulaCell,
  options: FormulaExport,
): ExcelJS.CellValue {
  const { warn, result, functions } = options;
  const formula = raw.trimStart().slice(1);
  const node = parseExportFormula(formula, warn);
  if (!node) return raw;
  const value = projectFormulaResult(node, target.value, options);
  if (value === undefined) return raw;

  const losses: FormulaWarning[] = [];
  const inlined = inlineHostOperands(node, {
    formula,
    sheetId: target.sheetId,
    x: target.x,
    y: target.y,
    result,
    functions,
    loss: (code, message) => losses.push({ code, message }),
  });
  if (isError(value) && !EXCEL_ERRORS.includes(value.code))
    losses.push({
      code: 'formula-error',
      message: 'Excel cannot represent this calculation error as an error value.',
    });
  if (inlined !== null && excelTextProblem(inlined))
    losses.push({
      code: 'formula-text',
      message: 'The formula contains text Excel cannot represent.',
    });
  if (inlined === null || losses.length) return storedFormulaValue({ value, raw }, losses, warn);
  if (unsupportedCall(node, functions)) {
    warn({
      code: 'unsupported-formula',
      message: 'An unsupported or undeclared host formula is retained without a cache.',
      disposition: 'not-exported',
    });
    return { formula: inlined };
  }
  return formulaWithCache(inlined, value, warn);
}
