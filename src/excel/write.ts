import ExcelJS from 'exceljs';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { CalculationResult } from '../calculation/evaluator.js';
import { formatCellValue } from '../calculation/format.js';
import { assertCalculationResult } from '../calculation/http.js';
import { type CellValue, isError, isNumeric } from '../calculation/values.js';
import { parseCellStyle } from '../model/editing.js';
import { type FormulaNode, parseFormula } from '../model/formula-parser.js';
import { requireDefined } from '../model/require-defined.js';
import type { CellFormat } from '../model/review.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
} from '../model/spreadsheet-addressing.js';
import { assertWorkbook, type Workbook } from '../model/workbook.js';
import { hostFunctionMap, validateFormulaCapabilities } from './capabilities.js';
import { capturedText } from './captured-text.js';
import type {
  ExcelFormulaCapabilities,
  ExcelHostFunction,
  XlsxExport,
  XlsxReport,
} from './contracts.js';
import { excelValue, unsupportedCall } from './export-values.js';
import { inlineHostOperands } from './inline-operands.js';
import { exportCellStyle, exportNumberFormat, numberFormatLoss } from './styles.js';
import { createWarnings, type ReportedWarning } from './warnings.js';
import { writeManualWorkbook } from './write-archive.js';

type Sheet = Workbook['worksheets'][number];
type Warn = (warning: ReportedWarning) => void;

/** What every worksheet of one export shares. */
interface ExportContext {
  warn: Warn;
  report: XlsxReport;
  result: CalculationResult;
  functions: ReadonlyMap<string, ExcelHostFunction>;
}

/** One authored cell being exported, with its stored value and number format. */
interface ExportCell {
  sheet: Sheet;
  cell: ExcelJS.Cell;
  x: number;
  y: number;
  value: CellValue;
  format: CellFormat | undefined;
  location: string;
}

const unrepresentableError = (value: CellValue) =>
  isError(value) && ['#CYCLE!', '#LIMIT!', '#ERROR!'].includes(value.code);

/** Metadata XLSX cannot keep for a cell, reported before the cell is written. */
function metadataLosses(context: ExportContext, target: ExportCell) {
  const { format, value, location, sheet } = target;
  const formattingLoss = numberFormatLoss(format);
  if (formattingLoss)
    context.warn({
      code: 'number-format',
      message: formattingLoss,
      location,
      disposition: 'not-exported',
    });
  const address = formatSpreadsheetAddress(target.x, target.y);
  const canonical = sheet.inputValues?.[address];
  if (canonical)
    context.warn({
      code: 'canonical-input',
      message:
        'Canonical numeric metadata is projected as the stored value; XLSX does not retain the canonical override or its provenance.',
      location,
      disposition: 'not-exported',
    });
  if (isNumeric(value) && ('temperature' in value || value.quantityKind || 'provenance' in value))
    context.warn({
      code: 'quantity-metadata',
      message:
        'Quantity and provenance metadata are retained in displayed stored text where applicable, but XLSX does not retain the typed metadata.',
      location,
      disposition: 'not-exported',
    });
}

/** A frozen formula result: Excel keeps the stored value, or the formula with its cache. */
function writeParsedFormula(
  context: ExportContext,
  target: ExportCell,
  { formula, node }: { formula: string; node: FormulaNode },
) {
  const { cell, value, location } = target;
  const { functions } = context;
  // A direct rich host call has the same display projection in its
  // frozen cache and its expanded formula; Excel does not recalculate.
  const projected =
    node.kind === 'call' &&
    functions.get(node.name)?.capturedText === 'html-to-text' &&
    typeof value === 'string'
      ? capturedText(value, 'html-to-text', 32767)
      : value;
  const inlined = inlineHostOperands(node, {
    formula,
    sheetId: target.sheet.id,
    x: target.x,
    y: target.y,
    result: context.result,
    functions,
    loss: (code, message) => context.warn({ code, message, location }),
  });
  if (inlined === null || unrepresentableError(value)) {
    cell.value = excelValue(projected);
    if (unrepresentableError(value))
      context.warn({
        code: 'formula-error',
        message:
          'Excel cannot represent this calculation error; its stored error text is exported.',
        location,
      });
    return;
  }
  const cache = excelValue(projected),
    unsupported = unsupportedCall(node, functions);
  if (unsupported)
    context.warn({
      code: 'unsupported-formula',
      message: 'An unsupported or undeclared host formula is retained without a cache.',
      location,
      disposition: 'not-exported',
    });
  cell.value = {
    formula: inlined,
    result:
      cache === null || unsupported ? undefined : (cache as ExcelJS.CellFormulaValue['result']),
  };
}

function writeFormula(context: ExportContext, target: ExportCell, raw: string) {
  try {
    const formula = raw.trimStart().slice(1);
    writeParsedFormula(context, target, { formula, node: parseFormula(formula) });
  } catch {
    context.warn({
      code: 'malformed-formula',
      message: 'The formula is exported as authored text because it cannot be parsed.',
      location: target.location,
      disposition: 'not-exported',
    });
    target.cell.value = raw;
  }
}

/** Quantities become stored text, formulas keep their frozen results, other values export as-is. */
function writeContent(context: ExportContext, target: ExportCell, input: unknown) {
  const { cell, value, format, location } = target;
  const raw = typeof input === 'string' ? input : null;
  const formula = raw?.trimStart().startsWith('=') ? raw : null;
  if (formula) context.report.formulas++;
  if (isNumeric(value) && (format?.kind === 'unit' || value.dimension || value.quantityKind)) {
    cell.value =
      format?.kind === 'unit'
        ? formatWorksheetCell(value, format, { significantDigits: 10000 })
        : formatCellValue(value, { significantDigits: 10000 });
    context.warn({
      code: 'quantity-text',
      message:
        'A quantity is exported as stored text to preserve its units and temperature meaning.',
      location,
    });
  } else if (formula) writeFormula(context, target, formula);
  else cell.value = typeof value === 'string' ? value : excelValue(value);
}

function writeCell(
  context: ExportContext,
  { sheet, output, values }: { sheet: Sheet; output: ExcelJS.Worksheet; values: CellValue[][] },
  { input, x, y }: { input: unknown; x: number; y: number },
) {
  const address = formatSpreadsheetAddress(x, y),
    location = `${sheet.name}!${address}`;
  const target: ExportCell = {
    sheet,
    cell: output.getCell(y + 1, x + 1),
    x,
    y,
    value: values[y]?.[x] ?? null,
    format: sheet.formats?.[address],
    location,
  };
  context.report.cells++;
  metadataLosses(context, target);
  writeContent(context, target, input);
  const { cell } = target;
  if (typeof cell.value === 'string' && cell.value.length > 32767)
    throw new RangeError(`${location}: Excel cell strings have a 32767-character limit.`);
  cell.style = {
    ...exportCellStyle(parseCellStyle(sheet.style?.[address])),
    numFmt: exportNumberFormat(target.format),
  };
}

const hasCellValue = (input: unknown, result: CellValue | undefined) =>
  (input != null && input !== '') || result != null;

/** Write each populated cell and any additional independently authored formatting. */
function writeWorksheetCells(
  context: ExportContext,
  place: { sheet: Sheet; output: ExcelJS.Worksheet; values: CellValue[][] },
) {
  const { sheet, values } = place;
  sheet.data.forEach((row, y) =>
    row.forEach((input, x) => {
      if (hasCellValue(input, values[y][x])) writeCell(context, place, { input, x, y });
    }),
  );
  const formatted = new Set([
    ...Object.keys(sheet.style ?? {}),
    ...Object.keys(sheet.formats ?? {}),
  ]);
  for (const address of formatted) {
    const { x, y } = requireDefined(parseSpreadsheetAddress(address));
    const input = sheet.data[y]?.[x];
    if (!hasCellValue(input, values[y]?.[x])) writeCell(context, place, { input, x, y });
  }
}

function checkSheetName(name: string, names: Set<string>) {
  const key = name.toLowerCase();
  if (!name || name.length > 31 || /[[\]:*?/\\]/.test(name) || key === 'history' || names.has(key))
    throw new TypeError('Worksheet names must follow Excel naming rules.');
  names.add(key);
}

function writeLayout(context: ExportContext, sheet: Sheet, output: ExcelJS.Worksheet) {
  for (const [address, [columns, rows]] of Object.entries(sheet.mergeCells ?? {})) {
    context.report.mergedRanges++;
    const anchor = requireDefined(parseSpreadsheetAddress(address));
    output.mergeCells(anchor.y + 1, anchor.x + 1, anchor.y + rows, anchor.x + columns);
  }
  if (sheet.freezeColumns)
    output.views = [{ state: 'frozen', xSplit: sheet.freezeColumns, ySplit: 0 }];
}

function writeWorksheet(
  context: ExportContext,
  document: ExcelJS.Workbook,
  { sheet, values }: { sheet: Sheet; values: CellValue[][] },
) {
  const output = document.addWorksheet(sheet.name);
  sheet.columns?.forEach((column, x) => {
    output.getColumn(x + 1).width = Math.max(1, ((column.width ?? 100) - 5) / 7);
  });
  for (const [row, geometry] of Object.entries(sheet.rows ?? {}))
    output.getRow(Number(row) + 1).height = (geometry.height * 3) / 4;
  writeWorksheetCells(context, { sheet, output, values });
  writeLayout(context, sheet, output);
}

/** Captured-result export with explicit grouped losses; no live resolver is available. */
export async function exportXlsxWithReport(
  workbook: Workbook,
  result: CalculationResult,
  options: { formulaCapabilities?: ExcelFormulaCapabilities } = {},
): Promise<XlsxExport> {
  const functions = hostFunctionMap(validateFormulaCapabilities(options.formulaCapabilities));
  const warnings = createWarnings('stored-value');
  const report: XlsxReport = {
    worksheets: workbook.worksheets.length,
    cells: 0,
    formulas: 0,
    mergedRanges: 0,
    warnings: [],
  };
  assertWorkbook(workbook, { maxCells: 250000, maxWorksheets: 200 });
  assertCalculationResult(result, workbook);
  const document = new ExcelJS.Workbook(),
    names = new Set<string>();
  const context = { warn: warnings.warn, report, result, functions };
  workbook.worksheets.forEach((sheet, index) => {
    checkSheetName(sheet.name, names);
    writeWorksheet(context, document, { sheet, values: result.worksheets[index].values });
  });
  const bytes = await writeManualWorkbook(document);
  report.warnings = warnings.list();
  return { bytes, report };
}
