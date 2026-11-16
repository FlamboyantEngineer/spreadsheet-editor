import type ExcelJS from 'exceljs';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';
import { formatCellValue } from '../calculation/format.js';
import type { Arithmetic } from '../calculation/rational.js';
import { serializeCellStyle } from '../model/editing.js';
import { type FormulaNode, parseFormula } from '../model/formula-parser.js';
import { requireDefined } from '../model/require-defined.js';
import type { CellFormat } from '../model/review.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import type { XlsxReport } from './contracts.js';
import { importCellStyle, importNumberFormat } from './styles.js';
import type { ReportedWarning } from './warnings.js';

/** What every part of one import shares: its warnings, counts and conversion settings. */
export interface ImportContext {
  warn(warning: ReportedWarning): void;
  report: XlsxReport;
  theme: readonly string[];
  arithmetic: Arithmetic;
  /** Whether the workbook counts dates from 1904, so date serials shift by 1462 days. */
  date1904: boolean;
  external: ReadonlyMap<string, unknown>;
}

/** The worksheet being filled, and the plain numbers its XML authored. */
export interface SheetImport {
  sheet: Worksheet;
  name: string;
  authoredNumbers: ReadonlyMap<string, string>;
}

/** Function names a formula calls that neither the catalog nor the host provides. */
function unsupported(node: FormulaNode, external: ReadonlyMap<string, unknown>): string[] {
  if (node.kind === 'call')
    return [
      ...(!SUPPORTED_FUNCTIONS.includes(node.name) && !external.has(node.name) ? [node.name] : []),
      ...node.args.flatMap((arg) => unsupported(arg, external)),
    ];
  if (node.kind === 'binary')
    return [...unsupported(node.left, external), ...unsupported(node.right, external)];
  return node.kind === 'unary' ? unsupported(node.operand, external) : [];
}

function formulaValue(context: ImportContext, formula: string, location: string): CellInput {
  context.report.formulas++;
  try {
    for (const name of unsupported(parseFormula(formula), context.external))
      context.warn({
        code: `formula:${name}`,
        message: `${name} is retained but is unavailable in the calculation catalog.`,
        location,
      });
  } catch {
    context.warn({
      code: 'malformed-formula',
      message: 'Formula is retained as authored but cannot be parsed.',
      location,
    });
  }
  return '=' + formula;
}

function authoredNumber(context: ImportContext, number: string, format: CellFormat | null) {
  if (!context.date1904 || !format || !['date', 'time', 'datetime'].includes(format.kind))
    return number;
  const { arithmetic } = context;
  return formatCellValue(arithmetic.add(arithmetic.parse(number), arithmetic.parse('1462')), {
    significantDigits: 10000,
    notation: 'decimal',
  });
}

type ObjectValue = Exclude<ExcelJS.CellValue, null | undefined | string | number | boolean | Date>;

function objectValue(
  context: ImportContext,
  value: ObjectValue,
  { cell, location }: { cell: ExcelJS.Cell; location: string },
): CellInput {
  if ('richText' in value) {
    context.warn({ code: 'rich-text', message: 'Rich text becomes plain text.', location });
    return value.richText.map((part) => part.text).join('');
  }
  if ('hyperlink' in value) {
    context.warn({ code: 'hyperlink', message: 'Hyperlinks become plain text.', location });
    return value.text;
  }
  return 'error' in value ? value.error : cell.text;
}

interface CellPlace {
  address: string;
  location: string;
  format: CellFormat | null;
}

function cellValue(
  context: ImportContext,
  {
    cell,
    authoredNumbers,
  }: { cell: ExcelJS.Cell; authoredNumbers: SheetImport['authoredNumbers'] },
  { address, location, format }: CellPlace,
): CellInput {
  if (cell.formula) return formulaValue(context, cell.formula, location);
  const number = authoredNumbers.get(address);
  if (number !== undefined) return authoredNumber(context, number, format);
  const { value } = cell;
  // Numeric date serials normally use the raw XML path above.
  if (value instanceof Date) return String((value.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
  if (typeof value === 'object' && value !== null)
    return objectValue(context, value, { cell, location });
  return value as CellInput;
}

/** Errors and booleans center, authored numbers align right and everything else left. */
function defaultAlignment(value: CellInput, authored: boolean): string {
  if (typeof value === 'boolean' || (typeof value === 'string' && value.startsWith('#')))
    return 'center';
  return authored ? 'right' : 'left';
}

function cellWarnings(context: ImportContext, cell: ExcelJS.Cell, place: CellPlace) {
  const { location } = place;
  const warn = (code: string, message: string) => context.warn({ code, message, location });
  if (cell.note) warn('note', 'Excel notes and comments are not imported.');
  if (
    cell.fill?.type === 'pattern' &&
    cell.fill.pattern !== 'solid' &&
    cell.fill.pattern !== 'none'
  )
    warn('pattern-fill', 'Pattern fills are not imported.');
  if (cell.alignment?.textRotation) warn('rotated-text', 'Text rotation is not imported.');
}

function legendWarning(context: ImportContext, value: CellInput, location: string) {
  if (typeof value === 'string' && /\b(?:ITAR|EAR99|ECCN|EXPORT CONTROLLED)\b/i.test(value))
    context.warn({
      code: 'export-control-legend',
      message: 'Export-control text is retained; set the document marking in the host application.',
      location,
    });
}

/** Imports one cell's value, number format and style into the worksheet. */
export function importCell(
  context: ImportContext,
  target: SheetImport,
  { cell, x, y }: { cell: ExcelJS.Cell; x: number; y: number },
) {
  if (cell.isMerged && cell.master.address !== cell.address) return;
  const { sheet } = target;
  const address = formatSpreadsheetAddress(x - 1, y - 1),
    location = `${target.name}!${address}`;
  const format = importNumberFormat(cell.numFmt);
  if (!format)
    context.warn({
      code: `number-format:${cell.numFmt}`,
      message: 'Unsupported number format falls back to General.',
      location,
    });
  if (format && format.kind !== 'general') requireDefined(sheet.formats)[address] = format;
  const place = { address, location, format };
  const value = cellValue(context, { cell, authoredNumbers: target.authoredNumbers }, place);
  const values = sheet.data[y - 1];
  while (values.length < x) values.push(null);
  values[x - 1] = value;
  const style = importCellStyle(cell.style, context.theme);
  if (!cell.alignment?.horizontal)
    style['text-align'] = defaultAlignment(value, target.authoredNumbers.has(address));
  requireDefined(sheet.style)[address] = serializeCellStyle(style);
  requireDefined(sheet.cellProvenance)[address] = {
    actor: { id: 'system', name: 'System' },
    source: 'system',
  };
  context.report.cells++;
  cellWarnings(context, cell, place);
  legendWarning(context, value, location);
}
