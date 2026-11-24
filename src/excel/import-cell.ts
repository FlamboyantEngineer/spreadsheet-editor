import type ExcelJS from 'exceljs';
import { isFormula, isNumericInputText } from '../calculation/cell-evaluation.js';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';
import { formatCellValue } from '../calculation/format.js';
import type { Arithmetic } from '../calculation/rational.js';
import { serializeCellStyle } from '../model/editing.js';
import { type FormulaNode, parseFormula } from '../model/formula-parser.js';
import { DEFAULT_WORKBOOK_LIMITS } from '../model/limits.js';
import { requireDefined } from '../model/require-defined.js';
import type { CellFormat } from '../model/review.js';
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
  if (!context.date1904 || !format || !['date', 'datetime'].includes(format.kind)) return number;
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
): string {
  if ('richText' in value) {
    context.warn({ code: 'rich-text', message: 'Rich text becomes plain text.', location });
    return value.richText.map((part) => part.text).join('');
  }
  if ('hyperlink' in value) {
    context.warn({ code: 'hyperlink', message: 'Hyperlinks become plain text.', location });
    return value.text;
  }
  if ('error' in value) {
    context.warn({
      code: 'static-error',
      message:
        'A stored Excel error becomes text because the workbook stores only authored inputs.',
      location,
    });
    return value.error;
  }
  return cell.text;
}

interface CellPlace {
  location: string;
  format: CellFormat | null;
}

/** Preserve Excel text when its spelling would be interpreted as a different authored type. */
function literalText(value: string, format: CellFormat | null) {
  return value.startsWith("'") ||
    isFormula(value) ||
    isNumericInputText(value) ||
    format?.kind === 'unit'
    ? "'" + value
    : value;
}

function cellValue(
  context: ImportContext,
  { cell, number }: { cell: ExcelJS.Cell; number: string | undefined },
  { location, format }: CellPlace,
): CellInput {
  if (cell.formula) return formulaValue(context, cell.formula, location);
  if (number !== undefined) {
    if (!isNumericInputText(number)) throw new TypeError(`${location}: invalid numeric cell.`);
    return authoredNumber(context, number, format);
  }
  const { value } = cell;
  if (typeof value === 'number' || value instanceof Date)
    throw new TypeError(`${location}: numeric cell has no exact XML value.`);
  if (typeof value === 'object' && value !== null)
    return literalText(objectValue(context, value, { cell, location }), format);
  return typeof value === 'string' ? literalText(value, format) : (value ?? null);
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

/** A value, unlike independent cell formatting, expands the authored rows and has a contributor. */
function storeImportedValue(
  context: ImportContext,
  sheet: Worksheet,
  {
    value,
    x,
    y,
    address,
    location,
  }: { value: Exclude<CellInput, null>; x: number; y: number; address: string; location: string },
) {
  if (typeof value === 'string' && value.length > DEFAULT_WORKBOOK_LIMITS.maxCellLength)
    throw new RangeError(`${location}: cell exceeds the workbook text limit.`);
  while (sheet.data.length <= y) sheet.data.push([]);
  const values = sheet.data[y];
  while (values.length <= x) values.push(null);
  values[x] = value;
  requireDefined(sheet.cellProvenance)[address] = {
    actor: { id: 'system', name: 'System' },
    source: 'system',
  };
  context.report.cells++;
  legendWarning(context, value, location);
}

/** Imports one cell's value, number format and style into the worksheet. */
export function importCell(
  context: ImportContext,
  sheet: Worksheet,
  {
    cell,
    address,
    x,
    y,
    number,
  }: { cell: ExcelJS.Cell; address: string; x: number; y: number; number: string | undefined },
) {
  if (cell.isMerged && cell.master.address !== cell.address) return;
  const location = `${sheet.name}!${address}`;
  const { format, lossy } = importNumberFormat(cell.numFmt);
  if (!format || lossy)
    context.warn({
      code: `number-format:${cell.numFmt}`,
      message: format
        ? 'Excel number format modifiers are omitted from its editor format.'
        : 'Unsupported number format falls back to General.',
      location,
    });
  if (format && format.kind !== 'general') requireDefined(sheet.formats)[address] = format;
  const place = { location, format };
  const value = cellValue(context, { cell, number }, place);
  const style = importCellStyle(cell.style, context.theme);
  if (value !== null && !cell.alignment?.horizontal)
    style['text-align'] = defaultAlignment(value, number !== undefined);
  if (Object.keys(style).length) requireDefined(sheet.style)[address] = serializeCellStyle(style);
  cellWarnings(context, cell, place);
  if (value !== null) storeImportedValue(context, sheet, { value, x, y, address, location });
}
