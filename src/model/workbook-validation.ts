import { assertNumericValue } from '../calculation/values.js';
import { assertCellFormat } from './cell-format.js';
import { worksheetExtent } from './extent.js';
import { DEFAULT_WORKBOOK_LIMITS } from './limits.js';
import { requireDefined } from './require-defined.js';
import { checkReviewData } from './review-validation.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { Workbook, WorkbookLimits, Worksheet } from './workbook.js';
import { worksheetNameKey } from './worksheet-name.js';

type Bounds = Required<WorkbookLimits>;

/** What validating one workbook has seen so far across its worksheets. */
interface Seen {
  ids: Set<string>;
  names: Set<string>;
  cells: number;
  rows: number;
}

const CELL_ADDRESS = /^[A-Z]+[1-9]\d*$/;
const isMap = (value: unknown) =>
  Boolean(value) && typeof value === 'object' && !Array.isArray(value);
/** Whether an optional field is present but not a plain map. */
const badMap = (value: unknown) => value !== undefined && !isMap(value);

function checkIdentity(sheet: Worksheet, seen: Seen) {
  if (
    !sheet ||
    typeof sheet.id !== 'string' ||
    !sheet.id ||
    sheet.id.length > 256 ||
    typeof sheet.name !== 'string' ||
    !sheet.name.trim() ||
    sheet.name.length > 256 ||
    seen.ids.has(sheet.id) ||
    seen.names.has(worksheetNameKey(sheet.name)) ||
    !Array.isArray(sheet.data)
  )
    throw new TypeError('Worksheets need unique ids, unique names, and row data.');
}

function checkLines(sheet: Worksheet, bounds: Bounds) {
  const dimensions = sheet.minDimensions;
  if (
    dimensions !== undefined &&
    (!Array.isArray(dimensions) ||
      dimensions.length !== 2 ||
      dimensions[0] * dimensions[1] > bounds.maxCells ||
      dimensions.some((n, i) => !Number.isInteger(n) || n < 1 || n > (i === 0 ? 16384 : 1048576)))
  )
    throw new TypeError('Invalid worksheet dimensions.');
  for (const axis of ['rowIds', 'columnIds'] as const) {
    const list = sheet[axis];
    if (
      list !== undefined &&
      (!Array.isArray(list) ||
        list.some((id) => typeof id !== 'string' || !id || id.length > 256) ||
        new Set(list).size !== list.length)
    )
      throw new TypeError('Invalid worksheet line identities.');
  }
}

const badColumn = (column: NonNullable<Worksheet['columns']>[number]) =>
  !column ||
  (column.width !== undefined &&
    (typeof column.width !== 'number' || !Number.isFinite(column.width) || column.width < 1)) ||
  (column.title !== undefined && typeof column.title !== 'string');

function checkMerges(sheet: Worksheet, bounds: Bounds) {
  if (badMap(sheet.mergeCells)) throw new TypeError('Invalid merged cells.');
  for (const [address, span] of Object.entries(sheet.mergeCells ?? {}))
    if (
      !CELL_ADDRESS.test(address) ||
      !Array.isArray(span) ||
      span.length !== 2 ||
      span[0] * span[1] > bounds.maxCells ||
      span.some((value) => !Number.isSafeInteger(value) || value < 1 || value > bounds.maxCells)
    )
      throw new TypeError('Invalid merged cells.');
}

function checkGeometry(sheet: Worksheet, bounds: Bounds) {
  const { columns } = sheet;
  if (
    columns !== undefined &&
    (!Array.isArray(columns) || columns.length > 16_384 || columns.some(badColumn))
  )
    throw new TypeError('Invalid worksheet columns.');
  checkMerges(sheet, bounds);
  if (badMap(sheet.rows)) throw new TypeError('Invalid worksheet rows.');
  for (const [index, row] of Object.entries(sheet.rows ?? {}))
    if (!/^\d+$/.test(index) || !row || !Number.isFinite(row.height) || row.height < 1)
      throw new TypeError('Invalid row height.');
}

function checkPresentation(sheet: Worksheet) {
  if (
    badMap(sheet.style) ||
    Object.values(sheet.style ?? {}).some((value) => typeof value !== 'string')
  )
    throw new TypeError('Invalid cell styles.');
  const frozen = sheet.freezeColumns;
  if (frozen !== undefined && (!Number.isSafeInteger(frozen) || frozen < 0 || frozen > 16_384))
    throw new TypeError('Invalid frozen column count.');
  if (badMap(sheet.formats)) throw new TypeError('Invalid cell formats.');
  for (const [address, format] of Object.entries(sheet.formats ?? {})) {
    if (!CELL_ADDRESS.test(address)) throw new TypeError('Invalid cell format address.');
    assertCellFormat(format);
  }
}

const badProvenance = ([address, entry]: [
  string,
  NonNullable<Worksheet['cellProvenance']>[string],
]) =>
  !CELL_ADDRESS.test(address) ||
  !entry ||
  !entry.actor ||
  typeof entry.actor.id !== 'string' ||
  !entry.actor.id ||
  entry.actor.id.length > 256 ||
  typeof entry.actor.name !== 'string' ||
  entry.actor.name.length > 512 ||
  !['human', 'ai', 'system'].includes(entry.source) ||
  (entry.at !== undefined &&
    (typeof entry.at !== 'string' || !Number.isFinite(Date.parse(entry.at))));

function checkInputs(sheet: Worksheet, bounds: Bounds) {
  if (
    badMap(sheet.cellProvenance) ||
    Object.entries(sheet.cellProvenance ?? {}).some(badProvenance)
  )
    throw new TypeError('Invalid cell provenance.');
  if (badMap(sheet.inputValues)) throw new TypeError('Invalid canonical input values.');
  for (const [address, canonical] of Object.entries(sheet.inputValues ?? {})) {
    const position = CELL_ADDRESS.test(address) ? parseSpreadsheetAddress(address) : null;
    const input = position && sheet.data[position.y]?.[position.x];
    if (
      !position ||
      input === undefined ||
      (typeof input === 'string' && input.trimStart().startsWith('='))
    )
      throw new TypeError('Canonical input requires a corresponding nonformula data entry.');
    assertNumericValue(canonical, bounds.maxCellLength);
  }
}

const metadataMaps = (sheet: Worksheet) => [
  sheet.style,
  sheet.formats,
  sheet.cellProvenance,
  sheet.inputValues,
  sheet.mergeCells,
];

function checkExtent(sheet: Worksheet, workbook: Workbook, bounds: Bounds) {
  const addresses = metadataMaps(sheet).flatMap((map) => Object.keys(map ?? {}));
  for (const address of addresses) {
    const position = CELL_ADDRESS.test(address) ? parseSpreadsheetAddress(address) : null;
    if (!position || position.x > 16383 || position.y > 1048575)
      throw new TypeError('Invalid cell metadata address.');
  }
  const extent = worksheetExtent(sheet, workbook);
  if (
    extent.columns > 16384 ||
    extent.rows > 1048576 ||
    extent.columns * extent.rows > bounds.maxCells
  )
    throw new RangeError('Worksheet extent exceeds the cell limit.');
  if (
    (sheet.rowIds && sheet.rowIds.length !== extent.rows) ||
    (sheet.columnIds && sheet.columnIds.length !== extent.columns)
  )
    throw new TypeError('Worksheet line identities must cover the effective extent.');
}

/** Cells that only metadata reaches, past the end of their data row, each counted once. */
function metadataCells(sheet: Worksheet): number {
  const beyondData = (address: string) => {
    const position = requireDefined(parseSpreadsheetAddress(address));
    return position.x >= (sheet.data[position.y]?.length ?? 0);
  };
  const counted = new Set(Object.keys(sheet.style ?? {}));
  let cells = [...counted].filter(beyondData).length;
  for (const map of [sheet.formats, sheet.cellProvenance, sheet.inputValues])
    for (const address of Object.keys(map ?? {})) {
      if (!counted.has(address) && beyondData(address)) cells++;
      counted.add(address);
    }
  return cells;
}

const badCell = (cell: unknown) =>
  cell !== null &&
  typeof cell !== 'string' &&
  typeof cell !== 'boolean' &&
  !(typeof cell === 'number' && Number.isFinite(cell));

function checkData(sheet: Worksheet, seen: Seen, bounds: Bounds) {
  for (const row of sheet.data) {
    if (!Array.isArray(row)) throw new TypeError('Expected a worksheet row.');
    if (row.length > 16_384) throw new RangeError('Worksheet exceeds the column limit.');
    seen.cells += row.length;
    if (seen.cells > bounds.maxCells) throw new RangeError('Workbook exceeds the cell limit.');
    for (const cell of row) {
      if (badCell(cell)) throw new TypeError('Invalid cell value.');
      if (typeof cell === 'string' && cell.length > bounds.maxCellLength)
        throw new RangeError('Cell exceeds the text limit.');
    }
  }
}

function checkWorksheet(
  sheet: Worksheet,
  { workbook, bounds, seen }: { workbook: Workbook; bounds: Bounds; seen: Seen },
) {
  checkIdentity(sheet, seen);
  checkLines(sheet, bounds);
  seen.ids.add(sheet.id);
  seen.names.add(worksheetNameKey(sheet.name));
  seen.rows += sheet.data.length;
  if (seen.rows > bounds.maxCells || sheet.data.length > 1_048_576)
    throw new RangeError('Worksheet exceeds the row limit.');
  checkGeometry(sheet, bounds);
  checkPresentation(sheet);
  checkInputs(sheet, bounds);
  checkExtent(sheet, workbook, bounds);
  seen.cells += metadataCells(sheet);
  checkData(sheet, seen, bounds);
}

/** Validate the same bounded JSON contract on the client and at the server boundary. */
export function assertWorkbook(
  value: unknown,
  limits: WorkbookLimits = {},
): asserts value is Workbook {
  const bounds = { ...DEFAULT_WORKBOOK_LIMITS, ...limits };
  if (!value || typeof value !== 'object') throw new TypeError('Expected a workbook.');
  const workbook = value as Partial<Workbook>;
  if (
    workbook.schemaVersion !== 1 ||
    !Array.isArray(workbook.worksheets) ||
    !workbook.worksheets.length ||
    workbook.worksheets.length > bounds.maxWorksheets
  )
    throw new TypeError('Invalid workbook version or worksheet count.');
  assertReviewData(workbook as Workbook);
  if (
    workbook.name !== undefined &&
    (typeof workbook.name !== 'string' || workbook.name.length > 512)
  )
    throw new TypeError('Invalid workbook name.');
  const seen: Seen = { ids: new Set(), names: new Set(), cells: 0, rows: 0 };
  for (const sheet of workbook.worksheets)
    checkWorksheet(sheet, { workbook: workbook as Workbook, bounds, seen });
}

export function assertReviewData(workbook: Workbook): void {
  checkReviewData(workbook, (snapshot) => assertWorkbook(snapshot));
}
