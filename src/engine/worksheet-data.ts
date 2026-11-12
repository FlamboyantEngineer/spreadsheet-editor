import { DEFAULT_WORKBOOK_LIMITS, type Worksheet } from '../model/workbook.js';
import {
  extendWorksheetIdentities,
  worksheetExtent,
  type WorkbookIdFactory,
  type WorksheetExtent,
} from '../model/extent.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { CellInput } from '../model/workbook.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

/** Capture a vendor batch without turning its padded blank viewport into authored data. */
export function captureWorksheetCellChanges(
  sheet: Worksheet,
  changes: readonly { x: number; y: number; value: CellInput | undefined }[],
  idFactory?: WorkbookIdFactory,
): Worksheet {
  if (changes.length > DEFAULT_WORKBOOK_LIMITS.maxCells)
    throw new RangeError('Grid changes exceed the cell limit.');
  const empty = (input: CellInput | undefined) => input == null || input === '';
  const extent = worksheetExtent(sheet);
  const normalized = changes.map((change) => {
    const x = Number(change.x),
      y = Number(change.y),
      value = change.value;
    if (
      !Number.isSafeInteger(x) ||
      !Number.isSafeInteger(y) ||
      x < 0 ||
      y < 0 ||
      x > 16_383 ||
      y > 1_048_575
    )
      throw new RangeError('Invalid grid cell coordinates.');
    if (typeof value === 'string' && value.length > DEFAULT_WORKBOOK_LIMITS.maxCellLength)
      throw new RangeError('Grid input exceeds the cell length limit.');
    const previous = sheet.data[y]?.[x];
    if (value !== previous && !(empty(value) && empty(previous))) {
      extent.columns = Math.max(extent.columns, x + 1);
      extent.rows = Math.max(extent.rows, y + 1);
    }
    if (extent.columns * extent.rows > DEFAULT_WORKBOOK_LIMITS.maxCells)
      throw new RangeError('Grid changes exceed the cell limit.');
    return { x, y, value };
  });
  const next = { ...sheet, data: sheet.data.map((row) => [...row]) };
  if (sheet.inputValues) next.inputValues = { ...sheet.inputValues };
  for (const { x, y, value } of normalized) {
    const previous = next.data[y]?.[x];
    if (Object.is(previous, value) || (empty(previous) && empty(value))) continue;
    while (next.data.length <= y) next.data.push([]);
    while (next.data[y].length <= x) next.data[y].push(null);
    next.data[y][x] = value ?? null;
    if (next.inputValues) delete next.inputValues[formatSpreadsheetAddress(x, y)];
  }
  extendWorksheetIdentities(next, idFactory);
  return next;
}

export function gridWorksheetExtent(
  sheet: Worksheet,
  logicalExtent?: WorksheetExtent,
): [number, number] {
  const extent = worksheetExtent(sheet);
  return [
    Math.max(extent.columns, logicalExtent?.columns ?? 0, sheet.minDimensions ? 1 : 26),
    Math.max(extent.rows, logicalExtent?.rows ?? 0, sheet.minDimensions ? 1 : 100),
  ];
}

/** setData rebuilds body rows, but CE keeps existing column/header structures.
 * Resize those structures first, then set current logical minima before body rebuild.
 * The surrounding controller suppresses authored callbacks for this presentation work.
 */
export function syncWorksheetData(
  engine: JSpreadsheetWorksheetInstance,
  next: Worksheet,
  logicalExtent?: WorksheetExtent,
) {
  const options = engine.options;
  if (!options) return;
  const [columns, rows] = gridWorksheetExtent(next, logicalExtent);
  const currentColumns = engine.cols?.length ?? options.columns?.length ?? columns;
  const permissions = {
    allowInsertColumn: options.allowInsertColumn,
    allowDeleteColumn: options.allowDeleteColumn,
  };
  const merges = options.mergeCells;
  options.mergeCells = {};
  options.minDimensions = [columns, rows];
  options.allowInsertColumn = true;
  options.allowDeleteColumn = true;
  try {
    if (currentColumns < columns) engine.insertColumn?.(columns - currentColumns);
    if (currentColumns > columns) engine.deleteColumn?.(columns, currentColumns - columns);
    options.columns = Array.from({ length: columns }, (_, x) => ({
      ...next.columns?.[x],
      type: 'text',
    }));
    options.rows = structuredClone(next.rows ?? {});
    options.mergeCells = structuredClone(next.mergeCells ?? {});
    engine.setData?.(next.data.map((row) => [...row]));
  } finally {
    options.allowInsertColumn = permissions.allowInsertColumn;
    options.allowDeleteColumn = permissions.allowDeleteColumn;
    if (!options.mergeCells) options.mergeCells = merges;
  }
}
