import {
  extendWorksheetIdentities,
  gridWorksheetExtent,
  type WorkbookIdFactory,
  type WorksheetExtent,
  worksheetExtent,
} from '../model/extent.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { type CellInput, DEFAULT_WORKBOOK_LIMITS, type Worksheet } from '../model/workbook.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { withLazyRebuild } from './row-window.js';

function copyCellStorage(sheet: Worksheet): Worksheet {
  return {
    ...sheet,
    data: [...sheet.data],
    ...(sheet.inputValues && { inputValues: { ...sheet.inputValues } }),
    ...(sheet.rowIds && { rowIds: [...sheet.rowIds] }),
    ...(sheet.columnIds && { columnIds: [...sheet.columnIds] }),
  };
}

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
  let next = sheet;
  const edits = new Set<number>();
  for (const { x, y, value } of normalized) {
    const previous = next.data[y]?.[x];
    if (Object.is(previous, value) || (empty(previous) && empty(value))) continue;
    if (next === sheet) next = copyCellStorage(sheet);
    while (next.data.length <= y) next.data.push([]);
    if (!edits.has(y)) {
      next.data[y] = [...next.data[y]];
      edits.add(y);
    }
    while (next.data[y].length <= x) next.data[y].push(null);
    next.data[y][x] = value ?? null;
    if (next.inputValues) delete next.inputValues[formatSpreadsheetAddress(x, y)];
  }
  if (next !== sheet) extendWorksheetIdentities(next, idFactory);
  return next;
}

/** At most this many cells are updated in place; a larger change rebuilds the rows. */
const PATCH_LIMIT = 2_000;

type CellPatch = { x: number; y: number; value: CellInput };

/** Adds one row's changed values, or returns false when the row has another width. */
function rowChanges(
  engine: JSpreadsheetWorksheetInstance,
  { data, next, columns }: { data: CellInput[][]; next: Worksheet; columns: number },
  { y, changes }: { y: number; changes: CellPatch[] },
): boolean {
  const current = data[y],
    source = next.data[y];
  if (current?.length !== columns || engine.records?.[y]?.length !== columns) return false;
  for (let x = 0; x < columns; x++) {
    // The engine holds an empty string where the workbook has no value.
    const value = source?.[x] ?? '';
    if (current[x] === value) continue;
    changes.push({ x, y, value });
    if (changes.length > PATCH_LIMIT) return false;
  }
  return true;
}

/** Update the engine's values in place when only cell values changed and the grid keeps
 * its extent: rebuilding every row makes the browser restyle and lay out every cell.
 * Returns false, changing nothing, when the rows must be rebuilt with `syncWorksheetData`.
 */
export function patchWorksheetData(
  engine: JSpreadsheetWorksheetInstance,
  next: Worksheet,
  logicalExtent?: WorksheetExtent,
): boolean {
  const data = engine.options?.data as CellInput[][] | undefined;
  const [columns, rows] = gridWorksheetExtent(next, logicalExtent);
  const records = engine.records;
  if (!data || data.length !== rows || records?.length !== rows) return false;
  const changes: CellPatch[] = [];
  for (let y = 0; y < rows; y++)
    if (!rowChanges(engine, { data, next, columns }, { y, changes })) return false;
  for (const { x, y, value } of changes) {
    data[y][x] = value;
    const cell = records[y][x]?.element;
    if (cell instanceof HTMLElement && !cell.classList.contains('editor'))
      cell.textContent = String(value);
  }
  return true;
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
    // A rebuild attaches only the first rows; the row window places the rest (SE-PERF-1.1).
    withLazyRebuild(engine, () => engine.setData?.(next.data.map((row) => [...row])));
  } finally {
    options.allowInsertColumn = permissions.allowInsertColumn;
    options.allowDeleteColumn = permissions.allowDeleteColumn;
    if (!options.mergeCells) options.mergeCells = merges;
  }
}
