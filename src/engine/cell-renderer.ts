import type { DisplayOptions } from '../calculation/format.js';
import type { CellValue } from '../calculation/values.js';
import type { CellFormat } from '../model/review.js';
import type { CellInput } from '../model/workbook.js';
import {
  type CellPortal,
  forgetCell,
  type GridCellPresentation,
  paintCell,
  paintDraftCell,
  type PaintFrame,
  releaseCell,
} from './cell-paint.js';
import { gridCellAt, gridWorksheet } from './grid-cell.js';
import { observeGridCellWrites } from './grid-cell-writes.js';
import type {
  JSpreadsheetHostHandle,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';
import { gridRowWindow } from './row-window.js';

export type { CellPortal, GridDraft } from './cell-paint.js';

const previousDraftCells = new WeakMap<JSpreadsheetHostHandle, HTMLElement>();
const cellWrites = new WeakMap<JSpreadsheetHostHandle, ReturnType<typeof observeGridCellWrites>>();
/**
 * What a row was last painted from. A values row keeps its identity only while no cell in it
 * changes value or readiness, so a row whose values, input, formats, hints and settings are all
 * the same objects is left as it is without visiting its cells (SE-PERF-1.1).
 */
interface PaintedRow {
  settings: object;
  values: readonly CellValue[] | undefined;
  raw: readonly CellInput[] | undefined;
  formats: Record<string, CellFormat>;
  hints: unknown;
  transient: boolean;
  cells: HTMLElement[];
  /** Each cell's first child after painting, as in PaintedCell. */
  nodes: (ChildNode | null)[];
  portals: CellPortal[];
}
const NO_FORMATS: Record<string, CellFormat> = Object.freeze({});
const paintedRows = new WeakMap<JSpreadsheetHostHandle, Map<number, PaintedRow>>();
/** Settings shared by every cell; a new object means every cell is painted again. */
const paintSettings = new WeakMap<
  JSpreadsheetHostHandle,
  { inputs: readonly unknown[]; token: object }
>();

/** A paint of one grid: the shared frame, the engine's cells and what the paint touched. */
interface GridPaint extends PaintFrame {
  worksheet: JSpreadsheetWorksheetInstance;
  rows: (CellValue | CellInput)[][];
  rowMemo: Map<number, PaintedRow>;
  dirtyRows: ReadonlySet<number> | undefined;
}
/** How calculated cells are painted; every field is optional. */
interface CellPaintOptions {
  display?: DisplayOptions | undefined;
  raw?: CellInput[][] | undefined;
  unavailableText?: string | undefined;
  formats?: Record<string, CellFormat> | undefined;
  showFormulas?: boolean | undefined;
  presentation?: GridCellPresentation | undefined;
  /** Paints every cell again, for when other presentation has changed the cells. */
  repaint?: boolean | undefined;
}

/** Release only presentation owned by this renderer; authored styles remain untouched. */
export function clearCalculatedPresentation(handle: JSpreadsheetHostHandle) {
  for (const row of paintedRows.get(handle)?.values() ?? [])
    for (const cell of row.cells) releaseCell(cell);
  const draft = previousDraftCells.get(handle);
  if (draft) releaseCell(draft);
  previousDraftCells.delete(handle);
  cellWrites.get(handle)?.dispose();
  cellWrites.delete(handle);
  paintSettings.delete(handle);
  paintedRows.delete(handle);
}

function observeWrites(handle: JSpreadsheetHostHandle, worksheet: JSpreadsheetWorksheetInstance) {
  let writes = cellWrites.get(handle);
  if (writes?.table !== worksheet.table) {
    writes?.dispose();
    writes = worksheet.table && observeGridCellWrites(worksheet.table);
    cellWrites.set(handle, writes);
  }
  return writes;
}

/** The settings token for this paint: unchanged while every shared input is unchanged. */
function settingsToken(
  handle: JSpreadsheetHostHandle,
  inputs: readonly unknown[],
  force: boolean,
): object {
  const previous = paintSettings.get(handle);
  if (
    !force &&
    previous &&
    previous.inputs.every((input, index) => Object.is(input, inputs[index]))
  )
    return previous.token;
  const token = {};
  paintSettings.set(handle, { inputs, token });
  return token;
}

/** Host presentation applies only to the publication the grid accepted. */
function acceptedHints(presentation: GridCellPresentation | undefined) {
  const accepted = presentation?.publication;
  const snapshot = presentation?.cellPresentation;
  return accepted &&
    snapshot?.inputKey === accepted.inputKey &&
    snapshot.revision === accepted.revision
    ? snapshot.cells
    : undefined;
}

function paintSettingsOf(handle: JSpreadsheetHostHandle, options: CellPaintOptions): object {
  const { presentation } = options,
    accepted = presentation?.publication;
  return settingsToken(
    handle,
    [
      options.display,
      options.unavailableText ?? '…',
      options.showFormulas ?? false,
      presentation?.sheetId,
      presentation?.renderCellContent,
      // Host content receives the publication; text alone does not depend on it.
      presentation?.renderCellContent && accepted?.inputKey,
      presentation?.renderCellContent && accepted?.revision,
    ],
    options.repaint ?? false,
  );
}

function startPaint(
  handle: JSpreadsheetHostHandle,
  worksheet: JSpreadsheetWorksheetInstance,
  {
    values,
    options,
    dirtyRows,
  }: {
    values: CellValue[][] | undefined;
    options: CellPaintOptions;
    dirtyRows: GridPaint['dirtyRows'];
  },
): GridPaint {
  const raw = options.raw ?? [],
    formats = options.formats ?? NO_FORMATS;
  const rows = options.raw ?? values ?? [];
  const rowMemo = paintedRows.get(handle) ?? new Map<number, PaintedRow>();
  paintedRows.set(handle, rowMemo);
  return {
    settings: paintSettingsOf(handle, options),
    values,
    raw,
    display: options.display,
    unavailableText: options.unavailableText ?? '…',
    formats,
    hasFormats: Object.keys(formats).length > 0,
    showFormulas: options.showFormulas ?? false,
    presentation: options.presentation,
    hints: acceptedHints(options.presentation),
    keyPrefix: `[${JSON.stringify(options.presentation?.sheetId)},`,
    portals: [],
    worksheet,
    rows,
    rowMemo,
    dirtyRows,
  };
}

/** Whether a row still shows what it was painted from, in the same engine cells. */
function rowUnchanged(paint: GridPaint, memo: PaintedRow, y: number): boolean {
  return (
    !memo.transient &&
    memo.values === paint.values?.[y] &&
    memo.raw === paint.raw[y] &&
    memo.settings === paint.settings &&
    memo.formats === paint.formats &&
    memo.hints === paint.hints &&
    paint.presentation?.draft?.y !== y &&
    // A rebuilt row has new cells; the engine replaces every cell of a row it rebuilds.
    memo.cells[0] === gridCellAt(paint.worksheet, 0, y) &&
    // Without an observable table, detect engine writes from the node identities.
    (paint.dirtyRows
      ? !paint.dirtyRows.has(y)
      : memo.cells.every((cell, index) => cell.firstChild === memo.nodes[index]))
  );
}

function releasePaint(cell: HTMLElement) {
  releaseCell(cell);
  if (cell.isConnected && !cell.classList.contains('editor')) cell.textContent = '';
}

function paintRow(paint: GridPaint, y: number) {
  const memo = paint.rowMemo.get(y);
  if (memo && rowUnchanged(paint, memo, y)) {
    for (const portal of memo.portals) paint.portals.push(portal);
    return;
  }
  const row = paint.rows[y];
  const cells: HTMLElement[] = [],
    rowPortals = paint.portals.length;
  let transient = paint.presentation?.draft?.y === y;
  for (let x = 0; x < row.length; x++) {
    const cell = gridCellAt(paint.worksheet, x, y);
    if (!cell) continue;
    if (cell.classList.contains('editor')) {
      // The engine rewrites the cell when its editor closes.
      forgetCell(cell);
      transient = true;
      continue;
    }
    cells.push(cell);
    paintCell(paint, cell, { x, y });
  }
  if (memo) {
    const current = new Set(cells);
    for (const cell of memo.cells) if (!current.has(cell)) releasePaint(cell);
  }
  paint.rowMemo.set(y, {
    settings: paint.settings,
    values: paint.values?.[y],
    raw: paint.raw[y],
    formats: paint.formats,
    hints: paint.hints,
    transient,
    cells,
    nodes: paint.dirtyRows ? [] : cells.map((cell) => cell.firstChild),
    portals: paint.portals.slice(rowPortals),
  });
}

function paintDraft(paint: GridPaint, previous: HTMLElement | undefined) {
  const draft = paint.presentation?.draft;
  // The row pass paints existing authored cells. A draft can also be in blank grid padding.
  const cell =
    draft && draft.x >= (paint.rows[draft.y]?.length ?? 0)
      ? gridCellAt(paint.worksheet, draft.x, draft.y)
      : null;
  if (previous && previous !== cell) releasePaint(previous);
  if (!cell?.isConnected || cell.classList.contains('editor')) return undefined;
  paintDraftCell(cell, draft?.text ?? '');
  return cell;
}

/** Only engine-owned cells are painted; raw data and formulas remain untouched. A cell
 * whose value, input, format, readiness and presentation are unchanged is left as it is.
 * `repaint` paints every cell again, for when other presentation has changed the cells.
 */
export function renderCalculatedCells(
  handle: JSpreadsheetHostHandle,
  values: CellValue[][] | undefined,
  options: CellPaintOptions = {},
): CellPortal[] {
  const worksheet = gridWorksheet(handle);
  if (!worksheet) return [];
  const writes = observeWrites(handle, worksheet);
  const paint = startPaint(handle, worksheet, {
    values,
    options,
    dirtyRows: writes?.changedRows(),
  });
  const visible = gridRowWindow(worksheet);
  const start = visible?.start ?? 0,
    end = Math.min(visible?.end ?? paint.rows.length, paint.rows.length);
  for (const [y, row] of paint.rowMemo) {
    if (y < start || y >= end) {
      for (const cell of row.cells) releasePaint(cell);
      paint.rowMemo.delete(y);
    }
  }
  const draft = paintDraft(paint, previousDraftCells.get(handle));
  if (draft) previousDraftCells.set(handle, draft);
  else previousDraftCells.delete(handle);
  for (let y = start; y < end; y++) paintRow(paint, y);
  writes?.ignorePaint();
  return paint.portals;
}
