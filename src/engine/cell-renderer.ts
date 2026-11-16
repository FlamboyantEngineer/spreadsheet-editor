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
import type {
  JSpreadsheetHostHandle,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';

export type { CellPortal, GridDraft } from './cell-paint.js';

export function worksheetFrom(
  handle: JSpreadsheetHostHandle,
): JSpreadsheetWorksheetInstance | undefined {
  const instance = handle.getInstance();
  if (!instance) return undefined;
  return Array.isArray(instance)
    ? instance[0]
    : 'worksheets' in instance
      ? instance.worksheets?.[0]
      : (instance as JSpreadsheetWorksheetInstance);
}
const previousPresentations = new WeakMap<JSpreadsheetHostHandle, Set<HTMLElement>>();
/**
 * What a row was last painted from. A values row keeps its identity only while no cell in it
 * changes value or readiness, so a row whose values, input, formats, hints and settings are all
 * the same objects is left as it is without visiting its cells (SE-PERF-1.1).
 */
interface PaintedRow {
  settings: object;
  /** The values row, or the input row when there are no values. */
  values: readonly (CellValue | CellInput)[];
  raw: readonly CellInput[] | undefined;
  formats: Record<string, CellFormat>;
  hints: unknown;
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

type CellLookup = Required<Pick<JSpreadsheetWorksheetInstance, 'getCellFromCoords'>>;
/** A paint of one grid: the shared frame, the engine's cells and what the paint touched. */
interface GridPaint extends PaintFrame {
  worksheet: CellLookup;
  rows: (CellValue | CellInput)[][];
  rowMemo: Map<number, PaintedRow>;
  touched: Set<HTMLElement>;
}
const hasCellLookup = (
  worksheet: JSpreadsheetWorksheetInstance | undefined,
): worksheet is JSpreadsheetWorksheetInstance & CellLookup => !!worksheet?.getCellFromCoords;
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
  for (const cell of previousPresentations.get(handle) ?? []) releaseCell(cell);
  previousPresentations.delete(handle);
  paintSettings.delete(handle);
  paintedRows.delete(handle);
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
  worksheet: CellLookup,
  { values, options }: { values: CellValue[][] | undefined; options: CellPaintOptions },
): GridPaint {
  const raw = options.raw ?? [],
    formats = options.formats ?? NO_FORMATS;
  const rows = values ?? raw;
  const rowMemo = paintedRows.get(handle) ?? new Map<number, PaintedRow>();
  paintedRows.set(handle, rowMemo);
  if (rowMemo.size > rows.length)
    for (const y of rowMemo.keys()) if (y >= rows.length) rowMemo.delete(y);
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
    touched: new Set(),
  };
}

/** Whether a row still shows what it was painted from, in the same engine cells. */
function rowUnchanged(
  paint: GridPaint,
  memo: PaintedRow | undefined,
  y: number,
): memo is PaintedRow {
  return (
    !!memo &&
    memo.values === paint.rows[y] &&
    memo.raw === paint.raw[y] &&
    memo.settings === paint.settings &&
    memo.formats === paint.formats &&
    memo.hints === paint.hints &&
    paint.presentation?.draft?.y !== y &&
    // A rebuilt row has new cells; the engine replaces every cell of a row it rebuilds.
    memo.cells[0] === paint.worksheet.getCellFromCoords(0, y) &&
    // The engine writes its raw value into a cell when an editor closes.
    memo.cells.every((cell, index) => cell.firstChild === memo.nodes[index])
  );
}

function paintRow(paint: GridPaint, y: number) {
  const memo = paint.rowMemo.get(y);
  if (rowUnchanged(paint, memo, y)) {
    for (const cell of memo.cells) paint.touched.add(cell);
    for (const portal of memo.portals) paint.portals.push(portal);
    return;
  }
  const row = paint.rows[y];
  const cells: HTMLElement[] = [],
    rowPortals = paint.portals.length;
  let settled = paint.presentation?.draft?.y !== y;
  for (let x = 0; x < row.length; x++) {
    const cell = paint.worksheet.getCellFromCoords(x, y);
    if (!cell) continue;
    if (cell.classList.contains('editor')) {
      // The engine rewrites the cell when its editor closes.
      forgetCell(cell);
      settled = false;
      continue;
    }
    cells.push(cell);
    paint.touched.add(cell);
    paintCell(paint, cell, { x, y });
  }
  if (!settled) return void paint.rowMemo.delete(y);
  paint.rowMemo.set(y, {
    settings: paint.settings,
    values: row,
    raw: paint.raw[y],
    formats: paint.formats,
    hints: paint.hints,
    cells,
    nodes: cells.map((cell) => cell.firstChild),
    portals: paint.portals.slice(rowPortals),
  });
}

function paintDraft(paint: GridPaint) {
  const draft = paint.presentation?.draft;
  if (!draft) return;
  const cell = paint.worksheet.getCellFromCoords(draft.x, draft.y);
  if (!cell || cell.classList.contains('editor')) return;
  paintDraftCell(cell, draft.text);
  paint.touched.add(cell);
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
  const worksheet = worksheetFrom(handle);
  if (!hasCellLookup(worksheet)) return [];
  const previous = previousPresentations.get(handle) ?? new Set<HTMLElement>();
  const paint = startPaint(handle, worksheet, { values, options });
  for (let y = 0; y < paint.rows.length; y++) paintRow(paint, y);
  paintDraft(paint);
  for (const cell of previous)
    if (!paint.touched.has(cell)) {
      releaseCell(cell);
      if (!cell.classList.contains('editor')) cell.textContent = '';
    }
  previousPresentations.set(handle, paint.touched);
  return paint.portals;
}
