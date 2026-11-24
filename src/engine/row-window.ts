import { requireDefined } from '../model/require-defined.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';

/**
 * The engine builds a row element for every worksheet row; only the rows near the viewport are
 * attached to its table body (SE-PERF-1.1). Two spacer rows stand in for the detached rows above
 * and below, so the scrollbar spans the whole worksheet and every attached row sits where it
 * would if all rows were attached.
 */
interface RowWindowState {
  start: number;
  end: number;
  top: HTMLTableRowElement;
  bottom: HTMLTableRowElement;
  /** The table body the window was built in; the engine replaces its rows when it rebuilds. */
  rows: readonly unknown[];
  offsets: Float64Array | undefined;
  onAttach: (rows: number[]) => void;
}
const windows = new WeakMap<JSpreadsheetWorksheetInstance, RowWindowState>();

/** Rows kept attached beyond each edge of the viewport. */
const OVERSCAN_ROWS = 20;
/** The window moves only once the viewport comes within this many rows of its edge. */
const MARGIN_ROWS = 5;
/** Rows the engine attaches when it builds or rebuilds its table body. */
const INITIAL_ROWS = 100;
/** Attached rows below this count are all attached; windowing would save nothing. */
const MIN_WINDOWED_ROWS = 2 * INITIAL_ROWS;

export const GRID_SPACER_CLASS = 'se-editor__grid-spacer';

function rowElement(engine: JSpreadsheetWorksheetInstance, y: number) {
  return engine.rows?.[y]?.element as HTMLTableRowElement | undefined;
}

function spacer(engine: JSpreadsheetWorksheetInstance): HTMLTableRowElement {
  const document = requireDefined(engine.tbody).ownerDocument;
  const row = document.createElement('tr');
  row.className = GRID_SPACER_CLASS;
  row.setAttribute('aria-hidden', 'true');
  const cell = document.createElement('td');
  // Geometry only: the spacer's height is the detached rows' height and nothing else.
  cell.style.padding = '0';
  cell.style.border = '0';
  row.append(cell);
  return row;
}

/** Cached offsets allow scroll to find rows without reading every row's style. */
function rowOffsets(engine: JSpreadsheetWorksheetInstance, state: RowWindowState): Float64Array {
  if (state.offsets) return state.offsets;
  const rows = engine.rows ?? [];
  const offsets = new Float64Array(rows.length + 1);
  const fallback = Number(engine.options?.defaultRowHeight) || 24;
  rows.forEach((row, index) => {
    const style = row?.element?.style;
    const shown = style && style.display !== 'none';
    const height = Number.parseFloat(style?.height ?? '');
    offsets[index + 1] =
      offsets[index] + (shown ? (Number.isFinite(height) ? height : fallback) : 0);
  });
  state.offsets = offsets;
  return offsets;
}

/** A changed row measure invalidates the prefix sums for the next geometry read. */
export function invalidateGridRowHeights(engine: JSpreadsheetWorksheetInstance) {
  const state = windows.get(engine);
  if (state) state.offsets = undefined;
}

/** Rows covered by a merge that starts in [start, end), or that crosses either edge. */
function widenForMerges(
  engine: JSpreadsheetWorksheetInstance,
  start: number,
  end: number,
): [number, number] {
  const merges = engine.options?.mergeCells as Record<string, [number, number]> | undefined;
  if (!merges) return [start, end];
  let changed = true;
  while (changed) {
    changed = false;
    for (const [address, span] of Object.entries(merges)) {
      const anchor = parseSpreadsheetAddress(address);
      if (!anchor) continue;
      const first = anchor.y,
        last = anchor.y + (span?.[1] || 1);
      if (first < end && last > start && (first < start || last > end)) {
        start = Math.min(start, first);
        end = Math.max(end, last);
        changed = true;
      }
    }
  }
  return [start, end];
}

/** The attached rows, as [start, end), or undefined when every row is attached. */
export function gridRowWindow(
  engine: JSpreadsheetWorksheetInstance,
): { start: number; end: number } | undefined {
  const state = windows.get(engine);
  return state && state.rows === engine.rows ? { start: state.start, end: state.end } : undefined;
}

/** Give the spacers the height of the detached rows above and below the window. */
function sizeSpacers(engine: JSpreadsheetWorksheetInstance, state: RowWindowState): void {
  const tbody = requireDefined(engine.tbody);
  const offsets = rowOffsets(engine, state);
  const above = offsets[state.start],
    below = offsets[offsets.length - 1] - offsets[state.end];
  // With no rows detached on a side, the first or last row keeps its :first-child or
  // :last-child edge in the stylesheet.
  const columns = (engine.cols?.length ?? 0) + 1;
  for (const [row, height, before] of [
    [state.top, above, rowElement(engine, state.start) ?? null],
    [state.bottom, below, null],
  ] as const) {
    if (!height) {
      row.remove();
      continue;
    }
    const cell = row.firstElementChild as HTMLTableCellElement;
    if (cell.colSpan !== columns) cell.colSpan = columns;
    if (cell.style.height !== `${height}px`) cell.style.height = `${height}px`;
    if (row.parentNode !== tbody || (before && row.nextSibling !== before))
      tbody.insertBefore(row, before);
  }
}

type Span = [start: number, end: number];

/** Appends the existing rows [from, to) before a node, returning the rows it attached. */
function attachRows(
  engine: JSpreadsheetWorksheetInstance,
  [from, to]: Span,
  before: Node | null,
): number[] {
  const tbody = requireDefined(engine.tbody);
  const rows = tbody.ownerDocument.createDocumentFragment();
  const attached: number[] = [];
  for (let y = from; y < to; y++) {
    const element = rowElement(engine, y);
    if (!element) continue;
    rows.append(element);
    attached.push(y);
  }
  tbody.insertBefore(rows, before);
  return attached;
}

/** Moves the window to [start, end), returning the rows it attached. */
function moveWindow(
  engine: JSpreadsheetWorksheetInstance,
  state: RowWindowState,
  [start, end]: Span,
) {
  state.top.remove();
  state.bottom.remove();
  const kept = start < state.end && end > state.start;
  for (let y = state.start; y < state.end; y++)
    if (!kept || y < start || y >= end) rowElement(engine, y)?.remove();
  if (!kept) return attachRows(engine, [start, end], null);
  return [
    ...(start < state.start
      ? attachRows(engine, [start, state.start], rowElement(engine, state.start) ?? null)
      : []),
    ...(end > state.end ? attachRows(engine, [state.end, end], null) : []),
  ];
}

/** Attach rows [start, end) and detach the rest, moving only rows that change sides. */
function place(engine: JSpreadsheetWorksheetInstance, state: RowWindowState, span: Span): void {
  const count = engine.rows?.length ?? 0;
  const [start, end] = widenForMerges(engine, Math.max(0, span[0]), Math.min(count, span[1]));
  if (start === state.start && end === state.end) return;
  const attached = moveWindow(engine, state, [start, end]);
  state.start = start;
  state.end = end;
  sizeSpacers(engine, state);
  state.onAttach(attached);
}

/** The rows the viewport shows, by their heights below the column header. */
function visibleRows(engine: JSpreadsheetWorksheetInstance, state: RowWindowState): Span {
  const viewport = requireDefined(engine.content);
  const count = engine.rows?.length ?? 0;
  const offsets = rowOffsets(engine, state);
  const header = engine.table?.tHead?.offsetHeight ?? 0;
  const top = viewport.scrollTop - header,
    bottom = viewport.scrollTop + viewport.clientHeight - header;
  const first = Math.min(Math.max(count - 1, 0), Math.max(offsetIndex(offsets, top, true) - 1, 0));
  const last = Math.max(first, Math.min(count, offsetIndex(offsets, bottom, false)));
  return [first, last];
}

/** First offset meeting the threshold, skipping equal offsets for the top edge. */
function offsetIndex(offsets: Float64Array, target: number, skipEqual: boolean) {
  let low = 0,
    high = offsets.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    const offset = offsets[middle];
    if (offset < target || (skipEqual && offset === target)) low = middle + 1;
    else high = middle;
  }
  return low;
}

/** Whether the current window still covers the visible rows with their margin. */
function covers(engine: JSpreadsheetWorksheetInstance, state: RowWindowState, [first, last]: Span) {
  const count = engine.rows?.length ?? 0;
  return (
    state.rows === engine.rows &&
    state.start <= Math.max(0, first - MARGIN_ROWS) &&
    state.end >= Math.min(count, last + MARGIN_ROWS)
  );
}

/** The window that covers the viewport with overscan, or the current one while it still does. */
function wanted(
  engine: JSpreadsheetWorksheetInstance,
  state: RowWindowState,
  reveal?: [number, number],
): Span {
  const viewport = engine.content;
  if (!viewport) return [0, engine.rows?.length ?? 0];
  let [first, last] = visibleRows(engine, state);
  if (reveal) {
    first = Math.min(first, reveal[0]);
    last = Math.max(last, reveal[1] + 1);
  } else if (covers(engine, state, [first, last])) return [state.start, state.end];
  return [first - OVERSCAN_ROWS, last + OVERSCAN_ROWS];
}

/**
 * Window the engine's rows. `onAttach` receives the rows each move attaches, after they are in
 * place, so presentation that needs layout can be applied to them. Returns a disposer.
 */
export function windowGridRows(
  engine: JSpreadsheetWorksheetInstance,
  onAttach: (rows: number[]) => void,
): () => void {
  const viewport = engine.content;
  if (!engine.tbody || !viewport) return () => {};
  const state: RowWindowState = {
    start: 0,
    end: 0,
    top: spacer(engine),
    bottom: spacer(engine),
    rows: [],
    offsets: undefined,
    onAttach,
  };
  windows.set(engine, state);
  resetGridRowWindow(engine);
  const update = () => {
    if (state.rows !== engine.rows) resetGridRowWindow(engine);
    else {
      if (!state.offsets) sizeSpacers(engine, state);
      place(engine, state, wanted(engine, state));
    }
  };
  viewport.addEventListener('scroll', update, { passive: true });
  return () => {
    viewport.removeEventListener('scroll', update);
    windows.delete(engine);
  };
}

/**
 * Take over the rows the engine attached after it built or rebuilt its table body: the first
 * INITIAL_ROWS rows when it built with `lazyLoading`, or all of them otherwise.
 */
function resetGridRowWindow(engine: JSpreadsheetWorksheetInstance): void {
  const state = windows.get(engine);
  if (!state || !engine.tbody) return;
  const count = engine.rows?.length ?? 0;
  state.rows = engine.rows ?? [];
  state.offsets = undefined;
  state.top.remove();
  state.bottom.remove();
  // Whatever is attached now is contiguous from row 0.
  let attached = 0;
  while (attached < count && rowElement(engine, attached)?.parentNode === engine.tbody) attached++;
  state.start = 0;
  state.end = attached;
  if (count < MIN_WINDOWED_ROWS) {
    place(engine, state, [0, count]);
    return;
  }
  const [first, last] = visibleRows(engine, state);
  place(engine, state, [first - OVERSCAN_ROWS, last + OVERSCAN_ROWS]);
}

/** Resize the spacers after row heights change, as zoom and row resizing do, and follow the viewport. */
export function refreshGridRowWindow(engine: JSpreadsheetWorksheetInstance): void {
  const state = windows.get(engine);
  if (!state) return;
  if (state.rows !== engine.rows) {
    resetGridRowWindow(engine);
    return;
  }
  sizeSpacers(engine, state);
  place(engine, state, wanted(engine, state));
}

/** Attach rows y1 to y2 before something measures them or scrolls to them. */
export function revealGridRows(engine: JSpreadsheetWorksheetInstance, y1: number, y2 = y1): void {
  const state = windows.get(engine);
  if (!state || state.rows !== engine.rows) return;
  const first = Math.min(y1, y2),
    last = Math.max(y1, y2);
  if (first >= state.start && last < state.end) return;
  // A far row joins a window around itself rather than stretching the current one.
  const span = state.end - state.start;
  if (first > state.end + span || last < state.start - span) {
    place(engine, state, [first - OVERSCAN_ROWS, last + 1 + OVERSCAN_ROWS]);
    return;
  }
  place(engine, state, wanted(engine, state, [first, last]));
}

/** Run an engine rebuild so that it attaches only its first rows, as its lazy mode does. */
export function withLazyRebuild<T>(engine: JSpreadsheetWorksheetInstance, rebuild: () => T): T {
  const options = engine.options as { lazyLoading?: boolean } | undefined;
  if (!options || !windows.has(engine)) return rebuild();
  options.lazyLoading = true;
  try {
    return rebuild();
  } finally {
    options.lazyLoading = false;
    resetGridRowWindow(engine);
  }
}
