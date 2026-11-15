import { requireDefined } from '../model/require-defined.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

/**
 * The engine builds a row element for every worksheet row; only the rows near the viewport are
 * attached to its table body (SE-PERF-1). Two spacer rows stand in for the detached rows above
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

/** Each row's height in CSS pixels: its written measure, or none while it is hidden. */
function rowHeights(engine: JSpreadsheetWorksheetInstance): number[] {
  const fallback = Number(engine.options?.defaultRowHeight) || 24;
  return (engine.rows ?? []).map((row) => {
    const element = row?.element as HTMLElement | undefined;
    if (!element || element.style.display === 'none') return 0;
    const height = Number.parseFloat(element.style.height);
    return Number.isFinite(height) ? height : fallback;
  });
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
  const count = engine.rows?.length ?? 0;
  const heights = rowHeights(engine);
  let above = 0,
    below = 0;
  for (let y = 0; y < state.start; y++) above += heights[y];
  for (let y = state.end; y < count; y++) below += heights[y];
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

/** Attach rows [start, end) and detach the rest, moving only rows that change sides. */
function place(
  engine: JSpreadsheetWorksheetInstance,
  state: RowWindowState,
  start: number,
  end: number,
): void {
  const tbody = requireDefined(engine.tbody);
  const count = engine.rows?.length ?? 0;
  [start, end] = widenForMerges(engine, Math.max(0, start), Math.min(count, end));
  if (start === state.start && end === state.end) return;
  state.top.remove();
  state.bottom.remove();
  const kept = start < state.end && end > state.start;
  for (let y = state.start; y < state.end; y++)
    if (!kept || y < start || y >= end) rowElement(engine, y)?.remove();
  const attached: number[] = [];
  const fragment = () => tbody.ownerDocument.createDocumentFragment();
  const attach = (from: number, to: number, before: Node | null) => {
    const rows = fragment();
    for (let y = from; y < to; y++) {
      const element = rowElement(engine, y);
      if (!element) continue;
      rows.append(element);
      attached.push(y);
    }
    tbody.insertBefore(rows, before);
  };
  if (!kept) attach(start, end, null);
  else {
    if (start < state.start) attach(start, state.start, rowElement(engine, state.start) ?? null);
    if (end > state.end) attach(state.end, end, null);
  }
  state.start = start;
  state.end = end;
  sizeSpacers(engine, state);
  if (attached.length) state.onAttach(attached);
}

/** The window that covers the viewport with overscan, or the current one while it still does. */
function wanted(
  engine: JSpreadsheetWorksheetInstance,
  state: RowWindowState,
  reveal?: [number, number],
): [number, number] {
  const viewport = engine.content;
  const count = engine.rows?.length ?? 0;
  if (!viewport) return [0, count];
  const heights = rowHeights(engine);
  const header = engine.table?.tHead?.offsetHeight ?? 0;
  const top = viewport.scrollTop - header,
    bottom = viewport.scrollTop + viewport.clientHeight - header;
  let first = 0,
    offset = 0;
  while (first < count - 1 && offset + heights[first] <= top) offset += heights[first++];
  let last = first;
  while (last < count && offset < bottom) offset += heights[last++];
  if (reveal) {
    first = Math.min(first, reveal[0]);
    last = Math.max(last, reveal[1] + 1);
  }
  if (
    !reveal &&
    state.rows === engine.rows &&
    state.start <= Math.max(0, first - MARGIN_ROWS) &&
    state.end >= Math.min(count, last + MARGIN_ROWS)
  )
    return [state.start, state.end];
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
    onAttach,
  };
  windows.set(engine, state);
  resetGridRowWindow(engine);
  const update = () => {
    if (state.rows !== engine.rows) resetGridRowWindow(engine);
    else {
      const [start, end] = wanted(engine, state);
      place(engine, state, start, end);
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
  state.top.remove();
  state.bottom.remove();
  // Whatever is attached now is contiguous from row 0.
  let attached = 0;
  while (attached < count && rowElement(engine, attached)?.parentNode === engine.tbody) attached++;
  state.start = 0;
  state.end = attached;
  if (count < MIN_WINDOWED_ROWS) {
    place(engine, state, 0, count);
    return;
  }
  const [start, end] = wanted(engine, { ...state, rows: [] });
  place(engine, state, start, end);
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
  const [start, end] = wanted(engine, state);
  place(engine, state, start, end);
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
    place(engine, state, first - OVERSCAN_ROWS, last + 1 + OVERSCAN_ROWS);
    return;
  }
  const [start, end] = wanted(engine, state, [first, last]);
  place(engine, state, start, end);
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
