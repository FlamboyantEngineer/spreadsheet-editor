import type { CellSelection } from '../model/editing.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { Worksheet } from '../model/workbook.js';
import { gridCellAt } from './grid-cell.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { gridRowWindow, revealGridRows } from './row-window.js';

/** Zoom is presentation state; all engine measurements are already CSS pixels. */
export function gridZoomScale(zoom = 100): number {
  return Number.isFinite(zoom) && zoom > 0 ? zoom / 100 : 1;
}

export const DEFAULT_COLUMN_WIDTH = 100;
export const DEFAULT_ROW_HEIGHT = 24;
export const DEFAULT_ROW_HEADER_WIDTH = 50;

export function gridViewport(engine: JSpreadsheetWorksheetInstance): HTMLElement | undefined {
  return engine.content;
}

/** Preserve a logical scroll anchor without asking the browser to scroll any ancestor. */
export function restoreGridScroll(
  viewport: HTMLElement | undefined,
  anchor: { left: number; top: number },
  scale: { from: number; to: number },
) {
  if (!viewport) return;
  viewport.scrollLeft = (anchor.left * scale.to) / scale.from;
  viewport.scrollTop = (anchor.top * scale.to) / scale.from;
}

/** A column's rendered width, or nothing while it is hidden. */
function columnWidth(engine: JSpreadsheetWorksheetInstance, x: number): number {
  const col = engine.cols?.[x]?.colElement;
  if (col?.style.display === 'none') return 0;
  return Number.parseFloat(col?.style.width || col?.getAttribute('width') || '0') || 0;
}

/** The engine's first-row cells of a column, each visited once across merged rows. */
function* columnCells(
  records: NonNullable<JSpreadsheetWorksheetInstance['records']>,
  x: number,
  positioned: Set<HTMLElement>,
): Generator<HTMLElement> {
  for (const row of records) {
    const cell = row?.[x]?.element;
    if (!(cell instanceof HTMLElement) || positioned.has(cell)) continue;
    positioned.add(cell);
    yield cell;
  }
}

const cornerOf = (engine: JSpreadsheetWorksheetInstance) => engine.table?.tHead?.rows[0]?.cells[0];
const previousFrozenCounts = new WeakMap<JSpreadsheetWorksheetInstance, number>();

/** Sticky row headers and corner, so they stay over scrolled cells. */
function pinHeaders(engine: JSpreadsheetWorksheetInstance, corner: HTMLElement | undefined) {
  if (corner) position(corner, { position: 'sticky', left: '0px', zIndex: '4' });
  for (const row of engine.rows ?? []) {
    const header = row?.element?.firstElementChild;
    if (header instanceof HTMLElement)
      position(header, { position: 'sticky', left: '0px', zIndex: '2' });
  }
}

/** Marks one column frozen or not, at the given sticky offset. */
function freezeColumn(
  engine: JSpreadsheetWorksheetInstance,
  { x, frozen, left }: { x: number; frozen: boolean; left: number },
  positioned: Set<HTMLElement>,
) {
  const offset = frozen ? `${left}px` : '';
  const header = engine.headers?.[x];
  if (header) {
    header.classList.toggle('jss_freezed', frozen);
    position(header, { position: header.style.position, left: offset, zIndex: frozen ? '3' : '' });
  }
  for (const cell of columnCells(engine.records ?? [], x, positioned)) {
    cell.classList.toggle('jss_freezed', frozen);
    position(cell, { position: frozen ? 'sticky' : '', left: offset, zIndex: frozen ? '1' : '' });
  }
}

/** CE's freeze implementation rounds widths and offsets body cells relative to scrollLeft.
 * Keep its flags, but use measured, fractional sticky offsets for the rendered view.
 */
export function syncFrozenColumns(engine: JSpreadsheetWorksheetInstance) {
  const count = engine.options?.freezeColumns ?? 0;
  const previous = previousFrozenCounts.get(engine) ?? 0;
  previousFrozenCounts.set(engine, count);
  const corner = cornerOf(engine);
  // Measuring forces layout; only frozen columns need the row header's width.
  const rowHeaderWidth =
    (count > 0 ? corner?.getBoundingClientRect().width : undefined) ?? DEFAULT_ROW_HEADER_WIDTH;
  pinHeaders(engine, corner);
  let left = rowHeaderWidth;
  const positioned = new Set<HTMLElement>();
  const last = Math.min(Math.max(count, previous), engine.cols?.length ?? 0);
  for (let x = 0; x < last; x++) {
    const frozen = x < count;
    freezeColumn(engine, { x, frozen, left }, positioned);
    if (frozen) left += columnWidth(engine, x);
  }
}

/** Write only changed values: every write is a mutation the engine's observers handle. */
function position(
  element: HTMLElement,
  { position, left, zIndex }: { position: string; left: string; zIndex: string },
) {
  if (element.style.position !== position) element.style.position = position;
  if (element.style.left !== left) element.style.left = left;
  if (element.style.zIndex !== zIndex) element.style.zIndex = zIndex;
}

function mountedRecords(engine: JSpreadsheetWorksheetInstance, entering?: readonly number[]) {
  const records = engine.records ?? [];
  if (entering) return entering.map((row) => records[row]);
  const attached = gridRowWindow(engine);
  return attached ? records.slice(attached.start, attached.end) : records;
}

/** Repair visible offsets after CE scrolls, on resize previews, or as rows come into view. */
export function syncFrozenColumnOffsets(
  engine: JSpreadsheetWorksheetInstance,
  entering?: readonly number[],
) {
  const count = engine.options?.freezeColumns ?? 0;
  if (!count) return;
  let left = cornerOf(engine)?.getBoundingClientRect().width ?? DEFAULT_ROW_HEADER_WIDTH;
  const records = mountedRecords(engine, entering);
  const positioned = new Set<HTMLElement>();
  for (let x = 0; x < count; x++) {
    const offset = `${left}px`,
      header = engine.headers?.[x];
    if (header && header.style.left !== offset) header.style.left = offset;
    for (const cell of columnCells(records, x, positioned))
      if (cell.style.left !== offset) cell.style.left = offset;
    left += columnWidth(engine, x);
  }
}

/** Resolve hidden merged-over endpoints to the complete selected merge span. */
export function expandMergedSelection(range: CellSelection, sheet: Worksheet): CellSelection {
  const result = {
    x1: Math.min(range.x1, range.x2),
    y1: Math.min(range.y1, range.y2),
    x2: Math.max(range.x1, range.x2),
    y2: Math.max(range.y1, range.y2),
  };
  let changed = true;
  while (changed) {
    changed = false;
    for (const [address, span] of Object.entries(sheet.mergeCells ?? {})) {
      const start = parseSpreadsheetAddress(address);
      if (!start) continue;
      const right = start.x + span[0] - 1,
        bottom = start.y + span[1] - 1;
      if (start.x > result.x2 || right < result.x1 || start.y > result.y2 || bottom < result.y1)
        continue;
      const next = {
        x1: Math.min(result.x1, start.x),
        y1: Math.min(result.y1, start.y),
        x2: Math.max(result.x2, right),
        y2: Math.max(result.y2, bottom),
      };
      if (
        Object.keys(next).some(
          (key) => next[key as keyof CellSelection] !== result[key as keyof CellSelection],
        )
      ) {
        Object.assign(result, next);
        changed = true;
      }
    }
  }
  return result;
}

interface Box {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** The sticky corner, the first column header and the first attached row header. */
function headerRects(engine: JSpreadsheetWorksheetInstance) {
  const corner = cornerOf(engine)?.getBoundingClientRect();
  let column: DOMRect | undefined;
  for (const header of engine.headers ?? []) {
    const rect = header?.getBoundingClientRect();
    if (rect && rect.height > 0) {
      column = rect;
      break;
    }
  }
  const row = firstRowHeaderRect(engine);
  return { corner, column, row };
}

function firstRowHeaderRect(engine: JSpreadsheetWorksheetInstance) {
  const attached = gridRowWindow(engine);
  const last = attached?.end ?? engine.rows?.length ?? 0;
  for (let y = attached?.start ?? 0; y < last; y++) {
    const rect = engine.rows?.[y]?.element?.firstElementChild?.getBoundingClientRect();
    if (rect && rect.width > 0) return rect;
  }
  return undefined;
}

/** The visible cell body excludes both sticky header bands and scrollbar gutters. */
export function gridBodyBounds(engine: JSpreadsheetWorksheetInstance, viewport: HTMLElement) {
  const bounds = viewport.getBoundingClientRect();
  const origin = { left: bounds.left + viewport.clientLeft, top: bounds.top + viewport.clientTop };
  const right = origin.left + viewport.clientWidth,
    bottom = origin.top + viewport.clientHeight;
  const { corner, column, row } = headerRects(engine);
  const left = Math.max(origin.left, corner?.right ?? origin.left, row?.right ?? origin.left);
  const top = Math.max(origin.top, corner?.bottom ?? origin.top, column?.bottom ?? origin.top);
  return { origin, left: Math.min(right, left), top: Math.min(bottom, top), right, bottom };
}

/**
 * Place the fill handle at the selection's last cell as the engine does. The engine measures
 * that cell when the selection changes, which reads nothing while its row is detached.
 */
export function placeGridCorner(engine: JSpreadsheetWorksheetInstance) {
  const last = engine.highlighted?.at(-1)?.element as HTMLElement | undefined,
    corner = engine.corner,
    content = engine.content;
  if (!last?.isConnected || !corner || !content) return;
  const bounds = content.getBoundingClientRect(),
    rect = last.getBoundingClientRect();
  const top = `${rect.top - bounds.top + content.scrollTop + rect.height - 4}px`,
    left = `${rect.left - bounds.left + content.scrollLeft + rect.width - 4}px`;
  if (corner.style.top !== top) corner.style.top = top;
  if (corner.style.left !== left) corner.style.left = left;
}

/** CE's absolutely positioned fill handle must not paint or intercept clicks over headers. */
export function clipGridSelectionHandle(engine: JSpreadsheetWorksheetInstance, body: Box) {
  const handle = engine.corner;
  if (!handle) return;
  const rect = handle.getBoundingClientRect();
  const top = Math.max(0, body.top - rect.top),
    right = Math.max(0, rect.right - body.right),
    bottom = Math.max(0, rect.bottom - body.bottom),
    left = Math.max(0, body.left - rect.left);
  const clip =
    top + bottom >= rect.height || left + right >= rect.width
      ? 'inset(50%)'
      : top || right || bottom || left
        ? `inset(${top}px ${right}px ${bottom}px ${left}px)`
        : '';
  if (handle.style.clipPath !== clip) handle.style.clipPath = clip;
}

/** Where cells stop being covered on the left: past the frozen columns, else the row headers. */
function coveredLeft(engine: JSpreadsheetWorksheetInstance, bounds: DOMRect): number {
  const frozen = engine.options?.freezeColumns ?? 0;
  return frozen
    ? (engine.headers?.[frozen - 1]?.getBoundingClientRect().right ?? bounds.left)
    : (cornerOf(engine)?.getBoundingClientRect().right ?? bounds.left);
}

/** The scroll change that brings a span into view, or zero when it is already in view. */
function scrollDelta(span: { start: number; end: number }, view: { start: number; end: number }) {
  if (span.start < view.start) return span.start - view.start;
  if (span.end > view.end) return span.end - view.end;
  return 0;
}

export function scrollGridCellIntoView(
  engine: JSpreadsheetWorksheetInstance,
  x: number,
  y: number,
) {
  revealGridRows(engine, y);
  const viewport = gridViewport(engine),
    cell = gridCellAt(engine, x, y);
  if (!viewport || !cell) return;
  const rect = cell.getBoundingClientRect(),
    bounds = viewport.getBoundingClientRect();
  const headerBottom = cornerOf(engine)?.getBoundingClientRect().bottom ?? bounds.top;
  const left = coveredLeft(engine, bounds);
  if (x >= (engine.options?.freezeColumns ?? 0)) {
    const delta = scrollDelta(
      { start: rect.left, end: rect.right },
      { start: left, end: bounds.left + viewport.clientWidth },
    );
    if (delta) viewport.scrollLeft += delta;
  }
  const delta = scrollDelta(
    { start: rect.top, end: rect.bottom },
    { start: headerBottom, end: bounds.top + viewport.clientHeight },
  );
  if (delta) viewport.scrollTop += delta;
}
