import type { CSSProperties } from 'react';
import type { CellSelection } from '../model/editing.js';
import type { Worksheet } from '../model/workbook.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

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
  previousScale: number,
  nextScale: number,
) {
  if (!viewport) return;
  viewport.scrollLeft = (anchor.left * nextScale) / previousScale;
  viewport.scrollTop = (anchor.top * nextScale) / previousScale;
}

/** CE's freeze implementation rounds widths and offsets body cells relative to scrollLeft.
 * Keep its flags, but use measured, fractional sticky offsets for the rendered view.
 */
export function syncFrozenColumns(engine: JSpreadsheetWorksheetInstance) {
  const count = engine.options?.freezeColumns ?? 0;
  const table = engine.table;
  const corner = table?.tHead?.rows[0]?.cells[0];
  const rowHeaderWidth = corner?.getBoundingClientRect().width ?? DEFAULT_ROW_HEADER_WIDTH;
  if (corner) {
    corner.style.position = 'sticky';
    corner.style.left = '0px';
    corner.style.zIndex = '4';
  }
  for (const row of engine.rows ?? []) {
    const header = row?.element?.firstElementChild;
    if (header instanceof HTMLElement) {
      header.style.position = 'sticky';
      header.style.left = '0px';
      header.style.zIndex = '2';
    }
  }
  let left = rowHeaderWidth;
  const positioned = new Set<HTMLElement>();
  for (let x = 0; x < (engine.cols?.length ?? 0); x++) {
    const frozen = x < count;
    const header = engine.headers?.[x];
    if (header) {
      header.classList.toggle('jss_freezed', frozen);
      header.style.left = frozen ? `${left}px` : '';
      header.style.zIndex = frozen ? '3' : '';
    }
    for (const row of engine.records ?? []) {
      const cell = row?.[x]?.element;
      if (!(cell instanceof HTMLElement) || positioned.has(cell)) continue;
      positioned.add(cell);
      cell.classList.toggle('jss_freezed', frozen);
      cell.style.position = frozen ? 'sticky' : '';
      cell.style.left = frozen ? `${left}px` : '';
      cell.style.zIndex = frozen ? '1' : '';
    }
    const col = engine.cols?.[x]?.colElement;
    if (frozen && col?.style.display !== 'none')
      left += Number.parseFloat(col?.style.width || col?.getAttribute('width') || '0') || 0;
  }
}

/** Restore sticky offsets after CE's scroll handler writes relative-scroll offsets over them. */
export function syncFrozenColumnOffsets(engine: JSpreadsheetWorksheetInstance) {
  const count = engine.options?.freezeColumns ?? 0;
  if (!count) return;
  let left =
    engine.table?.tHead?.rows[0]?.cells[0]?.getBoundingClientRect().width ??
    DEFAULT_ROW_HEADER_WIDTH;
  const positioned = new Set<HTMLElement>();
  for (let x = 0; x < count; x++) {
    const offset = `${left}px`,
      header = engine.headers?.[x];
    if (header && header.style.left !== offset) header.style.left = offset;
    for (const row of engine.records ?? []) {
      const cell = row?.[x]?.element;
      if (!(cell instanceof HTMLElement) || positioned.has(cell)) continue;
      positioned.add(cell);
      if (cell.style.left !== offset) cell.style.left = offset;
    }
    const col = engine.cols?.[x]?.colElement;
    if (col?.style.display !== 'none')
      left += Number.parseFloat(col?.style.width || col?.getAttribute('width') || '0') || 0;
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

/** The visible cell body excludes both sticky header bands and scrollbar gutters. */
export function gridBodyBounds(engine: JSpreadsheetWorksheetInstance, viewport: HTMLElement) {
  const bounds = viewport.getBoundingClientRect();
  const origin = { left: bounds.left + viewport.clientLeft, top: bounds.top + viewport.clientTop };
  const right = origin.left + viewport.clientWidth,
    bottom = origin.top + viewport.clientHeight;
  const corner = engine.table?.tHead?.rows[0]?.cells[0]?.getBoundingClientRect();
  const column = engine.headers
    ?.find((header) => header?.getBoundingClientRect().height > 0)
    ?.getBoundingClientRect();
  const row = engine.rows
    ?.find((row) => (row?.element?.firstElementChild?.getBoundingClientRect().width ?? 0) > 0)
    ?.element.firstElementChild?.getBoundingClientRect();
  return {
    origin,
    left: Math.min(
      right,
      Math.max(origin.left, corner?.right ?? origin.left, row?.right ?? origin.left),
    ),
    top: Math.min(
      bottom,
      Math.max(origin.top, corner?.bottom ?? origin.top, column?.bottom ?? origin.top),
    ),
    right,
    bottom,
  };
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

/** CSS-pixel outline fragments, clipped behind headers and frozen cells.
 * Clipped edges carry no stroke: scrolling must not invent a selection edge.
 */
export function measureGridSelection(
  engine: JSpreadsheetWorksheetInstance,
  sheet: Worksheet,
  selection: CellSelection,
  viewport: HTMLElement,
): CSSProperties[] {
  const range = expandMergedSelection(selection, sheet);
  const body = gridBodyBounds(engine, viewport),
    { origin } = body;
  const viewportRight = body.right,
    viewportBottom = body.bottom,
    headerBottom = body.top,
    cornerRight = body.left;
  const frozenCount = sheet.freezeColumns ?? 0;
  const frozenRight = frozenCount
    ? Math.max(
        cornerRight,
        engine.headers?.[frozenCount - 1]?.getBoundingClientRect().right ?? cornerRight,
      )
    : cornerRight;
  const boxes: [Box | undefined, Box | undefined] = [undefined, undefined];
  const hasMerges = Object.entries(sheet.mergeCells ?? {}).some(([address, [width, height]]) => {
    const start = parseSpreadsheetAddress(address);
    return (
      start &&
      start.x <= range.x2 &&
      start.x + width > range.x1 &&
      start.y <= range.y2 &&
      start.y + height > range.y1
    );
  });
  const measured = new Map<HTMLElement, DOMRect | null>();
  function cellRect(x: number, y: number) {
    const cell = engine.getCellFromCoords?.(x, y);
    if (!cell || cell.style.display === 'none') return null;
    if (measured.has(cell)) return measured.get(cell)!;
    const rect = cell.getBoundingClientRect(),
      visible = rect.width > 0 && rect.height > 0 ? rect : null;
    measured.set(cell, visible);
    return visible;
  }
  function include(rect: DOMRect, part: 0 | 1) {
    const old = boxes[part];
    boxes[part] = old
      ? {
          left: Math.min(old.left, rect.left),
          top: Math.min(old.top, rect.top),
          right: Math.max(old.right, rect.right),
          bottom: Math.max(old.bottom, rect.bottom),
        }
      : { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
  }
  const lastY = Math.min(range.y2, (engine.records?.length ?? range.y2 + 1) - 1);
  const firstY = Math.max(0, range.y1),
    lastX = Math.min(range.x2, (engine.cols?.length ?? range.x2 + 1) - 1);
  for (const part of [0, 1] as const) {
    const firstX = Math.max(0, range.x1, part === 1 ? frozenCount : 0),
      endX = Math.min(lastX, part === 0 ? frozenCount - 1 : lastX);
    if (endX < firstX || lastY < firstY) continue;
    // Ordinary rectangular grids need only their corners, independent of selected area.
    const corners = hasMerges
      ? []
      : [
          [firstX, firstY],
          [endX, firstY],
          [firstX, lastY],
          [endX, lastY],
        ].map(([x, y]) => cellRect(x, y));
    if (corners.length && corners.every((rect): rect is DOMRect => rect !== null)) {
      for (const rect of corners) include(rect, part);
    } else
      for (let y = firstY; y <= lastY; y++)
        for (let x = firstX; x <= endX; x++) {
          const rect = cellRect(x, y);
          if (rect) include(rect, part);
        }
  }
  return boxes.flatMap((box, index) => {
    if (!box) return [];
    const left = Math.max(box.left, origin.left, index === 0 ? cornerRight : frozenRight),
      right = Math.min(box.right, viewportRight),
      top = Math.max(box.top, origin.top, headerBottom),
      bottom = Math.min(box.bottom, viewportBottom);
    if (right <= left || bottom <= top) return [];
    return [
      {
        position: 'absolute' as const,
        pointerEvents: 'none' as const,
        boxSizing: 'border-box' as const,
        left: left - origin.left,
        top: top - origin.top,
        width: right - left,
        height: bottom - top,
        ...(left > box.left + 0.1 || (index === 1 && boxes[0]) ? { borderLeftWidth: 0 } : {}),
        ...(right < box.right - 0.1 || (index === 0 && boxes[1]) ? { borderRightWidth: 0 } : {}),
        ...(top > box.top + 0.1 ? { borderTopWidth: 0 } : {}),
        ...(bottom < box.bottom - 0.1 ? { borderBottomWidth: 0 } : {}),
      },
    ];
  });
}

export function scrollGridCellIntoView(
  engine: JSpreadsheetWorksheetInstance,
  x: number,
  y: number,
) {
  const viewport = gridViewport(engine),
    cell = engine.getCellFromCoords?.(x, y);
  if (!viewport || !cell) return;
  const rect = cell.getBoundingClientRect(),
    bounds = viewport.getBoundingClientRect();
  const headerBottom =
    engine.table?.tHead?.rows[0]?.cells[0]?.getBoundingClientRect().bottom ?? bounds.top;
  const frozen = engine.options?.freezeColumns ?? 0;
  const left = frozen
    ? (engine.headers?.[frozen - 1]?.getBoundingClientRect().right ?? bounds.left)
    : (engine.table?.tHead?.rows[0]?.cells[0]?.getBoundingClientRect().right ?? bounds.left);
  if (x >= frozen) {
    if (rect.left < left) viewport.scrollLeft -= left - rect.left;
    else if (rect.right > bounds.left + viewport.clientWidth)
      viewport.scrollLeft += rect.right - bounds.left - viewport.clientWidth;
  }
  if (rect.top < headerBottom) viewport.scrollTop -= headerBottom - rect.top;
  else if (rect.bottom > bounds.top + viewport.clientHeight)
    viewport.scrollTop += rect.bottom - bounds.top - viewport.clientHeight;
}
