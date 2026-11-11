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
  const bounds = viewport.getBoundingClientRect();
  const origin = { left: bounds.left + viewport.clientLeft, top: bounds.top + viewport.clientTop };
  const viewportRight = origin.left + viewport.clientWidth,
    viewportBottom = origin.top + viewport.clientHeight;
  const table = engine.table;
  const headerBottom =
    table?.tHead?.rows[0]?.cells[0]?.getBoundingClientRect().bottom ?? origin.top;
  const cornerRight = table?.tHead?.rows[0]?.cells[0]?.getBoundingClientRect().right ?? origin.left;
  const frozenCount = sheet.freezeColumns ?? 0;
  const frozenRight = frozenCount
    ? Math.max(
        cornerRight,
        engine.headers?.[frozenCount - 1]?.getBoundingClientRect().right ?? cornerRight,
      )
    : cornerRight;
  const boxes: [Box | undefined, Box | undefined] = [undefined, undefined];
  const measured = new Set<HTMLElement>();
  const lastY = Math.min(range.y2, (engine.records?.length ?? range.y2 + 1) - 1);
  for (let y = Math.max(0, range.y1); y <= lastY; y++) {
    const lastX = Math.min(range.x2, (engine.cols?.length ?? range.x2 + 1) - 1);
    for (let x = Math.max(0, range.x1); x <= lastX; x++) {
      const cell = engine.getCellFromCoords?.(x, y);
      if (!cell || cell.style.display === 'none' || measured.has(cell)) continue;
      measured.add(cell);
      const rect = cell.getBoundingClientRect();
      if (rect.width <= 0 || rect.height <= 0) continue;
      const part = x < frozenCount ? 0 : 1;
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
