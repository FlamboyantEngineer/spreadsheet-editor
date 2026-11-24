import type { CSSProperties } from 'react';
import type { CellSelection } from '../model/editing.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { Worksheet } from '../model/workbook.js';
import { gridCellAt } from './grid-cell.js';
import { expandMergedSelection, gridBodyBounds } from './grid-geometry.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { gridRowWindow } from './row-window.js';

interface Box {
  left: number;
  top: number;
  right: number;
  bottom: number;
}
type Part = 0 | 1;
type Body = ReturnType<typeof gridBodyBounds>;
/** The cells of a selection that can be seen: its attached rows and existing columns. */
interface SelectionSpan {
  firstY: number;
  lastY: number;
  lastX: number;
  frozenCount: number;
  hasMerges: boolean;
}

/** Whether a merge overlaps the range, so its corners alone cannot bound it. */
const overlapsMerge = (sheet: Worksheet, range: CellSelection) =>
  Object.entries(sheet.mergeCells ?? {}).some(([address, [width, height]]) => {
    const start = parseSpreadsheetAddress(address);
    return (
      start &&
      start.x <= range.x2 &&
      start.x + width > range.x1 &&
      start.y <= range.y2 &&
      start.y + height > range.y1
    );
  });

/** Measures each engine cell once; a hidden cell measures as null. */
function cellRects(engine: JSpreadsheetWorksheetInstance) {
  const measured = new Map<HTMLElement, DOMRect | null>();
  return (x: number, y: number): DOMRect | null => {
    const cell = gridCellAt(engine, x, y);
    if (!cell || cell.style.display === 'none') return null;
    // A hidden cell is measured as null, so presence is read with has().
    if (measured.has(cell)) return measured.get(cell) as DOMRect | null;
    const rect = cell.getBoundingClientRect(),
      visible = rect.width > 0 && rect.height > 0 ? rect : null;
    measured.set(cell, visible);
    return visible;
  };
}

const including = (box: Box | undefined, rect: DOMRect): Box =>
  box
    ? {
        left: Math.min(box.left, rect.left),
        top: Math.min(box.top, rect.top),
        right: Math.max(box.right, rect.right),
        bottom: Math.max(box.bottom, rect.bottom),
      }
    : { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };

function selectionSpan(
  engine: JSpreadsheetWorksheetInstance,
  sheet: Worksheet,
  range: CellSelection,
): SelectionSpan {
  // Detached rows lie outside the viewport; the attached rows bound what can be seen.
  const attached = gridRowWindow(engine);
  return {
    lastY: Math.min(
      range.y2,
      (engine.records?.length ?? range.y2 + 1) - 1,
      (attached?.end ?? Infinity) - 1,
    ),
    firstY: Math.max(0, range.y1, attached?.start ?? 0),
    lastX: Math.min(range.x2, (engine.cols?.length ?? range.x2 + 1) - 1),
    frozenCount: sheet.freezeColumns ?? 0,
    hasMerges: overlapsMerge(sheet, range),
  };
}

/** The bounding box of one part of the selection: the frozen columns (0) or the rest (1). */
function partBox(
  rectAt: ReturnType<typeof cellRects>,
  { span, range }: { span: SelectionSpan; range: CellSelection },
  part: Part,
): Box | undefined {
  const { firstY, lastY, lastX, frozenCount } = span;
  const firstX = Math.max(0, range.x1, part === 1 ? frozenCount : 0),
    endX = Math.min(lastX, part === 0 ? frozenCount - 1 : lastX);
  if (endX < firstX || lastY < firstY) return undefined;
  // Ordinary rectangular grids need only their corners, independent of selected area.
  const corners = span.hasMerges
    ? []
    : [
        [firstX, firstY],
        [endX, firstY],
        [firstX, lastY],
        [endX, lastY],
      ].map(([x, y]) => rectAt(x, y));
  if (corners.length && corners.every((rect): rect is DOMRect => rect !== null))
    return corners.reduce<Box | undefined>(including, undefined);
  let box: Box | undefined;
  for (let y = firstY; y <= lastY; y++)
    for (let x = firstX; x <= endX; x++) {
      const rect = rectAt(x, y);
      if (rect) box = including(box, rect);
    }
  return box;
}

/** Where cells stop being covered on the left, for each part. */
function coveredLefts(engine: JSpreadsheetWorksheetInstance, body: Body, frozenCount: number) {
  const cornerRight = body.left;
  const frozenRight = frozenCount
    ? Math.max(
        cornerRight,
        engine.headers?.[frozenCount - 1]?.getBoundingClientRect().right ?? cornerRight,
      )
    : cornerRight;
  return [cornerRight, frozenRight] as const;
}

/** Geometry shared by all outlines in one frame. */
export function gridSelectionGeometry(
  engine: JSpreadsheetWorksheetInstance,
  body: Body,
  frozenCount: number,
) {
  return { body, lefts: coveredLefts(engine, body, frozenCount), rectAt: cellRects(engine) };
}

/** One outline fragment, without a stroke on an edge that is clipped or meets the other part. */
function fragment(
  box: Box,
  { body, left: coveredLeft, joined }: { body: Body; left: number; joined: [boolean, boolean] },
): CSSProperties | undefined {
  const { origin } = body;
  const left = Math.max(box.left, origin.left, coveredLeft),
    right = Math.min(box.right, body.right),
    top = Math.max(box.top, origin.top, body.top),
    bottom = Math.min(box.bottom, body.bottom);
  if (right <= left || bottom <= top) return undefined;
  return {
    position: 'absolute' as const,
    pointerEvents: 'none' as const,
    boxSizing: 'border-box' as const,
    left: left - origin.left,
    top: top - origin.top,
    width: right - left,
    height: bottom - top,
    ...(left > box.left + 0.1 || joined[0] ? { borderLeftWidth: 0 } : {}),
    ...(right < box.right - 0.1 || joined[1] ? { borderRightWidth: 0 } : {}),
    ...(top > box.top + 0.1 ? { borderTopWidth: 0 } : {}),
    ...(bottom < box.bottom - 0.1 ? { borderBottomWidth: 0 } : {}),
  };
}

/** CSS-pixel outline fragments, clipped behind headers and frozen cells.
 * Clipped edges carry no stroke: scrolling must not invent a selection edge.
 */
export function measureGridSelection(
  engine: JSpreadsheetWorksheetInstance,
  {
    sheet,
    selection,
    geometry,
  }: {
    sheet: Worksheet;
    selection: CellSelection;
    geometry?: ReturnType<typeof gridSelectionGeometry>;
  },
  viewport: HTMLElement,
): CSSProperties[] {
  const range = expandMergedSelection(selection, sheet);
  const { body, lefts, rectAt } =
    geometry ??
    gridSelectionGeometry(engine, gridBodyBounds(engine, viewport), sheet.freezeColumns ?? 0);
  const span = selectionSpan(engine, sheet, range);
  const boxes = ([0, 1] as const).map((part) => partBox(rectAt, { span, range }, part));
  return boxes.flatMap((box, index) => {
    if (!box) return [];
    const joined: [boolean, boolean] = [index === 1 && !!boxes[0], index === 0 && !!boxes[1]];
    return fragment(box, { body, left: lefts[index], joined }) ?? [];
  });
}
