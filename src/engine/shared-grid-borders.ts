import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

type Side = 'Top' | 'Right' | 'Bottom' | 'Left';
const hidden = new WeakMap<HTMLElement, Partial<Record<Side, string>>>();
const hiddenCells = new WeakMap<JSpreadsheetWorksheetInstance, Set<HTMLElement>>();

/** Restore only our presentation colors before the next authored-style reconciliation. */
export function restoreSharedGridBorders(engine: JSpreadsheetWorksheetInstance) {
  for (const cell of hiddenCells.get(engine) ?? []) {
    const colors = hidden.get(cell);
    if (!colors) continue;
    for (const side of ['Top', 'Right', 'Bottom', 'Left'] as const)
      if (colors[side] !== undefined) cell.style[`border${side}Color`] = colors[side];
    hidden.delete(cell);
  }
  hiddenCells.delete(engine);
}

interface EdgeCell {
  cell: HTMLElement;
  x: number;
  y: number;
  right: number;
  bottom: number;
  visible: Record<Side, boolean>;
}

/** Cells that can paint a right or bottom edge: the resting grid paints only top and
 * left seams, so a duplicate seam needs an inline (authored) border or a clipboard edge.
 */
const EDGE_CELLS =
  'tbody td[data-x][data-y]:is([style*="border"], .copying-top, .copying-right, .copying-bottom, .copying-left, .selection-top, .selection-right, .selection-bottom, .selection-left)';

/** Separate-border tables reserve both sides of a seam. Hide one duplicate paint
 * channel rather than changing layout width or the canonical authored border.
 * A longer merged edge wins over a shorter neighbor, retaining its continuous edge.
 * Only cells with an inline border or a clipboard edge, and their right and bottom
 * neighbors, are visited.
 */
export function syncSharedGridBorders(engine: JSpreadsheetWorksheetInstance) {
  restoreSharedGridBorders(engine);
  const cells = new Map<HTMLElement, EdgeCell | null>();
  function edge(cell: HTMLElement | undefined): EdgeCell | undefined {
    if (!cell) return undefined;
    const known = cells.get(cell);
    if (known !== undefined) return known ?? undefined;
    const x = Number(cell.dataset.x),
      y = Number(cell.dataset.y);
    if (
      cell.style.display === 'none' ||
      cell.classList.contains('editor') ||
      !Number.isInteger(x) ||
      !Number.isInteger(y) ||
      x < 0 ||
      y < 0
    ) {
      cells.set(cell, null);
      return undefined;
    }
    const style = getComputedStyle(cell);
    const visible = {} as Record<Side, boolean>;
    for (const side of ['Top', 'Right', 'Bottom', 'Left'] as const) {
      const color = style[`border${side}Color`];
      visible[side] =
        style[`border${side}Style`] !== 'none' &&
        Number.parseFloat(style[`border${side}Width`]) > 0 &&
        !!color &&
        color !== 'transparent' &&
        !/^rgba\([^)]*,\s*0\s*\)$/.test(color);
    }
    const entry = {
      cell,
      x,
      y,
      right: x + (Number(cell.getAttribute('colspan')) || 1),
      bottom: y + (Number(cell.getAttribute('rowspan')) || 1),
      visible,
    };
    cells.set(cell, entry);
    return entry;
  }
  const at = (x: number, y: number) =>
    engine.getCellFromCoords?.(x, y) as HTMLElement | null | undefined;
  const visit = new Set<EdgeCell>();
  for (const cell of engine.table?.querySelectorAll<HTMLElement>(EDGE_CELLS) ?? []) {
    const current = edge(cell);
    if (!current) continue;
    visit.add(current);
    for (let y = current.y; y < current.bottom; y++) {
      const neighbor = edge(at(current.right, y) ?? undefined);
      if (neighbor) visit.add(neighbor);
    }
    for (let x = current.x; x < current.right; x++) {
      const neighbor = edge(at(x, current.bottom) ?? undefined);
      if (neighbor) visit.add(neighbor);
    }
  }
  const changed = new Set<HTMLElement>();
  hiddenCells.set(engine, changed);
  function hide(cell: HTMLElement, side: Side) {
    const colors = hidden.get(cell) ?? {};
    colors[side] ??= cell.style[`border${side}Color`];
    hidden.set(cell, colors);
    changed.add(cell);
    cell.style[`border${side}Color`] = 'transparent';
  }
  function join(current: EdgeCell, neighbor: EdgeCell | undefined, side: 'Left' | 'Top') {
    const other = side === 'Left' ? 'Right' : 'Bottom';
    if (!neighbor || neighbor === current || !current.visible[side] || !neighbor.visible[other])
      return;
    const ownCopy = current.cell.classList.contains(`copying-${side.toLowerCase()}`),
      otherCopy = neighbor.cell.classList.contains(`copying-${other.toLowerCase()}`);
    if (ownCopy && otherCopy) return;
    if (ownCopy) {
      hide(neighbor.cell, other);
      return;
    }
    if (otherCopy) {
      hide(current.cell, side);
      return;
    }
    const start = side === 'Left' ? 'y' : 'x',
      end = side === 'Left' ? 'bottom' : 'right';
    if (current[start] >= neighbor[start] && current[end] <= neighbor[end])
      hide(current.cell, side);
    else if (neighbor[start] >= current[start] && neighbor[end] <= current[end])
      hide(neighbor.cell, other);
  }
  // Measure every visited cell before hiding any edge, as one complete pass did.
  for (const current of visit) {
    for (let y = current.y; y < current.bottom && current.x > 0; y++)
      edge(at(current.x - 1, y) ?? undefined);
    for (let x = current.x; x < current.right && current.y > 0; x++)
      edge(at(x, current.y - 1) ?? undefined);
  }
  for (const current of visit) {
    for (let y = current.y; y < current.bottom && current.x > 0; y++)
      join(current, edge(at(current.x - 1, y) ?? undefined), 'Left');
    for (let x = current.x; x < current.right && current.y > 0; x++)
      join(current, edge(at(x, current.y - 1) ?? undefined), 'Top');
  }
}

/** Native clipboard outlines can appear without a React render. Reconcile only
 * clipboard-class transitions, leaving ordinary selection changes inexpensive.
 */
export function observeGridBorderCopies(engine: JSpreadsheetWorksheetInstance): () => void {
  if (!engine.table) return () => {};
  const observer = new MutationObserver((records) => {
    if (
      records.some(
        (record) =>
          /copying-/.test(record.oldValue ?? '') ||
          (record.target instanceof HTMLElement && /copying-/.test(record.target.className)),
      )
    )
      syncSharedGridBorders(engine);
  });
  observer.observe(engine.table, {
    subtree: true,
    attributes: true,
    attributeFilter: ['class'],
    attributeOldValue: true,
  });
  return () => observer.disconnect();
}
