import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

type Side = 'Top' | 'Right' | 'Bottom' | 'Left';
const hidden = new WeakMap<HTMLElement, Partial<Record<Side, string>>>();
const hiddenCells = new WeakMap<JSpreadsheetWorksheetInstance, Set<HTMLElement>>();

/** Restore only our presentation colors before the next authored-style reconciliation. */
export function restoreSharedGridBorders(engine: JSpreadsheetWorksheetInstance) {
  for (const cell of hiddenCells.get(engine) ?? []) {
    const colors = hidden.get(cell);
    if (!colors) continue;
    for (const side of SIDES)
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

const SIDES = ['Top', 'Right', 'Bottom', 'Left'] as const;
type Seam = 'Left' | 'Top';
const OPPOSITE = { Left: 'Right', Top: 'Bottom' } as const;

const transparent = (color: string) =>
  !color || color === 'transparent' || /^rgba\([^)]*,\s*0\s*\)$/.test(color);

function visibleSides(style: CSSStyleDeclaration): Record<Side, boolean> {
  const visible = {} as Record<Side, boolean>;
  for (const side of SIDES)
    visible[side] =
      style[`border${side}Style`] !== 'none' &&
      Number.parseFloat(style[`border${side}Width`]) > 0 &&
      !transparent(style[`border${side}Color`]);
  return visible;
}

/** A cell whose borders can be measured: attached, shown, not editing, at grid coordinates. */
const measurable = (cell: HTMLElement, x: number, y: number) =>
  // A detached row has no computed style; the row window reconciles it once attached.
  cell.isConnected &&
  cell.style.display !== 'none' &&
  !cell.classList.contains('editor') &&
  Number.isInteger(x) &&
  Number.isInteger(y) &&
  x >= 0 &&
  y >= 0;

/** One reconciliation of a grid's shared seams, measuring each cell at most once. */
class BorderPass {
  private readonly cells = new Map<HTMLElement, EdgeCell | null>();
  readonly changed = new Set<HTMLElement>();

  constructor(private readonly engine: JSpreadsheetWorksheetInstance) {}

  edge(cell: HTMLElement | null | undefined): EdgeCell | undefined {
    if (!cell) return undefined;
    const known = this.cells.get(cell);
    if (known !== undefined) return known ?? undefined;
    const x = Number(cell.dataset.x),
      y = Number(cell.dataset.y);
    if (!measurable(cell, x, y)) {
      this.cells.set(cell, null);
      return undefined;
    }
    const entry = {
      cell,
      x,
      y,
      right: x + (Number(cell.getAttribute('colspan')) || 1),
      bottom: y + (Number(cell.getAttribute('rowspan')) || 1),
      visible: visibleSides(getComputedStyle(cell)),
    };
    this.cells.set(cell, entry);
    return entry;
  }

  at(x: number, y: number): EdgeCell | undefined {
    return this.edge(this.engine.getCellFromCoords?.(x, y) as HTMLElement | null | undefined);
  }

  /** Cells with an edge to reconcile, and their right and bottom neighbours. */
  visited(): Set<EdgeCell> {
    const visit = new Set<EdgeCell>();
    const add = (cell: EdgeCell | undefined) => {
      if (cell) visit.add(cell);
    };
    for (const cell of this.engine.table?.querySelectorAll<HTMLElement>(EDGE_CELLS) ?? []) {
      const current = this.edge(cell);
      if (!current) continue;
      visit.add(current);
      for (let y = current.y; y < current.bottom; y++) add(this.at(current.right, y));
      for (let x = current.x; x < current.right; x++) add(this.at(x, current.bottom));
    }
    return visit;
  }

  /** The cells across a cell's left and top seams, left ones first. */
  *seams(current: EdgeCell): Generator<[EdgeCell | undefined, Seam]> {
    for (let y = current.y; y < current.bottom && current.x > 0; y++)
      yield [this.at(current.x - 1, y), 'Left'];
    for (let x = current.x; x < current.right && current.y > 0; x++)
      yield [this.at(x, current.y - 1), 'Top'];
  }

  hide(cell: HTMLElement, side: Side) {
    const colors = hidden.get(cell) ?? {};
    colors[side] ??= cell.style[`border${side}Color`];
    hidden.set(cell, colors);
    this.changed.add(cell);
    cell.style[`border${side}Color`] = 'transparent';
  }

  /** A clipboard edge wins its seam; returns whether one of the two cells had one. */
  private joinCopies(current: EdgeCell, neighbor: EdgeCell, side: Seam): boolean {
    const other = OPPOSITE[side];
    const ownCopy = current.cell.classList.contains(`copying-${side.toLowerCase()}`),
      otherCopy = neighbor.cell.classList.contains(`copying-${other.toLowerCase()}`);
    if (ownCopy && !otherCopy) this.hide(neighbor.cell, other);
    if (otherCopy && !ownCopy) this.hide(current.cell, side);
    return ownCopy || otherCopy;
  }

  join(current: EdgeCell, neighbor: EdgeCell | undefined, side: Seam) {
    const other = OPPOSITE[side];
    if (!neighbor || neighbor === current || !current.visible[side] || !neighbor.visible[other])
      return;
    if (this.joinCopies(current, neighbor, side)) return;
    const start = side === 'Left' ? 'y' : 'x',
      end = side === 'Left' ? 'bottom' : 'right';
    if (current[start] >= neighbor[start] && current[end] <= neighbor[end])
      this.hide(current.cell, side);
    else if (neighbor[start] >= current[start] && neighbor[end] <= current[end])
      this.hide(neighbor.cell, other);
  }
}

/** Separate-border tables reserve both sides of a seam. Hide one duplicate paint
 * channel rather than changing layout width or the canonical authored border.
 * A longer merged edge wins over a shorter neighbor, retaining its continuous edge.
 * Only cells with an inline border or a clipboard edge, and their right and bottom
 * neighbors, are visited.
 */
export function syncSharedGridBorders(engine: JSpreadsheetWorksheetInstance) {
  restoreSharedGridBorders(engine);
  const pass = new BorderPass(engine);
  const visit = pass.visited();
  hiddenCells.set(engine, pass.changed);
  // Measure every visited cell before hiding any edge, as one complete pass did.
  for (const current of visit) for (const _seam of pass.seams(current));
  for (const current of visit)
    for (const [neighbor, side] of pass.seams(current)) pass.join(current, neighbor, side);
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
