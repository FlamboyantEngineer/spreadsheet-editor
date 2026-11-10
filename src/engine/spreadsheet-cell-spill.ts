/**
 * SE-CR-2.2a: how far each cell's text may spill to its right, as in Excel. Text paints over the
 * empty cells after it and stops at the next non-empty cell or at the worksheet's last column.
 */
export interface SpreadsheetSpillCell {
  /** The cell takes part in the row's layout; hidden and merged-over cells do not. */
  readonly rendered: boolean;
  /** The cell shows nothing, as CSS `:empty` sees it. */
  readonly empty: boolean;
  /** The cell carries a class other than the read-only marker, so it keeps its opaque layer. */
  readonly decorated: boolean;
  /** The cell's text may paint past its edge: plain text that is not frozen, merged or edited. */
  readonly spillable: boolean;
}

/** Each cell's spill width: the combined width of the undecorated empty cells that follow it. */
export function resolveSpreadsheetRowSpillWidths(
  cells: ReadonlyArray<SpreadsheetSpillCell>,
  columnWidths: ReadonlyArray<number>,
): number[] {
  const spillWidths = cells.map(() => 0);
  let emptyWidth = 0;
  for (let x = cells.length - 1; x >= 0; x -= 1) {
    const cell = cells[x];
    if (!cell?.rendered) continue;
    if (cell.empty && !cell.decorated) {
      emptyWidth += Math.max(0, columnWidths[x] ?? 0);
      continue;
    }
    if (!cell.empty && cell.spillable) spillWidths[x] = emptyWidth;
    emptyWidth = 0;
  }
  return spillWidths;
}
