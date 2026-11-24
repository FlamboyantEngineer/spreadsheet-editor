import { moveAddress, moveRange, rewriteStructureReferences } from './formula-references.js';
import { requireDefined } from './require-defined.js';
import type { SuggestedCellChange } from './review.js';
import { parseSpreadsheetRange } from './spreadsheet-addressing.js';
import type { StructureChange } from './structure.js';
import type { Workbook, Worksheet } from './workbook.js';

export interface ReviewReferenceMove {
  target: Pick<Worksheet, 'id' | 'name'>;
  change: StructureChange;
  sheetNames: Record<string, string>;
}

/** Captures the sheet names used by a line move without retaining its cell data. */
export function createReviewReferenceMove(
  workbook: Workbook,
  { id, name }: Pick<Worksheet, 'id' | 'name'>,
  change: StructureChange,
): ReviewReferenceMove {
  return {
    target: { id, name },
    change,
    sheetNames: Object.fromEntries(workbook.worksheets.map((sheet) => [sheet.id, sheet.name])),
  };
}

function moveSuggestedCell(
  cell: SuggestedCellChange,
  { target, change, sheetNames }: ReviewReferenceMove,
) {
  const moved = cell.sheetId === target.id ? moveAddress(cell.address, change) : cell.address;
  if (moved) cell.address = moved;
  const currentSheet = Object.hasOwn(sheetNames, cell.sheetId)
    ? sheetNames[cell.sheetId]
    : undefined;
  if (currentSheet)
    for (const field of ['before', 'after'] as const) {
      const value = cell[field];
      if (typeof value === 'string')
        cell[field] = rewriteStructureReferences(value, change, {
          currentSheet,
          targetSheet: target.name,
        });
    }
  return moved !== null;
}

/** Moves annotation anchors and proposed formulas in the same coordinate frame as authored cells. */
export function moveReviewReferences(
  review: Pick<Workbook, 'comments' | 'suggestions'>,
  move: ReviewReferenceMove,
): void {
  const { target, change } = move;
  for (const thread of review.comments ?? []) {
    if (thread.sheetId !== target.id || thread.orphaned) continue;
    const moved = moveRange(thread.range ?? '', change);
    if (moved) thread.range = requireDefined(parseSpreadsheetRange(moved)).address;
    else thread.orphaned = true;
  }
  for (const item of review.suggestions ?? [])
    for (const cell of item.changes) if (!moveSuggestedCell(cell, move)) item.outdated = true;
}
