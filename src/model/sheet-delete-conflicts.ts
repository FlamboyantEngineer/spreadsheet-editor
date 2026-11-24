import { formulaReferenceTokens } from './formula-reference-token.js';
import { WorkbookOperationConflict } from './operation-delta.js';
import type { WorkbookChangeContext } from './operations.js';
import type { SuggestedCellChange } from './review.js';
import { type CellInput, type Workbook, type Worksheet, worksheetNameKey } from './workbook.js';

type DeletedSheet = Pick<Worksheet, 'id' | 'name'>;

function namesDeletedSheet(value: CellInput, name: string): boolean {
  if (typeof value !== 'string' || !value.trimStart().startsWith('=')) return false;
  for (const token of formulaReferenceTokens(value))
    if (token.coordinates && token.sheetName !== null && worksheetNameKey(token.sheetName) === name)
      return true;
  return false;
}

function hasNewFormulas(baseline: Workbook, current: Workbook, deleted: DeletedSheet): boolean {
  const previous = new Map(baseline.worksheets.map((sheet) => [sheet.id, sheet]));
  const name = worksheetNameKey(deleted.name);
  return current.worksheets.some((sheet) => {
    if (sheet.id === deleted.id) return false;
    const prior = previous.get(sheet.id)?.data;
    return sheet.data.some(
      (row, y) =>
        row !== prior?.[y] &&
        row.some((value, x) => value !== prior?.[y]?.[x] && namesDeletedSheet(value, name)),
    );
  });
}

function hasNewComments(baseline: Workbook, current: Workbook, deleted: DeletedSheet): boolean {
  const prior = new Map((baseline.comments ?? []).map((thread) => [thread.id, thread]));
  return (current.comments ?? []).some((thread) => {
    if (thread.sheetId !== deleted.id || thread.orphaned) return false;
    const before = prior.get(thread.id);
    return before?.sheetId !== thread.sheetId || before.range !== thread.range || before.orphaned;
  });
}

const cellKey = ({ sheetId, address, before, after }: SuggestedCellChange) =>
  JSON.stringify([sheetId, address, before, after]);

function hasNewSuggestions(baseline: Workbook, current: Workbook, deleted: DeletedSheet): boolean {
  const prior = new Map(
    (baseline.suggestions ?? []).map((suggestion) => [suggestion.id, suggestion]),
  );
  const name = worksheetNameKey(deleted.name);
  return (current.suggestions ?? []).some((suggestion) => {
    const previous = prior.get(suggestion.id);
    if (suggestion === previous) return false;
    const cells = new Set(previous?.changes.map(cellKey));
    return suggestion.changes.some(
      (cell) =>
        !cells.has(cellKey(cell)) &&
        (cell.sheetId === deleted.id ||
          namesDeletedSheet(cell.before, name) ||
          namesDeletedSheet(cell.after, name)),
    );
  });
}

/** A stale sheet deletion can rebase independent edits but must acknowledge new references. */
export function checkSheetDeletions(
  current: Workbook,
  { baseline, operations }: WorkbookChangeContext,
) {
  if (current === baseline) return;
  for (const operation of operations) {
    if (operation.kind !== 'sheet-delete') continue;
    const deleted = operation.sheet;
    if (
      hasNewFormulas(baseline, current, deleted) ||
      hasNewComments(baseline, current, deleted) ||
      hasNewSuggestions(baseline, current, deleted)
    )
      throw new WorkbookOperationConflict('referential-delete-baseline');
  }
}
