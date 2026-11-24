import { writeAuthoredCellState } from './authored-cell.js';
import type { WorkbookIdFactory } from './extent.js';
import type { WorkbookOperation } from './operations.js';
import { requireDefined } from './require-defined.js';
import type { SpreadsheetSuggestion, SuggestedCellChange } from './review.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import { changeWorkbookStructureWithOperations } from './structure-operations.js';
import type { Workbook, Worksheet } from './workbook.js';

type Structure = NonNullable<SpreadsheetSuggestion['structure']>;
type Point = { x: number; y: number };

const json = (value: unknown) => JSON.stringify(value);

/** Threads and suggestions on sheets a structural suggestion removed become orphaned or outdated. */
function markRemovedSheets(next: Workbook) {
  const ids = new Set(next.worksheets.map((sheet) => sheet.id));
  next.comments?.forEach((thread) => {
    if (thread.scope !== 'workbook' && !ids.has(thread.sheetId ?? '')) thread.orphaned = true;
  });
  next.suggestions?.forEach((item) => {
    if (item.changes.some((cell) => !ids.has(cell.sheetId))) item.outdated = true;
  });
}

function acceptStructure(next: Workbook, structure: Structure, idFactory?: WorkbookIdFactory) {
  if (json(next.worksheets) !== json(structure.before))
    throw new Error('The workbook changed since this structural suggestion.');
  if (structure.operation) {
    const { sheetId, change } = structure.operation;
    return changeWorkbookStructureWithOperations(
      next,
      { sheetId, idFactory, materialize: false },
      change,
    );
  }
  next.worksheets = structuredClone(structure.after);
  markRemovedSheets(next);
  return { workbook: next };
}

/** Whether a cell still holds the state the suggestion saw. */
function unchanged(sheet: Worksheet | undefined, point: Point | null, change: SuggestedCellChange) {
  return (
    !!sheet &&
    !!point &&
    (sheet.data[point.y]?.[point.x] ?? null) === change.before &&
    sheet.style?.[change.address] === change.beforeStyle &&
    json(sheet.formats?.[change.address]) === json(change.beforeFormat) &&
    json(sheet.inputValues?.[change.address]) === json(change.beforeInputValue)
  );
}

export function writeSuggestedCell(sheet: Worksheet, change: SuggestedCellChange) {
  writeAuthoredCellState(sheet, change.address, {
    value: change.after,
    inputValue: change.afterInputValue,
    style: change.afterStyle,
    format: change.afterFormat,
    provenance: sheet.cellProvenance?.[change.address],
  });
}

/** Applies a cell suggestion only when none of its cells changed since it was made. */
function acceptCells(next: Workbook, changes: SuggestedCellChange[]): Workbook {
  for (const change of changes) {
    const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId);
    if (!unchanged(sheet, parseSpreadsheetAddress(change.address), change))
      throw new Error(
        `Cell ${change.address} changed since this suggestion. It has not been overwritten.`,
      );
  }
  for (const change of changes)
    writeSuggestedCell(
      requireDefined(next.worksheets.find((sheet) => sheet.id === change.sheetId)),
      change,
    );
  return next;
}

export function acceptSuggestion(
  next: Workbook,
  suggestion: SpreadsheetSuggestion,
  idFactory?: WorkbookIdFactory,
): { workbook: Workbook; operations?: WorkbookOperation[] } {
  if (suggestion.outdated)
    throw new Error('The suggested cells were deleted. Reject this outdated suggestion.');
  return suggestion.structure
    ? acceptStructure(next, suggestion.structure, idFactory)
    : { workbook: acceptCells(next, suggestion.changes) };
}

/** Accepted values are the suggestion author's, by provenance. */
export function recordProvenance(next: Workbook, suggestion: SpreadsheetSuggestion) {
  for (const change of suggestion.changes) {
    if (
      change.before === change.after &&
      json(change.beforeInputValue) === json(change.afterInputValue)
    )
      continue;
    const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId);
    if (sheet)
      (sheet.cellProvenance ??= {})[change.address] = {
        actor: { ...suggestion.author },
        source: 'human',
        at: suggestion.updatedAt ?? suggestion.createdAt,
      };
  }
}
