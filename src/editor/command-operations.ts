import { orderedJson } from '../calculation/publication.js';
import {
  annotationOperation,
  collectCellOperations,
  collectLineGrowthOperations,
  geometryOperation,
  mergeOperation,
  sheetOperation,
  type WorkbookOperation,
} from '../model/operations.js';
import type { Workbook } from '../model/workbook.js';

/** Deletes, renames, geometry and merge changes of the worksheets `before` holds. */
function sheetChanges(before: Workbook, after: Workbook): WorkbookOperation[] {
  const operations: WorkbookOperation[] = [];
  for (const sheet of before.worksheets) {
    const next = after.worksheets.find((value) => value.id === sheet.id);
    if (!next) {
      operations.push(sheetOperation({ kind: 'sheet-delete', sheetId: sheet.id }, before, after));
      continue;
    }
    if (next.name !== sheet.name)
      operations.push(sheetOperation({ kind: 'sheet-rename', sheetId: sheet.id }, before, after));
    const geometry = geometryOperation(sheet, next);
    if (
      geometry.kind === 'geometry' &&
      orderedJson(geometry.before) !== orderedJson(geometry.after)
    )
      operations.push(geometry);
    if (orderedJson(sheet.mergeCells) !== orderedJson(next.mergeCells))
      operations.push(mergeOperation(sheet, next));
  }
  return operations;
}

/** Whether `after` holds exactly the worksheets of `before`, in another order. */
const reordered = (before: Workbook, after: Workbook) =>
  before.worksheets.length === after.worksheets.length &&
  before.worksheets.every((sheet) => after.worksheets.some((value) => value.id === sheet.id)) &&
  before.worksheets.some((sheet, index) => after.worksheets[index].id !== sheet.id);

/** Workbook-level changes: added sheets, a reorder, and the comments and suggestions. */
function workbookChanges(before: Workbook, after: Workbook): WorkbookOperation[] {
  const operations: WorkbookOperation[] = [];
  for (const sheet of after.worksheets)
    if (!before.worksheets.some((value) => value.id === sheet.id))
      operations.push(sheetOperation({ kind: 'sheet-add', sheetId: sheet.id }, before, after));
  if (reordered(before, after))
    operations.push(sheetOperation({ kind: 'sheet-reorder' }, before, after));
  if (
    orderedJson(before.comments) !== orderedJson(after.comments) ||
    orderedJson(before.suggestions) !== orderedJson(after.suggestions)
  )
    operations.push(annotationOperation(before, after));
  return operations;
}

/**
 * Sparse cell/layout patches complement commands which carry their own line operations.
 * `targets` limits the cell patches to the addresses the command wrote.
 */
export function commandOperations(
  before: Workbook,
  after: Workbook,
  targets?: ReadonlyMap<string, Iterable<string>>,
): WorkbookOperation[] {
  const operations: WorkbookOperation[] = [];
  for (const sheet of before.worksheets) {
    const next = after.worksheets.find((value) => value.id === sheet.id);
    if (next) operations.push(...collectLineGrowthOperations(sheet, next));
  }
  operations.push(...collectCellOperations(before, after, targets));
  if (before.name !== after.name)
    operations.push({
      kind: 'workbook-metadata',
      before: { name: before.name },
      after: { name: after.name },
    });
  operations.push(...sheetChanges(before, after), ...workbookChanges(before, after));
  return operations;
}
