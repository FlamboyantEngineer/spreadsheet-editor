import type { Workbook } from '../model/workbook.js';
import {
  annotationOperation,
  collectCellOperations,
  collectLineGrowthOperations,
  geometryOperation,
  mergeOperation,
  sheetOperation,
  type WorkbookOperation,
} from '../model/operations.js';
import { orderedJson } from '../calculation/publication.js';

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
  for (const sheet of before.worksheets) {
    const next = after.worksheets.find((value) => value.id === sheet.id);
    if (!next) {
      operations.push(sheetOperation('sheet-delete', before, after, sheet.id));
      continue;
    }
    if (next.name !== sheet.name)
      operations.push(sheetOperation('sheet-rename', before, after, sheet.id));
    const geometry = geometryOperation(sheet, next);
    if (
      geometry.kind === 'geometry' &&
      orderedJson(geometry.before) !== orderedJson(geometry.after)
    )
      operations.push(geometry);
    if (orderedJson(sheet.mergeCells) !== orderedJson(next.mergeCells))
      operations.push(mergeOperation(sheet, next));
  }
  for (const sheet of after.worksheets)
    if (!before.worksheets.some((value) => value.id === sheet.id))
      operations.push(sheetOperation('sheet-add', before, after, sheet.id));
  if (
    before.worksheets.length === after.worksheets.length &&
    before.worksheets.every((sheet) => after.worksheets.some((value) => value.id === sheet.id)) &&
    before.worksheets.some((sheet, index) => after.worksheets[index].id !== sheet.id)
  )
    operations.push(sheetOperation('sheet-reorder', before, after));
  if (
    orderedJson(before.comments) !== orderedJson(after.comments) ||
    orderedJson(before.suggestions) !== orderedJson(after.suggestions)
  )
    operations.push(annotationOperation(before, after));
  return operations;
}
