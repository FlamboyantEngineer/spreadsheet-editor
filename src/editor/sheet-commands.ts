import type { CellSelection } from '../model/editing.js';
import { createWorkbookEntityId } from '../model/extent.js';
import { deleteWorksheet } from '../model/structure.js';
import { appendWorksheet, copyWorkbook, type Workbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';

export type SheetCommand = 'add' | 'insert' | 'duplicate' | 'delete' | 'left' | 'right';

/**
 * Applies a worksheet tab command to `base`. `activatedId` names the worksheet to show next: an
 * added one, or the neighbour of the deleted active worksheet.
 */
export function applySheetCommand(
  base: Workbook,
  command: SheetCommand,
  targetSheetId: string,
  activeSheetId: string,
  idFactory: SpreadsheetEditorProps['idFactory'],
): { workbook: Workbook; activatedId?: string } {
  let activatedId: string | undefined;
  let next = copyWorkbook(base);
  const targetSheet = next.worksheets.find((value) => value.id === targetSheetId)!;
  if (command === 'add' || command === 'insert' || command === 'duplicate') {
    next = appendWorksheet(base);
    const added = next.worksheets.at(-1)!;
    added.id = (idFactory ?? createWorkbookEntityId)('sheet');
    if (command === 'insert') {
      next.worksheets.pop();
      next.worksheets.splice(
        next.worksheets.findIndex((value) => value.id === targetSheetId),
        0,
        added,
      );
    }
    if (command === 'duplicate')
      next.worksheets[next.worksheets.length - 1] = {
        ...structuredClone(targetSheet),
        id: added.id,
        name: added.name,
      };
    activatedId = added.id;
  }
  if (command === 'delete') {
    const index = next.worksheets.findIndex((value) => value.id === targetSheetId);
    next = deleteWorksheet(base, targetSheetId);
    if (targetSheetId === activeSheetId)
      activatedId = next.worksheets[Math.min(index, next.worksheets.length - 1)].id;
  }
  if (command === 'left' || command === 'right') {
    const index = next.worksheets.findIndex((value) => value.id === targetSheetId),
      target =
        (index + (command === 'left' ? -1 : 1) + next.worksheets.length) % next.worksheets.length;
    [next.worksheets[index], next.worksheets[target]] = [
      next.worksheets[target],
      next.worksheets[index],
    ];
  }
  return { workbook: next, activatedId };
}

/** Sets the height of the selected rows or the width of the selected columns to `measure`. */
export function resizeSelection(
  base: Workbook,
  sheetId: string,
  selection: CellSelection,
  dialog: 'row-height' | 'column-width',
  measure: number,
): Workbook {
  const next = copyWorkbook(base),
    target = next.worksheets.find((value) => value.id === sheetId)!;
  if (dialog === 'row-height') {
    target.rows = { ...target.rows };
    for (let y = selection.y1; y <= selection.y2; y++) target.rows[y] = { height: measure };
  } else {
    target.columns = [...(target.columns ?? [])];
    while (target.columns.length <= selection.x2) target.columns.push({});
    for (let x = selection.x1; x <= selection.x2; x++)
      target.columns[x] = { ...target.columns[x], width: measure };
  }
  return next;
}
