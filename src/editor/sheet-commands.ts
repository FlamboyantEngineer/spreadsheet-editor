import type { CellSelection } from '../model/editing.js';
import { createWorkbookEntityId } from '../model/extent.js';
import { requireDefined } from '../model/require-defined.js';
import { deleteWorksheet } from '../model/structure.js';
import { appendWorksheet, copyWorkbook, type Workbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';

export type SheetCommand = 'add' | 'insert' | 'duplicate' | 'delete' | 'left' | 'right';

/** The worksheet a tab command acts on, and the worksheet the editor shows. */
interface SheetCommandTarget {
  targetSheetId: string;
  activeSheetId: string;
  idFactory: SpreadsheetEditorProps['idFactory'];
}

interface SheetCommandResult {
  workbook: Workbook;
  activatedId?: string;
}

const indexOf = (workbook: Workbook, id: string) =>
  workbook.worksheets.findIndex((value) => value.id === id);

function addSheet(
  base: Workbook,
  command: 'add' | 'insert' | 'duplicate',
  { targetSheetId, idFactory }: SheetCommandTarget,
): SheetCommandResult {
  const targetSheet = requireDefined(base.worksheets.find((value) => value.id === targetSheetId));
  const next = appendWorksheet(base);
  const added = requireDefined(next.worksheets.at(-1));
  added.id = (idFactory ?? createWorkbookEntityId)('sheet');
  if (command === 'insert') {
    next.worksheets.pop();
    next.worksheets.splice(indexOf(next, targetSheetId), 0, added);
  }
  if (command === 'duplicate')
    next.worksheets[next.worksheets.length - 1] = {
      ...structuredClone(targetSheet),
      id: added.id,
      name: added.name,
    };
  return { workbook: next, activatedId: added.id };
}

function deleteSheet(
  base: Workbook,
  { targetSheetId, activeSheetId }: SheetCommandTarget,
): SheetCommandResult {
  const index = indexOf(base, targetSheetId);
  const next = deleteWorksheet(base, targetSheetId);
  if (targetSheetId !== activeSheetId) return { workbook: next };
  return {
    workbook: next,
    activatedId: next.worksheets[Math.min(index, next.worksheets.length - 1)].id,
  };
}

function moveSheet(base: Workbook, command: 'left' | 'right', targetSheetId: string) {
  const next = copyWorkbook(base);
  const index = indexOf(next, targetSheetId),
    target =
      (index + (command === 'left' ? -1 : 1) + next.worksheets.length) % next.worksheets.length;
  [next.worksheets[index], next.worksheets[target]] = [
    next.worksheets[target],
    next.worksheets[index],
  ];
  return { workbook: next };
}

/**
 * Applies a worksheet tab command to `base`. `activatedId` names the worksheet to show next: an
 * added one, or the neighbour of the deleted active worksheet.
 */
export function applySheetCommand(
  base: Workbook,
  command: SheetCommand,
  target: SheetCommandTarget,
): SheetCommandResult {
  requireDefined(base.worksheets.find((value) => value.id === target.targetSheetId));
  if (command === 'delete') return deleteSheet(base, target);
  if (command === 'left' || command === 'right')
    return moveSheet(base, command, target.targetSheetId);
  return addSheet(base, command, target);
}

/** A row height or column width applied to the selected lines. */
interface SelectionResize {
  dialog: 'row-height' | 'column-width';
  measure: number;
}

/** Sets the height of the selected rows or the width of the selected columns to `measure`. */
export function resizeSelection(
  base: Workbook,
  { sheetId, selection }: { sheetId: string; selection: CellSelection },
  { dialog, measure }: SelectionResize,
): Workbook {
  const next = copyWorkbook(base),
    target = requireDefined(next.worksheets.find((value) => value.id === sheetId));
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
