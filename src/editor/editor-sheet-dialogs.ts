import { calculationCellKey } from '../calculation/publication.js';
import { parseDelimited } from '../model/delimited.js';
import { INITIAL_SELECTION } from '../model/editing.js';
import { renameWorksheet, validateWorksheetName } from '../model/structure.js';
import { assertWorkbook, createWorkbook, type Workbook } from '../model/workbook.js';
import type { EditorCore, EditorDialog } from './editor-core.js';
import { selectRange } from './editor-navigation.js';
import { mutateWithDraftChange, replaceWorkbookChange } from './editor-transactions.js';
import { applySheetCommand, resizeSelection, type SheetCommand } from './sheet-commands.js';
import { type FindMatch, matchTargets, replaceMatches } from './use-find.js';

export function validateSheetName(core: EditorCore, id: string, name: string): string | null {
  try {
    validateWorksheetName(core.workbook, id, name);
    return null;
  } catch (error) {
    return error instanceof Error ? error.message : 'Choose a valid worksheet name.';
  }
}

const hasSheet = (core: EditorCore, id: string) =>
  core.workbook.worksheets.some((value) => value.id === id);

export function renameSheet(core: EditorCore, id: string, name: string): boolean {
  if (!core.canEdit || !hasSheet(core, id)) return false;
  return mutateWithDraftChange(core, (base) => renameWorksheet(base, id, name), {
    label: 'Rename worksheet',
  });
}

/** Runs a worksheet tab command, showing the worksheet it adds or leaves active. */
export function sheetCommand(core: EditorCore, command: SheetCommand, targetSheetId: string) {
  if (!core.canEdit || !hasSheet(core, targetSheetId)) return false;
  let activatedId: string | undefined;
  const changed = mutateWithDraftChange(
    core,
    (base) => {
      const result = applySheetCommand(base, command, {
        targetSheetId,
        activeSheetId: core.sheet.id,
        idFactory: core.props.idFactory,
      });
      activatedId = result.activatedId;
      return result.workbook;
    },
    { label: `${command} worksheet` },
  );
  if (changed && activatedId) {
    core.setSheetId(activatedId);
    core.setSelection(INITIAL_SELECTION);
  }
  return changed;
}

/** The value a dialog opens with: the current measure, or the workbook for export. */
function dialogInitialValue({ sheet, selection, workbook }: EditorCore, kind: EditorDialog) {
  if (kind === 'row-height') return String(sheet.rows?.[selection.y1]?.height ?? 26);
  if (kind === 'column-width') return String(sheet.columns?.[selection.x1]?.width ?? 100);
  return kind === 'export' ? JSON.stringify(workbook, null, 2) : '';
}

export function openDialog(core: EditorCore, kind: Exclude<EditorDialog, null>) {
  core.setDialog(kind);
  core.setDialogValue(dialogInitialValue(core, kind));
  core.grid.current?.releaseKeyboard();
}

const invalidMeasure = (value: string) =>
  !Number.isFinite(Number(value)) || Number(value) <= 0 || Number(value) > 2000;

/** Why the open dialog's value cannot be applied, or null. */
export function dialogError(dialog: EditorDialog, value: string) {
  return (dialog === 'row-height' || dialog === 'column-width') && invalidMeasure(value)
    ? 'Enter a positive size up to 2000 pixels.'
    : null;
}

function importDialog(core: EditorCore) {
  const { dialogValue } = core;
  const value = dialogValue.trimStart().startsWith('{')
    ? (JSON.parse(dialogValue) as Workbook)
    : createWorkbook(parseDelimited(dialogValue, ','));
  assertWorkbook(value);
  replaceWorkbookChange(core, value, { label: 'Import workbook', reason: 'import' });
}

function resizeDialog(core: EditorCore, dialog: 'row-height' | 'column-width') {
  const measure = Number(core.dialogValue);
  if (!Number.isFinite(measure) || measure <= 0 || measure > 2000)
    throw new RangeError('Enter a positive size up to 2000 pixels.');
  mutateWithDraftChange(
    core,
    (base) =>
      resizeSelection(
        base,
        { sheetId: core.sheet.id, selection: core.selection, idFactory: core.props.idFactory },
        { dialog, measure },
      ),
    { label: dialog === 'row-height' ? 'Set row height' : 'Set column width' },
  );
}

export function applyDialog(core: EditorCore) {
  const { dialog } = core;
  if (dialogError(dialog, core.dialogValue)) return;
  if (dialog === 'confirm-structure') {
    core.pendingStructure.current?.();
    core.pendingStructure.current = null;
    core.setDialog(null);
    return;
  }
  if (dialog === 'export') {
    core.setDialog(null);
    return;
  }
  if (!core.canEdit) return;
  if (dialog === 'import') importDialog(core);
  if (dialog === 'row-height' || dialog === 'column-width') resizeDialog(core, dialog);
  core.setDialog(null);
}

/** The find pane's matches, the active one, and the cell key a navigation makes active. */
export interface FindState {
  query: string;
  matches: FindMatch[];
  matchIndex: number;
  setActiveMatch(key: string | null): void;
}

export function navigateMatch(core: EditorCore, find: FindState, direction: number) {
  const { matches, matchIndex } = find;
  if (!matches.length) return;
  const current = matchIndex >= 0 ? matchIndex : direction < 0 ? 0 : -1;
  const index = (current + direction + matches.length) % matches.length,
    match = matches[index];
  if (selectRange(core, { x1: match.x, x2: match.x, y1: match.y, y2: match.y }, match.sheetId))
    find.setActiveMatch(calculationCellKey(match.sheetId, match.x, match.y));
}

/** Replaces the active match, or every match, in authored cells. Returns how many it replaced. */
export function replaceText(
  core: EditorCore,
  find: FindState,
  { all, text }: { all: boolean; text: string },
) {
  const { query, matches, matchIndex } = find;
  if (!core.canWrite || !query) return;
  const targets = (
    all ? matches : matches[matchIndex] ? [matches[matchIndex]] : matches.slice(0, 1)
  ).filter((match) => match.authored);
  if (!targets.length) return 0;
  mutateWithDraftChange(core, (base) => replaceMatches(base, targets, { query, text, all }), {
    label: all ? 'Replace all matches' : 'Replace match',
    written: matchTargets(targets),
  });
  return targets.length;
}
