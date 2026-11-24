import {
  type CellSelection,
  INITIAL_SELECTION,
  normalizeSelection,
  selectionFitsCellLimit,
} from '../model/editing.js';
import { gridWorksheetExtent, worksheetExtent } from '../model/extent.js';
import { collaboratorSheet } from '../model/presence.js';
import { copyWorkbook, type Workbook } from '../model/workbook.js';
import type { SpreadsheetMode } from './contracts.js';
import type { EditorCore } from './editor-core.js';
import { applyFormula, attemptAction } from './editor-transactions.js';

const sameSelection = (a: CellSelection, b: CellSelection) =>
  a.x1 === b.x1 && a.x2 === b.x2 && a.y1 === b.y1 && a.y2 === b.y2;

/** Go To and worksheet view restoration can only select a cell the grid can show. */
function canSelect(core: EditorCore, id: string, range: CellSelection) {
  const target = core.workbook.worksheets.find((sheet) => sheet.id === id);
  if (!target || !selectionFitsCellLimit(range)) return false;
  const [columns, rows] = gridWorksheetExtent(target, worksheetExtent(target, core.workbook));
  return range.x2 < columns && range.y2 < rows;
}

/** Keeps the active worksheet's selection and scroll position for when the reader returns. */
function rememberSheetView({ sheetViews, sheet, selection, grid }: EditorCore) {
  sheetViews.current.set(sheet.id, {
    selection: { ...selection },
    viewport: grid.current?.getViewport?.() ?? { left: 0, top: 0 },
  });
  grid.current?.releaseKeyboard();
}

/** Selects `range` on a worksheet, committing the active edit first. */
export function selectRange(core: EditorCore, range: CellSelection, targetSheetId = core.sheet.id) {
  const { sheet, grid } = core;
  const normalized = normalizeSelection(range);
  if (!canSelect(core, targetSheetId, normalized)) return false;
  grid.current?.commitDraft?.();
  if (!applyFormula(core)) return false;
  if (targetSheetId !== sheet.id) rememberSheetView(core);
  core.setSheetId(targetSheetId);
  core.setSelection(normalized);
  core.setDraft(null);
  if (targetSheetId === sheet.id) grid.current?.select(normalized);
  return true;
}

/** Shows worksheet `id` with the selection the reader left there. */
export function activateSheet(core: EditorCore, id: string) {
  if (id === core.sheet.id) return true;
  if (!core.workbook.worksheets.some((sheet) => sheet.id === id)) return false;
  core.grid.current?.commitDraft?.();
  if (!applyFormula(core)) return false;
  rememberSheetView(core);
  core.setSheetId(id);
  const saved = core.sheetViews.current.get(id)?.selection;
  core.setSelection(saved && canSelect(core, id, saved) ? saved : INITIAL_SELECTION);
  core.setDraft(null);
  return true;
}

export function commitDraftNow(core: EditorCore): Workbook {
  if (!core.alive.current) throw new Error('The editor session has ended.');
  core.grid.current?.commitDraft?.();
  if (!applyFormula(core)) throw new Error('The active cell edit could not be committed.');
  return copyWorkbook(core.latestWorkbook.current);
}

export function cancelDraftNow(core: EditorCore) {
  if (!core.alive.current) throw new Error('The editor session has ended.');
  core.grid.current?.cancelDraft?.();
  core.setNativeDraft(null);
  core.setDraft(null);
}

/** Follows a selection the grid made, unless the active edit cannot be committed. */
export function gridSelection(core: EditorCore, range: CellSelection) {
  attemptAction(core, () => {
    const normalized = normalizeSelection(range);
    if (!selectionFitsCellLimit(normalized)) {
      core.warn('Selection exceeds the cell limit.');
      core.grid.current?.select(core.selection);
      return;
    }
    if (sameSelection(core.selection, normalized)) return;
    if (!applyFormula(core)) {
      core.grid.current?.select(core.selection);
      return;
    }
    core.setSelection((previous) => (sameSelection(previous, normalized) ? previous : normalized));
    core.setDraft(null);
  });
}

export function setEditorMode(core: EditorCore, next: SpreadsheetMode) {
  if (core.locked || (!core.capabilities.edit && next !== core.mode)) return;
  if (!applyFormula(core)) return;
  core.setLocalMode(next);
  core.props.onModeChange?.(next);
  core.setDraft(null);
  core.grid.current?.releaseKeyboard();
}

const validDraftTarget = ({ x, y, text }: { x: number; y: number; text: string }) =>
  Number.isSafeInteger(x) &&
  Number.isSafeInteger(y) &&
  x >= 0 &&
  y >= 0 &&
  x <= 16383 &&
  y <= 1048575 &&
  text.length <= 10000;

/** Restores a host's saved draft into its cell (SE-HOST). */
export function restoreDraftAt(
  core: EditorCore,
  value: { sheetId: string; x: number; y: number; text: string },
) {
  if (!core.canWrite || !validDraftTarget(value))
    throw new Error('The draft cannot be restored in this session.');
  if (!selectRange(core, { x1: value.x, x2: value.x, y1: value.y, y2: value.y }, value.sheetId))
    throw new Error('The draft target is unavailable.');
  core.setDraft({ key: `${value.sheetId}:${value.x}:${value.y}`, value: value.text });
}

export function followCollaborator(core: EditorCore, id: string) {
  const person = core.peers.find((person) => person.id === id),
    target = person && collaboratorSheet(core.workbook, person);
  if (person?.selection && target) selectRange(core, person.selection, target.id);
}

export function openPanel(core: EditorCore, id: string) {
  core.setPanel(id);
  core.setPanelOpen(true);
  core.grid.current?.releaseKeyboard();
}

export function setHostPanel(core: EditorCore, id: string | null) {
  if (id === null) core.setPanelOpen(false);
  else openPanel(core, id);
}
