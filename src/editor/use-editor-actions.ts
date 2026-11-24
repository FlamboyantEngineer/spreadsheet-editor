'use client';

import type { CellSelection } from '../model/editing.js';
import type { WorkbookCommandResult } from '../model/operations.js';
import type { CellFormat } from '../model/review.js';
import type { StructureChange } from '../model/structure.js';
import type { CellInput, Workbook, Worksheet } from '../model/workbook.js';
import type { SpreadsheetMode } from './contracts.js';
import {
  type BorderKind,
  borderSelection,
  clearCells,
  formatSelection,
  historyStep,
  mergeCells,
  proposeCells,
  structureChange,
  updateSheet,
  type WorksheetEdit,
} from './editor-cell-commands.js';
import {
  copyCells,
  copyToClipboardEvent,
  pasteFromClipboard,
  pasteFromClipboardEvent,
  pasteValues,
} from './editor-clipboard.js';
import type { EditorCore, EditorDialog } from './editor-core.js';
import { exportExcel, importExcel, insertSelection, refreshData } from './editor-host-commands.js';
import {
  activateSheet,
  cancelDraftNow,
  commitDraftNow,
  followCollaborator,
  gridSelection,
  openPanel,
  restoreDraftAt,
  selectRange,
  setEditorMode,
  setHostPanel,
} from './editor-navigation.js';
import {
  applyDialog,
  dialogError,
  type FindState,
  navigateMatch,
  openDialog,
  renameSheet,
  replaceText,
  sheetCommand,
  validateSheetName,
} from './editor-sheet-dialogs.js';
import {
  applyFormula,
  attemptAction,
  commitChange,
  type CommitOptions,
  type DraftMutateOptions,
  mutateChange,
  type MutateOptions,
  mutateWithDraftChange,
  type ReplaceReason,
  replaceWorkbookChange,
  stampOf,
} from './editor-transactions.js';
import type { SheetCommand } from './sheet-commands.js';

/** The editor's workbook changes and cell commands, bound to this render's core. */
export function useEditorEditActions(core: EditorCore) {
  return {
    stamp: () => stampOf(core),
    commit: (next: Workbook, label = 'Edit workbook', options: CommitOptions = {}) =>
      commitChange(core, next, { ...options, label }),
    mutate: (next: Workbook, label: string, options: MutateOptions = {}) =>
      mutateChange(core, next, { ...options, label }),
    mutateWithDraft: (
      transform: (base: Workbook) => Workbook | WorkbookCommandResult,
      label: string,
      options: DraftMutateOptions = {},
    ) => mutateWithDraftChange(core, transform, { ...options, label }),
    replaceWorkbook: (next: Workbook, label: string, reason: ReplaceReason = 'import') =>
      replaceWorkbookChange(core, next, { label, reason }),
    attempt: (action: () => unknown) => attemptAction(core, action),
    applyFormula: () => applyFormula(core),
    cancelFormula: () => core.setDraft(null),
    setDraftValue: (value: string) => core.setDraft({ key: core.display.draftKey, value }),
    updateSheet: (next: Worksheet, change?: WorksheetEdit) => updateSheet(core, next, change),
    proposeCells: (values: CellInput[][], x: number, y: number) =>
      proposeCells(core, values, { x, y }),
    format: (properties: Record<string, string | null>, format?: CellFormat | null) =>
      formatSelection(core, properties, format),
    toggleFormat: (property: string, value: string) =>
      formatSelection(core, { [property]: core.display.style[property] === value ? null : value }),
    border: (kind: BorderKind) => borderSelection(core, kind),
    clear: (formatting = false) => clearCells(core, formatting),
    merge: (unmerge = false) => mergeCells(core, unmerge),
    structure: (axis: StructureChange['axis'], kind: StructureChange['kind'], after = false) =>
      structureChange(core, { axis, kind }, after),
    undo: () => historyStep(core, 'undo'),
    redo: () => historyStep(core, 'redo'),
    copy: (cut = false) => copyCells(core, cut),
    paste: () => pasteFromClipboard(core),
    transferCopy: (transfer: DataTransfer, cut: boolean) =>
      attemptAction(core, () => copyToClipboardEvent(core, transfer, cut)),
    transferPaste: (transfer: DataTransfer) =>
      attemptAction(core, () => pasteFromClipboardEvent(core, transfer)),
    pasteCells: (values: CellInput[][], x = core.selection.x1, y = core.selection.y1) =>
      pasteValues(core, values, { x, y }),
  };
}

/** Selection, worksheets, dialogs, panels and host commands, bound to this render's core. */
export function useEditorNavigationActions(core: EditorCore) {
  return {
    select: (range: CellSelection, targetSheetId?: string) =>
      selectRange(core, range, targetSheetId),
    activateSheet: (id: string) => activateSheet(core, id),
    onGridSelection: (range: CellSelection) => gridSelection(core, range),
    commitDraft: () => commitDraftNow(core),
    cancelDraft: () => cancelDraftNow(core),
    setMode: (next: SpreadsheetMode) => setEditorMode(core, next),
    restoreDraft: (value: { sheetId: string; x: number; y: number; text: string }) =>
      restoreDraftAt(core, value),
    followCollaborator: (id: string) => followCollaborator(core, id),
    publishPresence: (editing: boolean) => core.setEditingCell(editing),
    openPanel: (id: string) => openPanel(core, id),
    setHostPanel: (id: string | null) => setHostPanel(core, id),
    validateSheetName: (id: string, name: string) => validateSheetName(core, id, name),
    renameSheet: (id: string, name: string) => renameSheet(core, id, name),
    sheetCommand: (command: SheetCommand, targetSheetId = core.sheet.id) =>
      sheetCommand(core, command, targetSheetId),
    openDialog: (kind: Exclude<EditorDialog, null>) => openDialog(core, kind),
    dialogError: dialogError(core.dialog, core.dialogValue),
    applyDialog: () => applyDialog(core),
    insert: (kind: 'table' | 'spreadsheet') => insertSelection(core, kind),
    refresh: (selected = false) => refreshData(core, selected),
    importExcel: () => importExcel(core),
    exportExcel: () => exportExcel(core),
  };
}

/** Find navigation and replacement, bound to this render's core and find state. */
export function useEditorFindActions(core: EditorCore, find: FindState) {
  return {
    navigateMatch: (direction: number) => navigateMatch(core, find, direction),
    replace: (all: boolean, text: string) => replaceText(core, find, { all, text }),
  };
}
