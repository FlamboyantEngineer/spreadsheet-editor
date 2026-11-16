import type { Dispatch, RefObject, SetStateAction } from 'react';
import type { WorkbookVersion } from '../host/adapters.js';
import type { SpreadsheetCopy } from '../model/clipboard.js';
import type { CellSelection } from '../model/editing.js';
import type { SpreadsheetCollaborator } from '../model/presence.js';
import type { ReviewActor } from '../model/review.js';
import type { Workbook, Worksheet } from '../model/workbook.js';
import type { GridController, SpreadsheetEditorProps, SpreadsheetMode } from './contracts.js';
import type { editorAccess } from './editor-access.js';
import type { EditorHistory } from './history.js';
import type { useCellDisplay } from './use-cell-display.js';

export type EditorDialog =
  'row-height' | 'column-width' | 'import' | 'export' | 'confirm-structure' | null;
export type DraftText = { key: string; value: string } | null;
export type NativeDraft = { x: number; y: number; text: string } | null;
export type ViewedVersion = { version: WorkbookVersion; content: Workbook } | null;
/** Where the reader left a worksheet, restored when they return to it. */
export interface SheetView {
  selection: CellSelection;
  viewport: { left: number; top: number };
}
export type LiveEditor = { canEdit: boolean; canWrite: boolean; props: SpreadsheetEditorProps };

/**
 * One render's editor: its props, workbook, place, permissions, drafts, and the refs and setters
 * its commands use. Commands are module functions of this core; the editor hook binds them.
 */
export interface EditorCore extends ReturnType<typeof editorAccess> {
  props: SpreadsheetEditorProps;
  workbook: Workbook;
  viewedVersion: ViewedVersion;
  currentHistory: EditorHistory;
  setHistory(next: EditorHistory): void;
  setLocal(next: Workbook): void;
  setLocalMode(mode: SpreadsheetMode): void;
  sheet: Worksheet;
  selection: CellSelection;
  setSelection: Dispatch<SetStateAction<CellSelection>>;
  setSheetId(id: string): void;
  setEditingCell(editing: boolean): void;
  sheetViews: RefObject<Map<string, SheetView>>;
  actor: ReviewActor;
  peers: SpreadsheetCollaborator[];
  grid: RefObject<GridController | null>;
  draftRef: RefObject<DraftText>;
  setDraft(next: DraftText): void;
  setNativeDraft(next: NativeDraft): void;
  display: ReturnType<typeof useCellDisplay>;
  alive: RefObject<boolean>;
  latestWorkbook: RefObject<Workbook>;
  live: RefObject<LiveEditor>;
  copied: RefObject<SpreadsheetCopy | null>;
  pendingStructure: RefObject<(() => void) | null>;
  dialog: EditorDialog;
  dialogValue: string;
  setDialog(next: EditorDialog): void;
  setDialogValue(next: string): void;
  setRefreshVersion: Dispatch<SetStateAction<number>>;
  setNotice(message: string | null): void;
  warn(message: string): void;
  setPanel(id: string | null): void;
  setPanelOpen(open: boolean): void;
}
