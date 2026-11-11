import type { ReactNode } from 'react';
import type { CalculationProvider } from '../calculation/provider.js';
import type { DisplayOptions } from '../calculation/format.js';
import type {
  CalculationOptions,
  CalculationResult,
  FormulaHost,
} from '../calculation/evaluator.js';
import type { Workbook } from '../model/workbook.js';
import type { ReviewActor } from '../model/review.js';
import type { DocumentInsertion } from '../document/insertion.js';
import type { SpreadsheetCollaborator, SpreadsheetPresence } from '../model/presence.js';
import type { CellSelection } from '../model/editing.js';

export type SpreadsheetChrome = 'minimal' | 'embedded' | 'full';
export type SpreadsheetMode = 'author' | 'review' | 'read';
export type CalculationMode =
  | { mode: 'client'; host?: FormulaHost; options?: CalculationOptions }
  | { mode: 'server'; provider: CalculationProvider; options?: CalculationOptions };
export interface SpreadsheetClipboard {
  readText(): Promise<string>;
  writeText(text: string): Promise<void>;
}
export interface SpreadsheetPermissions {
  edit?: boolean;
  comment?: boolean;
  suggest?: boolean;
  decideSuggestions?: boolean;
}
export interface SpreadsheetHostPanel {
  id: string;
  label: string;
  content: ReactNode;
}
export interface SpreadsheetSaveContext {
  signal: AbortSignal;
  kind: 'manual' | 'autosave';
  requestId: string;
}
export interface SpreadsheetHistoryEntry {
  id: string;
  timestamp: string;
  author: ReviewActor;
  state?: string;
  revision?: string;
}
export interface SpreadsheetHistory {
  entries: readonly SpreadsheetHistoryEntry[];
  loading?: boolean;
  error?: string;
  onRefresh?(target?: { sheetId: string; range: string }): void | Promise<void>;
  onOpen(entry: SpreadsheetHistoryEntry): void | Promise<void>;
}
export interface SpreadsheetFeatures {
  comments?: boolean;
  suggestions?: boolean;
  search?: boolean;
  panels?: boolean;
  modeControl?: boolean;
}
export interface SpreadsheetHistoricalView {
  timestamp: string;
  onReturnCurrent(): void | Promise<void>;
}
export interface SpreadsheetEditorProps {
  /** Visible editor controls; full is the default. Context menus and permissions work in every preset. */
  chrome?: SpreadsheetChrome;
  sessionKey?: string;
  initialCommentId?: string;
  getCommentLink?(threadId: string): string;
  events?: { subscribeWindowBlur(callback: () => void): () => void };
  features?: SpreadsheetFeatures;
  autosaveMs?: number;
  history?: SpreadsheetHistory;
  collaborators?: readonly SpreadsheetCollaborator[];
  onPresenceChange?(presence: SpreadsheetPresence): void;
  historical?: SpreadsheetHistoricalView;
  value?: Workbook;
  defaultValue?: Workbook;
  onChange?(workbook: Workbook): void;
  calculation?: CalculationMode;
  display?: DisplayOptions;
  readOnly?: boolean;
  mode?: SpreadsheetMode;
  defaultMode?: SpreadsheetMode;
  onModeChange?(mode: SpreadsheetMode): void;
  currentUser?: ReviewActor;
  permissions?: SpreadsheetPermissions;
  title?: string;
  onSave?(workbook: Workbook, context: SpreadsheetSaveContext): void | Promise<void>;
  onRefresh?(target?: { sheetId: string; range: string }): void | Promise<void>;
  /** Hosts own file selection, preview and persistence. Node conversion helpers live in /excel. */
  onImportXlsx?(): Promise<Workbook | null>;
  onExportXlsx?(workbook: Workbook, result: CalculationResult): void | Promise<void>;
  clipboard?: SpreadsheetClipboard;
  panels?: readonly SpreadsheetHostPanel[];
  className?: string;
  'aria-label'?: string;
  onInsert?(insertion: DocumentInsertion): void;
  initialSheetId?: string;
}
export interface GridController {
  select(range: CellSelection): void;
  getAnchor(
    range: CellSelection,
  ): { left: number; top: number; right: number; bottom: number } | null;
  releaseKeyboard(): void;
  /** Returns keyboard focus to this instance after formula entry. */
  focus?(): void;
}
