import type { ReactNode, Ref } from 'react';
import type { ExternalCalculationSnapshot } from '../calculation/publication.js';
import type { WorkbookIdFactory } from '../model/extent.js';
import type { WorkbookChangeContext } from '../model/operations.js';
import type {
  CellPresentationSnapshot,
  FormulaFunctionDescriptor,
  RenderCellContent,
  SpreadsheetEditorHandle,
} from './public-host.js';
export * from './public-host.js';
import type {
  CalculationOptions,
  CalculationResult,
  FormulaHost,
} from '../calculation/evaluator.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CalculationProvider } from '../calculation/provider.js';
import type { DocumentInsertion } from '../document/insertion.js';
import type { IdentityAdapter, VersionStoreAdapter } from '../host/adapters.js';
import type { EditorCapabilities, SpreadsheetMode } from '../host/capabilities.js';
import type { PresenceAdapter } from '../host/presence-adapters.js';
import type { WorkbookSession } from '../host/workbook-session.js';
import type { CellSelection } from '../model/editing.js';
import type { CellFormat } from '../model/review.js';
import type { Workbook } from '../model/workbook.js';
export type { SpreadsheetMode } from '../host/capabilities.js';

export type SpreadsheetChrome = 'minimal' | 'embedded' | 'full';
export type CalculationMode =
  | { mode: 'client'; host?: FormulaHost; options?: CalculationOptions }
  | { mode: 'server'; provider: CalculationProvider; options?: CalculationOptions }
  | { mode: 'external'; snapshot: ExternalCalculationSnapshot | null };
export interface SpreadsheetClipboard {
  readText(): Promise<string>;
  writeText(text: string): Promise<void>;
}
export interface SpreadsheetHostPanel {
  id: string;
  label: string;
  content: ReactNode;
}
export interface SpreadsheetFeatures {
  newWorkbook?: boolean;
  import?: boolean;
  comments?: boolean;
  suggestions?: boolean;
  search?: boolean;
  panels?: boolean;
  modeControl?: boolean;
}
export interface SpreadsheetEditorProps {
  editorRef?: Ref<SpreadsheetEditorHandle>;
  idFactory?: WorkbookIdFactory;
  undo?: {
    canUndo: boolean;
    canRedo: boolean;
    undo(): void | Promise<void>;
    redo(): void | Promise<void>;
  };
  formulaFunctions?: readonly FormulaFunctionDescriptor[];
  /** An authoritative host may supply its complete supported formula catalog. */
  includeBuiltinFormulaFunctions?: boolean;
  formatting?: { kinds?: readonly CellFormat['kind'][]; decimalPlaces?: boolean };
  cellPresentation?: CellPresentationSnapshot;
  renderCellContent?: RenderCellContent;
  /** Visible editor controls; full is the default. Context menus and permissions work in every preset. */
  chrome?: SpreadsheetChrome;
  sessionKey?: string;
  initialCommentId?: string;
  getCommentLink?(threadId: string): string;
  events?: { subscribeWindowBlur(callback: () => void): () => void };
  features?: SpreadsheetFeatures;
  /**
   * The workbook's id in the host. It addresses the workbook in the host's adapters, and a new id
   * starts a new editor session, as a new `sessionKey` does.
   */
  workbookId?: string;
  /**
   * The stored workbook's session (`createWorkbookSession`): the editor shows its workbook,
   * records each accepted change in it, and saves and flushes through it (SE-SESSION-4). Null
   * shows the editor as opening, as `useWorkbookSession` returns before its session exists.
   */
  session?: WorkbookSession | null;
  /** Shares the reader's place with the other people in the workbook, given a `workbookId`. */
  presence?: PresenceAdapter;
  /**
   * The workbook's saved versions, given a `workbookId`. The history pane lists them; the editor
   * opens one read-only in place of the workbook and returns to the current workbook itself, and
   * saves, renames and restores versions where the store and the reader's capabilities allow
   * (SE-ADAPTERS-5.1).
   */
  versions?: VersionStoreAdapter;
  value?: Workbook;
  defaultValue?: Workbook;
  onChange?(workbook: Workbook, context: WorkbookChangeContext): void;
  calculation?: CalculationMode;
  /** Authoritative results for this actor's suggestion projection, with its own input key. */
  reviewCalculation?: ExternalCalculationSnapshot | null;
  /** Draft notifications let the host retain uncommitted input independently of accepted data. */
  onDraftChange?(draft: { sheetId: string; x: number; y: number; text: string } | null): void;
  display?: DisplayOptions;
  mode?: SpreadsheetMode;
  defaultMode?: SpreadsheetMode;
  onModeChange?(mode: SpreadsheetMode): void;
  /** Who the reader is: the actor of their comments, suggestions and decisions. */
  identity?: IdentityAdapter;
  /**
   * What the reader may do, as the host grants it; every capability by default. The editor never
   * offers more: a historical version is always read-only.
   */
  capabilities?: EditorCapabilities;
  title?: string;
  onRefresh?(target?: { sheetId: string; range: string }): void | Promise<void>;
  /** Hosts own file selection, preview and persistence. Node conversion helpers live in /excel. */
  onImportXlsx?(): Promise<Workbook | null>;
  onExportXlsx?(workbook: Workbook, result: CalculationResult): void | Promise<void>;
  clipboard?: SpreadsheetClipboard;
  panels?: readonly SpreadsheetHostPanel[];
  panelRef?: Ref<HTMLDivElement>;
  panelId?: string;
  panelData?: Readonly<Record<`data-${string}`, string | undefined>>;
  onPanelChange?(state: { id: string | null; expanded: boolean }): void;
  className?: string;
  'aria-label'?: string;
  onInsert?(insertion: DocumentInsertion): void;
  initialSheetId?: string;
}
export interface GridController {
  commitDraft?(): void;
  cancelDraft?(): void;
  getViewport?(): { left: number; top: number };
  restoreViewport?(viewport: { left: number; top: number }): void;
  select(range: CellSelection): void;
  getAnchor(
    range: CellSelection,
  ): { left: number; top: number; right: number; bottom: number } | null;
  releaseKeyboard(): void;
  /** Returns keyboard focus to this instance after formula entry. */
  focus?(): void;
}
