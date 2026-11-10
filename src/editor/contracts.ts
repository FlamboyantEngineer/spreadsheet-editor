import type { ReactNode } from 'react';
import type { CalculationProvider } from '../calculation/provider.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CalculationOptions, FormulaHost } from '../calculation/evaluator.js';
import type { Workbook } from '../model/workbook.js';
import type { ReviewActor } from '../model/review.js';
import type { DocumentInsertion } from '../document/insertion.js';
import type { CellSelection } from '../model/editing.js';

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
export interface SpreadsheetEditorProps {
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
  onSave?(workbook: Workbook): void | Promise<void>;
  onRefresh?(): void | Promise<void>;
  clipboard?: SpreadsheetClipboard;
  panels?: readonly SpreadsheetHostPanel[];
  className?: string;
  'aria-label'?: string;
  onInsert?(insertion: DocumentInsertion): void;
  initialSheetId?: string;
}
export interface GridController {
  select(range: CellSelection): void;
  getAnchor(range: CellSelection): { left: number; bottom: number } | null;
  releaseKeyboard(): void;
}
