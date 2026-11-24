import type { NumericValue } from '../calculation/values.js';
import type { CellOperationShape } from './operation-cell-shape.js';
import type { CellFormat } from './review.js';
import type { CellInput, CellProvenance, Workbook, Worksheet } from './workbook.js';

export {
  authoredCellState,
  collectCellOperations,
  collectLineGrowthOperations,
  writeCellsWithOperations,
} from './cell-operations.js';
export { invertWorkbookOperations } from './operation-inverse.js';
export {
  annotationOperation,
  changeWorkbookStructureWithOperations,
  geometryOperation,
  mergeOperation,
  type SheetCommand,
  sheetOperation,
  type StructureTarget,
} from './structure-operations.js';

export interface AuthoredCellState {
  value: CellInput;
  style?: string;
  format?: CellFormat;
  inputValue?: NumericValue;
  provenance?: CellProvenance;
}
export interface WorkbookCellChange {
  address: string;
  /** Old coordinates may differ after a line command; stable IDs identify the same cell. */
  beforeAddress?: string;
  rowId?: string;
  columnId?: string;
  /** Restoring a deleted line can restore an explicit blank value. */
  materializeValue?: true;
  before: AuthoredCellState;
  after: AuthoredCellState;
}
export type WorksheetGeometry = Pick<
  Worksheet,
  'minDimensions' | 'columns' | 'rows' | 'freezeColumns'
>;
export type WorkbookOperation =
  | { kind: 'workbook-metadata'; before: Pick<Workbook, 'name'>; after: Pick<Workbook, 'name'> }
  | { kind: 'cells'; sheetId: string; changes: WorkbookCellChange[]; shape?: CellOperationShape }
  | {
      kind: 'lines';
      sheetId: string;
      axis: 'row' | 'column';
      action: 'insert' | 'delete';
      index: number;
      count: number;
      ids: string[];
      /** An ordinary write extends the identities without moving existing cells or layout. */
      extentOnly?: true;
      /** A line command can first materialize an implicit viewport in an exact recorded frame. */
      preparation?: LinePreparation | undefined;
      restoreLayout?: true;
      /** Raw row counts preserve a deleted trailing empty row without authoring a value. */
      dataRows?: { before: number; after: number } | undefined;
      beforeId?: string;
      afterId?: string;
      /** Removed authored cells and sparse geometry support semantic standalone undo. */
      deletedCells?: WorkbookCellChange[];
      beforeGeometry?: WorksheetGeometry;
      afterGeometry?: WorksheetGeometry;
      beforeMerges?: Worksheet['mergeCells'];
      afterMerges?: Worksheet['mergeCells'];
    }
  | { kind: 'line-frame'; sheetId: string; before: LineFrame; after: LineFrame }
  | {
      kind: 'sheet-add';
      sheet: Worksheet;
      beforeOrder: string[];
      afterOrder: string[];
    }
  | {
      kind: 'sheet-delete';
      sheet: Worksheet;
      beforeOrder: string[];
      afterOrder: string[];
    }
  | { kind: 'sheet-rename'; sheetId: string; before: string; after: string }
  | { kind: 'sheet-reorder'; before: string[]; after: string[] }
  | { kind: 'geometry'; sheetId: string; before: WorksheetGeometry; after: WorksheetGeometry }
  | {
      kind: 'merges';
      sheetId: string;
      before: Worksheet['mergeCells'];
      after: Worksheet['mergeCells'];
    }
  | {
      kind: 'annotations';
      before: Pick<Workbook, 'comments' | 'suggestions'>;
      after: Pick<Workbook, 'comments' | 'suggestions'>;
    }
  | {
      kind: 'replace';
      reason: 'import' | 'new-workbook' | 'reset' | 'suggestion';
      before: Workbook;
      after: Workbook;
    };
export interface WorkbookChangeContext {
  transactionId: string;
  source: 'user' | 'undo' | 'redo' | 'host-command';
  baseline: Workbook;
  operations: WorkbookOperation[];
}
export interface WorkbookCommandResult {
  workbook: Workbook;
  operations: WorkbookOperation[];
}

/** The fields an editor line command may materialize before applying the requested change. */
type LineFrame = Pick<Worksheet, 'minDimensions' | 'rowIds' | 'columnIds'>;
export interface LinePreparation {
  before: LineFrame;
  after: LineFrame;
}
