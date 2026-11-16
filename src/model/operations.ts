import type { NumericValue } from '../calculation/values.js';
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
  before: AuthoredCellState;
  after: AuthoredCellState;
}
export type WorksheetGeometry = Pick<
  Worksheet,
  'minDimensions' | 'columns' | 'rows' | 'freezeColumns'
>;
export type WorkbookOperation =
  | { kind: 'workbook-metadata'; before: Pick<Workbook, 'name'>; after: Pick<Workbook, 'name'> }
  | { kind: 'cells'; sheetId: string; changes: WorkbookCellChange[] }
  | {
      kind: 'lines';
      sheetId: string;
      axis: 'row' | 'column';
      action: 'insert' | 'delete';
      index: number;
      count: number;
      ids: string[];
      beforeId?: string;
      afterId?: string;
      /** Removed authored cells and sparse geometry support semantic standalone undo. */
      deletedCells?: WorkbookCellChange[];
      beforeGeometry?: WorksheetGeometry;
      afterGeometry?: WorksheetGeometry;
      beforeMerges?: NonNullable<Worksheet['mergeCells']>;
      afterMerges?: NonNullable<Worksheet['mergeCells']>;
    }
  | {
      kind: 'sheet-add';
      sheet: Worksheet;
      index: number;
      beforeOrder: string[];
      afterOrder: string[];
    }
  | {
      kind: 'sheet-delete';
      sheet: Worksheet;
      index: number;
      beforeOrder: string[];
      afterOrder: string[];
    }
  | { kind: 'sheet-rename'; sheetId: string; before: string; after: string }
  | { kind: 'sheet-reorder'; before: string[]; after: string[] }
  | { kind: 'geometry'; sheetId: string; before: WorksheetGeometry; after: WorksheetGeometry }
  | {
      kind: 'merges';
      sheetId: string;
      before: NonNullable<Worksheet['mergeCells']>;
      after: NonNullable<Worksheet['mergeCells']>;
    }
  | {
      kind: 'annotations';
      before: Pick<Workbook, 'comments' | 'suggestions'>;
      after: Pick<Workbook, 'comments' | 'suggestions'>;
    }
  | {
      kind: 'replace';
      reason: 'import' | 'new-workbook' | 'reset';
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
