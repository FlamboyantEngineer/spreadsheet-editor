import type { NumericValue } from '../calculation/values.js';
import { DEFAULT_WORKBOOK_LIMITS } from './limits.js';
import type { CellFormat, SpreadsheetComment, SpreadsheetSuggestion } from './review.js';
import { assertWorkbook } from './workbook-validation.js';
import { worksheetNameKey } from './worksheet-name.js';

export { DEFAULT_WORKBOOK_LIMITS };
export { assertReviewData, assertWorkbook } from './workbook-validation.js';
export { worksheetNameKey } from './worksheet-name.js';

export type CellInput = string | number | boolean | null;

export interface CellProvenance {
  actor: { id: string; name: string };
  source: 'human' | 'ai' | 'system';
  at?: string;
}

export interface Worksheet {
  id: string;
  name: string;
  /** Keep numeric literals as strings to preserve all authored digits. */
  data: CellInput[][];
  inputValues?: Record<string, NumericValue>;
  rowIds?: string[];
  columnIds?: string[];
  minDimensions?: [number, number];
  columns?: { width?: number; title?: string }[];
  rows?: Record<string, { height: number }>;
  mergeCells?: Record<string, [number, number]>;
  style?: Record<string, string>;
  formats?: Record<string, CellFormat>;
  freezeColumns?: number;
  cellProvenance?: Record<string, CellProvenance>;
}

export interface Workbook {
  schemaVersion: 1;
  worksheets: Worksheet[];
  name?: string;
  comments?: SpreadsheetComment[];
  suggestions?: SpreadsheetSuggestion[];
}

export interface WorkbookLimits {
  maxCells?: number;
  maxCellLength?: number;
  maxWorksheets?: number;
}

export function createWorkbook(data: CellInput[][] = [['']]): Workbook {
  return { schemaVersion: 1, worksheets: [{ id: 'sheet-1', name: 'Sheet1', data }] };
}

export function appendWorksheet(workbook: Workbook): Workbook {
  const next = copyWorkbook(workbook);
  let index = next.worksheets.length + 1;
  while (
    next.worksheets.some(
      (sheet) => sheet.id === `sheet-${index}` || worksheetNameKey(sheet.name) === `SHEET${index}`,
    )
  )
    index++;
  next.worksheets.push({ id: `sheet-${index}`, name: `Sheet${index}`, data: [['']] });
  assertWorkbook(next);
  return next;
}

export function copyWorkbook(workbook: Workbook): Workbook {
  assertWorkbook(workbook);
  return structuredClone(workbook);
}
