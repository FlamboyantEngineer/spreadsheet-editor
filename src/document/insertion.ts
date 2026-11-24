import type { CalculationResult } from '../calculation/evaluator.js';
import type { DisplayOptions } from '../calculation/format.js';
import { assertCalculationResult } from '../calculation/result-validation.js';
import { assertWorkbook, type Workbook } from '../model/workbook.js';
import {
  createDocumentTable,
  type DocumentTableInsertion,
  resolveDocumentSelection,
} from './table-capture.js';

export type { DocumentTableCell, DocumentTableInsertion } from './table-capture.js';

export interface LiveSpreadsheetInsertion {
  kind: 'spreadsheet';
  schemaVersion: 1;
  sheetId: string;
  range: string;
  /** Full source workbook retains formula dependencies outside the selected range. */
  workbook: Workbook;
}
export type DocumentInsertion = DocumentTableInsertion | LiveSpreadsheetInsertion;
export interface InsertionOptions {
  mode: 'table' | 'spreadsheet';
  sheetId?: string;
  range?: string;
  display?: DisplayOptions;
}

/** Captures a calculated table or an independent, editable workbook. */
export function createDocumentInsertion(
  workbook: Workbook,
  result: CalculationResult | null,
  options: InsertionOptions,
): DocumentInsertion {
  assertWorkbook(workbook);
  if (options.mode === 'spreadsheet') {
    const { sheet, range } = resolveDocumentSelection(workbook, options);
    return {
      kind: 'spreadsheet',
      schemaVersion: 1,
      sheetId: sheet.id,
      range: range.address,
      workbook: structuredClone(workbook),
    };
  }
  if (!result) throw new Error('A document table needs current calculation results.');
  assertCalculationResult(result, workbook);
  return createDocumentTable(workbook, result, options);
}
