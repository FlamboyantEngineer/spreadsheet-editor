import type { Workbook } from '../model/workbook.js';
export interface XlsxWarning {
  code: string;
  disposition: 'skipped' | 'not-imported' | 'stored-value' | 'not-exported';
  message: string;
  count: number;
  locations: string[];
}
export interface XlsxReport {
  worksheets: number;
  cells: number;
  formulas: number;
  mergedRanges: number;
  warnings: XlsxWarning[];
}
export interface XlsxImport {
  workbook: Workbook;
  report: XlsxReport;
}

/** Serializable host formula declarations. No resolver is ever invoked by Excel conversion. */
export interface ExcelHostFunction {
  name: string;
  aliases?: readonly string[];
  capturedText?: 'plain' | 'html-to-text';
}
export interface ExcelFormulaCapabilities {
  hostFunctions: readonly ExcelHostFunction[];
}
export interface XlsxExport {
  bytes: Uint8Array;
  report: XlsxReport;
}
