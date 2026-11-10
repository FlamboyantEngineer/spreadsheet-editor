import type { Workbook } from '../model/workbook.js';
export interface XlsxWarning {
  code: string;
  disposition: 'skipped' | 'not-imported';
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
