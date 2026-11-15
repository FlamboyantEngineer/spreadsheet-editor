import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { CalculationResult } from '../calculation/evaluator.js';
import type { DisplayOptions } from '../calculation/format.js';
import { assertCalculationResult } from '../calculation/http.js';
import { requireDefined } from '../model/require-defined.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
  type SpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import { copyWorkbook, type Workbook } from '../model/workbook.js';

export interface DocumentTableCell {
  text: string;
  colspan: number;
  rowspan: number;
}
export interface DocumentTableInsertion {
  kind: 'table';
  schemaVersion: 1;
  sheetId: string;
  range: string;
  /** Calculated text cells; the document editor's `tableFromRows` turns them into its table. */
  rows: DocumentTableCell[][];
  html: string;
}
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

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) =>
    requireDefined(
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character],
    ),
  );
}

export function createDocumentInsertion(
  workbook: Workbook,
  result: CalculationResult | null,
  options: InsertionOptions,
): DocumentInsertion {
  const snapshot = copyWorkbook(workbook);
  const sheet = snapshot.worksheets.find(
    (sheet) => sheet.id === (options.sheetId ?? snapshot.worksheets[0].id),
  );
  if (!sheet) throw new RangeError('Unknown worksheet.');
  const width = Math.max(1, ...sheet.data.map((row) => row.length)),
    height = Math.max(1, sheet.data.length);
  const range: SpreadsheetRange | null = parseSpreadsheetRange(
    options.range ?? `A1:${formatSpreadsheetAddress(width - 1, height - 1)}`,
  );
  if (
    !range ||
    range.end.x > 16_383 ||
    range.end.y > 1_048_575 ||
    (range.end.x - range.start.x + 1) * (range.end.y - range.start.y + 1) > 50_000
  )
    throw new RangeError('Invalid insertion range.');
  if (options.mode === 'spreadsheet')
    return {
      kind: 'spreadsheet',
      schemaVersion: 1,
      sheetId: sheet.id,
      range: range.address,
      workbook: snapshot,
    };
  if (!result) throw new Error('A document table needs current calculation results.');
  assertCalculationResult(result, snapshot);
  const values = requireDefined(result.worksheets.find((output) => output.id === sheet.id)).values;
  const covered = new Set<string>(),
    spans = new Map<string, [number, number]>();
  for (const [address, [colspan, rowspan]] of Object.entries(sheet.mergeCells ?? {})) {
    const merge = parseSpreadsheetRange(address);
    if (
      !merge ||
      !Number.isSafeInteger(colspan) ||
      !Number.isSafeInteger(rowspan) ||
      colspan < 1 ||
      rowspan < 1
    )
      continue;
    const left = Math.max(merge.start.x, range.start.x),
      top = Math.max(merge.start.y, range.start.y);
    const right = Math.min(merge.start.x + colspan - 1, range.end.x),
      bottom = Math.min(merge.start.y + rowspan - 1, range.end.y);
    if (left > right || top > bottom) continue;
    spans.set(`${left}:${top}`, [right - left + 1, bottom - top + 1]);
    for (let y = top; y <= bottom; y++)
      for (let x = left; x <= right; x++) {
        if (x !== left || y !== top) covered.add(`${x}:${y}`);
      }
  }
  const rows: DocumentTableCell[][] = [];
  for (let y = range.start.y; y <= range.end.y; y++) {
    const row: DocumentTableCell[] = [];
    for (let x = range.start.x; x <= range.end.x; x++) {
      if (covered.has(`${x}:${y}`)) continue;
      const [colspan, rowspan] = spans.get(`${x}:${y}`) ?? [1, 1];
      row.push({
        text: formatWorksheetCell(
          values[y]?.[x] ?? null,
          sheet.formats?.[formatSpreadsheetAddress(x, y)],
          options.display,
        ),
        colspan,
        rowspan,
      });
    }
    rows.push(row);
  }
  const html = `<table><tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td colspan="${cell.colspan}" rowspan="${cell.rowspan}">${escapeHtml(cell.text)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  return {
    kind: 'table',
    schemaVersion: 1,
    sheetId: sheet.id,
    range: range.address,
    rows,
    html,
  };
}
