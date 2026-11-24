import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { CalculationResult } from '../calculation/evaluator.js';
import type { DisplayOptions } from '../calculation/format.js';
import { requireDefined } from '../model/require-defined.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
  parseSpreadsheetRange,
  type SpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import type { Workbook, Worksheet } from '../model/workbook.js';

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

interface DocumentRangeOptions {
  sheetId?: string;
  range?: string;
}

interface TableCaptureOptions extends DocumentRangeOptions {
  display?: DisplayOptions;
  assertReady?: (x: number, y: number) => void;
}

interface MergeCell {
  x: number;
  y: number;
  colspan: number;
  rowspan: number;
}

const maxColumns = 16_384;
const maxRows = 1_048_576;
const maxTableCells = 50_000;
const htmlEntities: Readonly<Record<string, string>> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => htmlEntities[character]);
}

function dataRange(sheet: Worksheet): string {
  const width = sheet.data.reduce((widest, row) => Math.max(widest, row.length), 1);
  return `A1:${formatSpreadsheetAddress(width - 1, Math.max(1, sheet.data.length) - 1)}`;
}

function insertionRange(sheet: Worksheet, address: string | undefined): SpreadsheetRange {
  const range = parseSpreadsheetRange(address ?? dataRange(sheet));
  if (
    !range ||
    range.end.x >= maxColumns ||
    range.end.y >= maxRows ||
    (range.end.x - range.start.x + 1) * (range.end.y - range.start.y + 1) > maxTableCells
  )
    throw new RangeError('Invalid insertion range.');
  return range;
}

/** Finds the worksheet and checks only the requested range when one was supplied. */
export function resolveDocumentSelection(
  workbook: Workbook,
  { sheetId = workbook.worksheets[0].id, range }: DocumentRangeOptions,
) {
  const sheetIndex = workbook.worksheets.findIndex((sheet) => sheet.id === sheetId);
  if (sheetIndex < 0) throw new RangeError('Unknown worksheet.');
  const sheet = workbook.worksheets[sheetIndex];
  return { sheet, sheetIndex, range: insertionRange(sheet, range) };
}

const cellIndex = (x: number, y: number) => y * maxColumns + x;

/** A clipped merge starts at its first included cell and still displays its original anchor. */
function addMerge(
  layout: Map<number, MergeCell | null>,
  range: SpreadsheetRange,
  [address, [colspan, rowspan]]: [string, [number, number]],
): void {
  const { x, y } = requireDefined(parseSpreadsheetAddress(address));
  const left = Math.max(x, range.start.x),
    top = Math.max(y, range.start.y),
    right = Math.min(x + colspan - 1, range.end.x),
    bottom = Math.min(y + rowspan - 1, range.end.y);
  if (left > right || top > bottom) return;
  for (let row = top; row <= bottom; row++)
    for (let column = left; column <= right; column++)
      if (column !== left || row !== top) layout.set(cellIndex(column, row), null);
  const key = cellIndex(left, top);
  if (layout.get(key) !== null)
    layout.set(key, { x, y, colspan: right - left + 1, rowspan: bottom - top + 1 });
}

function tableRows(
  sheet: Worksheet,
  values: CalculationResult['worksheets'][number]['values'],
  {
    range,
    display,
    assertReady,
  }: Pick<TableCaptureOptions, 'display' | 'assertReady'> & {
    range: SpreadsheetRange;
  },
): DocumentTableCell[][] {
  const merges = new Map<number, MergeCell | null>();
  for (const merge of Object.entries(sheet.mergeCells ?? {})) addMerge(merges, range, merge);
  const rows: DocumentTableCell[][] = [];
  for (let y = range.start.y; y <= range.end.y; y++) {
    const row: DocumentTableCell[] = [];
    for (let x = range.start.x; x <= range.end.x; x++) {
      const merge = merges.get(cellIndex(x, y));
      if (merge === null) continue;
      const {
        x: sourceX,
        y: sourceY,
        colspan,
        rowspan,
      } = merge ?? { x, y, colspan: 1, rowspan: 1 };
      assertReady?.(sourceX, sourceY);
      const format = sheet.formats?.[formatSpreadsheetAddress(sourceX, sourceY)];
      row.push({
        text: formatWorksheetCell(values[sourceY]?.[sourceX] ?? null, format, display),
        colspan,
        rowspan,
      });
    }
    rows.push(row);
  }
  return rows;
}

const tableHtml = (rows: DocumentTableCell[][]) =>
  `<table><tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td colspan="${cell.colspan}" rowspan="${cell.rowspan}">${escapeHtml(cell.text)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;

/** Projects a validated workbook and its matching result without copying either input. */
export function createDocumentTable(
  workbook: Workbook,
  result: CalculationResult,
  options: TableCaptureOptions,
): DocumentTableInsertion {
  const { sheet, sheetIndex, range } = resolveDocumentSelection(workbook, options);
  const rows = tableRows(sheet, result.worksheets[sheetIndex].values, {
    range,
    display: options.display,
    assertReady: options.assertReady,
  });
  return {
    kind: 'table',
    schemaVersion: 1,
    sheetId: sheet.id,
    range: range.address,
    rows,
    html: tableHtml(rows),
  };
}
