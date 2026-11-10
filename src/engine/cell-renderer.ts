import type { CellValue } from '../calculation/values.js';
import type { DisplayOptions } from '../calculation/format.js';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { CellInput } from '../model/workbook.js';
import type { CellFormat, SpreadsheetComment, SpreadsheetSuggestion } from '../model/review.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import type { JSpreadsheetHostHandle, JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

function worksheetFrom(handle: JSpreadsheetHostHandle): JSpreadsheetWorksheetInstance | undefined {
  const instance = handle.getInstance();
  if (!instance) return undefined;
  return Array.isArray(instance)
    ? instance[0]
    : 'worksheets' in instance
      ? instance.worksheets?.[0]
      : (instance as JSpreadsheetWorksheetInstance);
}
/** Only engine-owned cells are painted; raw data and formulas remain untouched. */
export function renderCalculatedCells(
  handle: JSpreadsheetHostHandle,
  values: CellValue[][] | undefined,
  display?: DisplayOptions,
  raw: CellInput[][] = [],
  unavailableText = '…',
  formats: Record<string, CellFormat> = {},
  showFormulas = false,
): void {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.getCellFromCoords) return;
  const visible: CellValue[][] =
    values ??
    raw.map((row) =>
      row.map((value) =>
        typeof value === 'string' && value.trimStart().startsWith('=')
          ? unavailableText
          : typeof value === 'number'
            ? String(value)
            : value,
      ),
    );
  visible.forEach((row, y) =>
    row.forEach((value, x) => {
      const cell = worksheet.getCellFromCoords?.(x, y);
      if (!cell || cell.classList.contains('editor')) return;
      const text =
        showFormulas &&
        typeof raw[y]?.[x] === 'string' &&
        String(raw[y][x]).trimStart().startsWith('=')
          ? String(raw[y][x])
          : formatWorksheetCell(value, formats[formatSpreadsheetAddress(x, y)], display);
      if (cell.textContent !== text) cell.textContent = text;
      cell.dataset.approximate = String(
        typeof value === 'object' && value !== null && value.kind === 'number' && value.approximate,
      );
    }),
  );
}
export function renderReviewMarkers(
  handle: JSpreadsheetHostHandle,
  sheetId: string,
  comments: SpreadsheetComment[] = [],
  suggestions: SpreadsheetSuggestion[] = [],
): void {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.getCellFromCoords) return;
  const root = worksheet.element ?? worksheet.root;
  root
    ?.querySelectorAll<HTMLElement>('[data-se-comments], [data-se-suggestions]')
    .forEach((cell) => {
      delete cell.dataset.seComments;
      delete cell.dataset.seSuggestions;
      cell.removeAttribute('title');
    });
  for (const thread of comments) {
    if (thread.sheetId !== sheetId || thread.resolved || thread.orphaned) continue;
    const range = parseSpreadsheetRange(thread.range);
    if (!range) continue;
    const cell = worksheet.getCellFromCoords(range.start.x, range.start.y);
    if (cell) {
      cell.dataset.seComments = String(Number(cell.dataset.seComments ?? 0) + 1);
      cell.title = `${cell.dataset.seComments} open comment(s)`;
    }
  }
  for (const suggestion of suggestions) {
    if (suggestion.status !== 'pending' || suggestion.outdated) continue;
    const addresses = suggestion.changes
      .filter((change) => change.sheetId === sheetId)
      .map((change) => change.address);
    if (suggestion.structure) addresses.push('A1');
    for (const address of addresses) {
      const range = parseSpreadsheetRange(address),
        cell = range && worksheet.getCellFromCoords(range.start.x, range.start.y);
      if (cell) {
        cell.dataset.seSuggestions = String(Number(cell.dataset.seSuggestions ?? 0) + 1);
        cell.title = [cell.title, `Suggestion: ${suggestion.summary}`].filter(Boolean).join('; ');
      }
    }
  }
}
