import type { CellValue } from '../calculation/values.js';
import { formatCellValue, type DisplayOptions } from '../calculation/format.js';
import type { CellInput } from '../model/workbook.js';
import type { JSpreadsheetHostHandle, JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

/** Only engine-owned cells are painted here; raw data and formulas remain untouched. */
export function renderCalculatedCells(
  handle: JSpreadsheetHostHandle,
  values: CellValue[][] | undefined,
  display?: DisplayOptions,
  raw: CellInput[][] = [],
  unavailableText = '…',
): void {
  const instance = handle.getInstance();
  if (!instance) return;
  const worksheet: JSpreadsheetWorksheetInstance | undefined = Array.isArray(instance)
    ? instance[0]
    : 'worksheets' in instance
      ? instance.worksheets?.[0]
      : (instance as JSpreadsheetWorksheetInstance);
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
      const text = formatCellValue(value, display);
      if (cell.textContent !== text) cell.textContent = text;
      cell.dataset.approximate = String(
        typeof value === 'object' && value !== null && value.kind === 'number' && value.approximate,
      );
    }),
  );
}
