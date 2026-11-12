import type { Worksheet } from '../model/workbook.js';
import { worksheetExtent, type WorksheetExtent } from '../model/extent.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

export function gridWorksheetExtent(
  sheet: Worksheet,
  logicalExtent?: WorksheetExtent,
): [number, number] {
  const extent = worksheetExtent(sheet);
  return [
    Math.max(extent.columns, logicalExtent?.columns ?? 0, sheet.minDimensions ? 1 : 26),
    Math.max(extent.rows, logicalExtent?.rows ?? 0, sheet.minDimensions ? 1 : 100),
  ];
}

/** setData rebuilds body rows, but CE keeps existing column/header structures.
 * Resize those structures first, then set current logical minima before body rebuild.
 * The surrounding controller suppresses authored callbacks for this presentation work.
 */
export function syncWorksheetData(
  engine: JSpreadsheetWorksheetInstance,
  next: Worksheet,
  logicalExtent?: WorksheetExtent,
) {
  const options = engine.options;
  if (!options) return;
  const [columns, rows] = gridWorksheetExtent(next, logicalExtent);
  const currentColumns = engine.cols?.length ?? options.columns?.length ?? columns;
  const permissions = {
    allowInsertColumn: options.allowInsertColumn,
    allowDeleteColumn: options.allowDeleteColumn,
  };
  const merges = options.mergeCells;
  options.mergeCells = {};
  options.minDimensions = [columns, rows];
  options.allowInsertColumn = true;
  options.allowDeleteColumn = true;
  try {
    if (currentColumns < columns) engine.insertColumn?.(columns - currentColumns);
    if (currentColumns > columns) engine.deleteColumn?.(columns, currentColumns - columns);
    options.columns = Array.from({ length: columns }, (_, x) => ({
      ...next.columns?.[x],
      type: 'text',
    }));
    options.rows = structuredClone(next.rows ?? {});
    options.mergeCells = structuredClone(next.mergeCells ?? {});
    engine.setData?.(next.data.map((row) => [...row]));
  } finally {
    options.allowInsertColumn = permissions.allowInsertColumn;
    options.allowDeleteColumn = permissions.allowDeleteColumn;
    if (!options.mergeCells) options.mergeCells = merges;
  }
}
