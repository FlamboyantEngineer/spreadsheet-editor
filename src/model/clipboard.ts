import type { CellInput, Workbook } from './workbook.js';
import { readSelection, writeCells, normalizeSelection, type CellSelection } from './editing.js';
import { translateFormulaReferences } from './structure.js';

export interface SpreadsheetCopy {
  text: string;
  sheetId: string;
  selection: CellSelection;
  cells: CellInput[][];
  cut: boolean;
}

export function createSpreadsheetCopy(
  workbook: Workbook,
  sheetId: string,
  selection: CellSelection,
  text: string,
  cut = false,
): SpreadsheetCopy {
  return {
    text,
    sheetId,
    selection: normalizeSelection(selection),
    cells: readSelection(workbook, sheetId, selection),
    cut,
  };
}

/** A matching plain-text clipboard carries authored formulas; foreign text never does. */
export function pasteSpreadsheetCopy(
  workbook: Workbook,
  sheetId: string,
  x: number,
  y: number,
  copy: SpreadsheetCopy,
): Workbook {
  const dx = x - copy.selection.x1,
    dy = y - copy.selection.y1;
  return writeCells(
    workbook,
    sheetId,
    x,
    y,
    copy.cells.map((row) =>
      row.map((cell) =>
        typeof cell === 'string' && !copy.cut ? translateFormulaReferences(cell, dx, dy) : cell,
      ),
    ),
  );
}
