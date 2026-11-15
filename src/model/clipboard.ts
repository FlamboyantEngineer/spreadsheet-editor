import {
  type CellSelection,
  normalizeSelection,
  readSelection,
  selectionCells,
} from './editing.js';
import { extendWorksheetIdentities, type WorkbookIdFactory } from './extent.js';
import { type AuthoredCellState, authoredCellState } from './operations.js';
import { formatSpreadsheetAddress } from './spreadsheet-addressing.js';
import { translateFormulaReferences } from './structure.js';
import {
  assertWorkbook,
  type CellInput,
  copyWorkbook,
  DEFAULT_WORKBOOK_LIMITS,
  type Workbook,
} from './workbook.js';

export interface SpreadsheetCopy {
  text: string;
  sheetId: string;
  selection: CellSelection;
  cells: CellInput[][];
  cellStates?: AuthoredCellState[][];
  cut: boolean;
}

export function createSpreadsheetCopy(
  workbook: Workbook,
  sheetId: string,
  selection: CellSelection,
  text: string,
  cut = false,
): SpreadsheetCopy {
  const range = normalizeSelection(selection),
    sheet = workbook.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  const cells = readSelection(workbook, sheetId, range);
  return {
    text,
    sheetId,
    selection: range,
    cells,
    cellStates: cells.map((row, dy) =>
      row.map((_cell, dx) =>
        authoredCellState(sheet, formatSpreadsheetAddress(range.x1 + dx, range.y1 + dy)),
      ),
    ),
    cut,
  };
}

/** Exact multiples tile against original source coordinates; cuts move once atomically. */
export function pasteSpreadsheetCopy(
  workbook: Workbook,
  sheetId: string,
  x: number,
  y: number,
  copy: SpreadsheetCopy,
  destination: CellSelection = { x1: x, y1: y, x2: x, y2: y },
  idFactory?: WorkbookIdFactory,
): Workbook {
  const sourceWidth = Math.max(0, ...copy.cells.map((row) => row.length)),
    sourceHeight = copy.cells.length;
  if (!sourceWidth || !sourceHeight) return copyWorkbook(workbook);
  const selected = normalizeSelection(destination),
    selectedWidth = selected.x2 - selected.x1 + 1,
    selectedHeight = selected.y2 - selected.y1 + 1;
  const tile =
    !copy.cut &&
    selectedWidth >= sourceWidth &&
    selectedHeight >= sourceHeight &&
    selectedWidth % sourceWidth === 0 &&
    selectedHeight % sourceHeight === 0;
  const width = tile ? selectedWidth : sourceWidth,
    height = tile ? selectedHeight : sourceHeight;
  normalizeSelection({ x1: x, y1: y, x2: x + width - 1, y2: y + height - 1 });
  if (width * height > DEFAULT_WORKBOOK_LIMITS.maxCells)
    throw new RangeError('Paste exceeds the cell limit.');
  // Validate all limits before source clearing, including sparse target expansion.
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  if (
    Math.max(sheet.minDimensions?.[0] ?? 1, x + width) *
      Math.max(sheet.minDimensions?.[1] ?? 1, y + height) >
    DEFAULT_WORKBOOK_LIMITS.maxCells
  )
    throw new RangeError('Paste exceeds the cell limit.');
  if (copy.cut) {
    const source = next.worksheets.find((candidate) => candidate.id === copy.sheetId);
    if (!source) throw new RangeError('The cut worksheet no longer exists.');
    for (const { address, x: sx, y: sy } of selectionCells(copy.selection)) {
      const expected = copy.cellStates?.[sy - copy.selection.y1]?.[sx - copy.selection.x1];
      if (
        expected &&
        JSON.stringify(authoredCellState(source, address)) !== JSON.stringify(expected)
      )
        throw new Error('A cut source cell changed. Copy it again before moving.');
    }
    for (const { address, x: sx, y: sy } of selectionCells(copy.selection)) {
      if (source.data[sy]?.[sx] !== undefined) source.data[sy][sx] = null;
      for (const map of [source.style, source.formats, source.inputValues, source.cellProvenance])
        if (map) delete map[address];
    }
  }
  for (let dy = 0; dy < height; dy++) {
    while (sheet.data.length <= y + dy) sheet.data.push([]);
    for (let dx = 0; dx < width; dx++) {
      const sourceX = dx % sourceWidth,
        sourceY = dy % sourceHeight,
        raw = copy.cells[sourceY][sourceX] ?? null,
        address = formatSpreadsheetAddress(x + dx, y + dy);
      while (sheet.data[y + dy].length <= x + dx) sheet.data[y + dy].push(null);
      sheet.data[y + dy][x + dx] =
        typeof raw === 'string' && !copy.cut
          ? translateFormulaReferences(
              raw,
              x + dx - copy.selection.x1 - sourceX,
              y + dy - copy.selection.y1 - sourceY,
            )
          : raw;
      const state = copy.cellStates?.[sourceY]?.[sourceX];
      for (const [field, stateField] of [
        ['style', 'style'],
        ['formats', 'format'],
        ['inputValues', 'inputValue'],
        ['cellProvenance', 'provenance'],
      ] as const) {
        if (state?.[stateField] !== undefined) {
          // Each field retains its own schema; assignment is bounded by the copy's validated workbook.
          const map = (sheet[field] ??= {}) as Record<string, unknown>;
          map[address] = structuredClone(state[stateField]);
        } else if (sheet[field]) delete sheet[field][address];
      }
    }
  }
  extendWorksheetIdentities(sheet, idFactory);
  assertWorkbook(next);
  return next;
}
