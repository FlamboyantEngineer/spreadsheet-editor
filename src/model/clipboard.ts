import {
  type CellPlacement,
  type CellSelection,
  normalizeSelection,
  readSelection,
  selectionCells,
  type SheetRange,
} from './editing.js';
import { extendWorksheetIdentities } from './extent.js';
import { translateFormulaReferences } from './formula-references.js';
import { type AuthoredCellState, authoredCellState } from './operations.js';
import { formatSpreadsheetAddress } from './spreadsheet-addressing.js';
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
  { sheetId, range: selection }: SheetRange,
  { text, cut = false }: { text: string; cut?: boolean | undefined },
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

type Sheet = Workbook['worksheets'][number];

/** Where a copy is pasted: its top-left cell, and the selection it may tile across. */
export interface PastePlacement extends CellPlacement {
  /** The selected destination; an exact multiple of the copy's size is tiled. Defaults to the cell. */
  destination?: CellSelection | undefined;
}

/** The size of the pasted block: the whole selection when the copy tiles it, else the copy's. */
function pasteSize(
  copy: SpreadsheetCopy,
  { sourceWidth, sourceHeight }: { sourceWidth: number; sourceHeight: number },
  destination: CellSelection,
) {
  const selected = normalizeSelection(destination),
    selectedWidth = selected.x2 - selected.x1 + 1,
    selectedHeight = selected.y2 - selected.y1 + 1;
  const tile =
    !copy.cut &&
    selectedWidth >= sourceWidth &&
    selectedHeight >= sourceHeight &&
    selectedWidth % sourceWidth === 0 &&
    selectedHeight % sourceHeight === 0;
  return {
    width: tile ? selectedWidth : sourceWidth,
    height: tile ? selectedHeight : sourceHeight,
  };
}

/** Clears a cut's source once, after checking none of its cells changed since the copy. */
function clearCutSource(next: Workbook, copy: SpreadsheetCopy) {
  const source = next.worksheets.find((candidate) => candidate.id === copy.sheetId);
  if (!source) throw new RangeError('The cut worksheet no longer exists.');
  const cells = selectionCells(copy.selection);
  for (const { address, x, y } of cells) {
    const expected = copy.cellStates?.[y - copy.selection.y1]?.[x - copy.selection.x1];
    if (expected && JSON.stringify(authoredCellState(source, address)) !== JSON.stringify(expected))
      throw new Error('A cut source cell changed. Copy it again before moving.');
  }
  for (const { address, x, y } of cells) {
    if (source.data[y]?.[x] !== undefined) source.data[y][x] = null;
    for (const map of [source.style, source.formats, source.inputValues, source.cellProvenance])
      if (map) delete map[address];
  }
}

const STATE_FIELDS = [
  ['style', 'style'],
  ['formats', 'format'],
  ['inputValues', 'inputValue'],
  ['cellProvenance', 'provenance'],
] as const;

/** Copies a source cell's style, format, input value and provenance, or clears absent ones. */
function pasteState(sheet: Sheet, address: string, state: AuthoredCellState | undefined) {
  for (const [field, stateField] of STATE_FIELDS) {
    const value = state?.[stateField];
    if (value === undefined) {
      if (sheet[field]) delete sheet[field][address];
      continue;
    }
    // Each field retains its own schema; assignment is bounded by the copy's validated workbook.
    const map = (sheet[field] ??= {}) as Record<string, unknown>;
    map[address] = structuredClone(value);
  }
}

/** One pasted cell: copies translate relative references, cuts move formulas unchanged. */
function pasteCell(
  sheet: Sheet,
  copy: SpreadsheetCopy,
  { x, y, sourceX, sourceY }: { x: number; y: number; sourceX: number; sourceY: number },
) {
  const raw = copy.cells[sourceY][sourceX] ?? null;
  while (sheet.data[y].length <= x) sheet.data[y].push(null);
  sheet.data[y][x] =
    typeof raw === 'string' && !copy.cut
      ? translateFormulaReferences(
          raw,
          x - copy.selection.x1 - sourceX,
          y - copy.selection.y1 - sourceY,
        )
      : raw;
  pasteState(sheet, formatSpreadsheetAddress(x, y), copy.cellStates?.[sourceY]?.[sourceX]);
}

/** Exact multiples tile against original source coordinates; cuts move once atomically. */
export function pasteSpreadsheetCopy(
  workbook: Workbook,
  copy: SpreadsheetCopy,
  { sheetId, x, y, destination = { x1: x, y1: y, x2: x, y2: y }, idFactory }: PastePlacement,
): Workbook {
  const sourceWidth = Math.max(0, ...copy.cells.map((row) => row.length)),
    sourceHeight = copy.cells.length;
  if (!sourceWidth || !sourceHeight) return copyWorkbook(workbook);
  const { width, height } = pasteSize(copy, { sourceWidth, sourceHeight }, destination);
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
  if (copy.cut) clearCutSource(next, copy);
  for (let dy = 0; dy < height; dy++) {
    while (sheet.data.length <= y + dy) sheet.data.push([]);
    for (let dx = 0; dx < width; dx++) {
      const position = {
        x: x + dx,
        y: y + dy,
        sourceX: dx % sourceWidth,
        sourceY: dy % sourceHeight,
      };
      pasteCell(sheet, copy, position);
    }
  }
  extendWorksheetIdentities(sheet, idFactory);
  assertWorkbook(next);
  return next;
}
