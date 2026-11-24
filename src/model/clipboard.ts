import { writeAuthoredCellState } from './authored-cell.js';
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
  rowIds?: string[] | undefined;
  columnIds?: string[] | undefined;
  /** Without complete stable line identities, a cut needs the original worksheet state. */
  legacySource?: string | undefined;
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
  const rowIds = sheet.rowIds?.slice(range.y1, range.y2 + 1);
  const columnIds = sheet.columnIds?.slice(range.x1, range.x2 + 1);
  const stableLines = rowIds?.length === cells.length && columnIds?.length === cells[0].length;
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
    rowIds,
    columnIds,
    legacySource: cut && !stableLines ? JSON.stringify(sheet) : undefined,
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

function assertOriginalCutLines(source: Sheet, copy: SpreadsheetCopy) {
  if (
    copy.rowIds?.some((id, offset) => source.rowIds?.[copy.selection.y1 + offset] !== id) ||
    copy.columnIds?.some((id, offset) => source.columnIds?.[copy.selection.x1 + offset] !== id)
  )
    throw new Error('A cut source row or column moved. Copy it again before moving.');
  if (copy.legacySource !== undefined && JSON.stringify(source) !== copy.legacySource)
    throw new Error('The cut source worksheet changed. Copy it again before moving.');
}

/** Clears a cut's source once, after checking none of its cells changed since the copy. */
function clearCutSource(next: Workbook, copy: SpreadsheetCopy) {
  const source = next.worksheets.find((candidate) => candidate.id === copy.sheetId);
  if (!source) throw new RangeError('The cut worksheet no longer exists.');
  assertOriginalCutLines(source, copy);
  const cells = selectionCells(copy.selection);
  for (const { address, x, y } of cells) {
    const dy = y - copy.selection.y1,
      dx = x - copy.selection.x1;
    const expected = copy.cellStates?.[dy]?.[dx] ?? { value: copy.cells[dy]?.[dx] ?? null };
    if (JSON.stringify(authoredCellState(source, address)) !== JSON.stringify(expected))
      throw new Error('A cut source cell changed. Copy it again before moving.');
  }
  for (const { address } of cells) writeAuthoredCellState(source, address, { value: null });
}

/** One pasted cell: copies translate relative references, cuts move formulas unchanged. */
function pasteCell(
  sheet: Sheet,
  copy: SpreadsheetCopy,
  { x, y, sourceX, sourceY }: { x: number; y: number; sourceX: number; sourceY: number },
) {
  const raw = copy.cells[sourceY][sourceX] ?? null;
  const value =
    typeof raw === 'string' && !copy.cut
      ? translateFormulaReferences(
          raw,
          x - copy.selection.x1 - sourceX,
          y - copy.selection.y1 - sourceY,
        )
      : raw;
  writeAuthoredCellState(sheet, formatSpreadsheetAddress(x, y), {
    ...copy.cellStates?.[sourceY]?.[sourceX],
    value,
    materialize: true,
  });
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
