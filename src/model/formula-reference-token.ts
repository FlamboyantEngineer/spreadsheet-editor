import {
  FORMULA_WORD_CHARACTER,
  MAX_COLUMN_INDEX,
  MAX_ROW_INDEX,
  parseSpreadsheetAddress,
  parseSpreadsheetColumnName,
  UNQUOTED_SHEET_NAME,
} from './spreadsheet-addressing.js';

export type CellCoords = { x: number; y: number };
export type ReferenceCoordinates =
  | { kind: 'reference'; start: CellCoords; end: CellCoords | null }
  | { kind: 'lines'; axis: 'column' | 'row'; start: number; end: number };

/** An authored reference, including its unmodified spelling and sheet qualifier spans. */
interface ReferenceToken {
  startIndex: number;
  addressIndex: number;
  endIndex: number;
  sheetName: string | null;
  coordinates: ReferenceCoordinates | null;
}

const SHEET_PREFIX = new RegExp(`(${UNQUOTED_SHEET_NAME})!\\s*`, 'uy');
const NAME_RUN = /[\p{L}\p{N}_ ]*/uy;
const CELL = String.raw`\$?[A-Za-z]+\$?[1-9]\d*`;
const COLUMN = String.raw`\$?[A-Za-z]+`;
const ROW = String.raw`\$?[1-9]\d*`;
const BOUNDARY = String.raw`(?![\p{L}\p{N}_.!\[\]]|\s*\()`;
const CELLS = new RegExp(`(${CELL})(?:\\s*:\\s*(${CELL}))?${BOUNDARY}`, 'uy');
const COLUMNS = new RegExp(`(${COLUMN})\\s*:\\s*(${COLUMN})${BOUNDARY}`, 'uy');
const ROWS = new RegExp(`(${ROW})\\s*:\\s*(${ROW})${BOUNDARY}`, 'uy');
const SPACE = /\s*/y;

export function quotedTextEnd(source: string, start: number): number | null {
  const quote = source[start];
  for (let index = start + 1; index < source.length; index++) {
    if (source[index] !== quote) continue;
    if (source[index + 1] !== quote) return index + 1;
    index++;
  }
  return null;
}

function sheetPrefix(source: string, start: number) {
  if (source[start] === "'") {
    const end = quotedTextEnd(source, start);
    if (end === null || source[end] !== '!') return null;
    SPACE.lastIndex = end + 1;
    SPACE.exec(source);
    return { name: source.slice(start + 1, end - 1).replaceAll("''", "'"), end: SPACE.lastIndex };
  }
  SHEET_PREFIX.lastIndex = start;
  const match = SHEET_PREFIX.exec(source);
  return match && { name: match[1], end: SHEET_PREFIX.lastIndex };
}

const inGrid = (cell: CellCoords | null): cell is CellCoords =>
  cell !== null && cell.x <= MAX_COLUMN_INDEX && cell.y <= MAX_ROW_INDEX;

function cellCoordinates(first: string, last: string | undefined): ReferenceCoordinates | null {
  const start = parseSpreadsheetAddress(first);
  const end = last ? parseSpreadsheetAddress(last) : null;
  return inGrid(start) && (!last || inGrid(end)) ? { kind: 'reference', start, end } : null;
}

function lineCoordinates(
  match: RegExpExecArray,
  axis: 'column' | 'row',
): ReferenceCoordinates | null {
  const coordinate = (value: string) => {
    const bare = value.replace('$', '');
    return axis === 'column' ? parseSpreadsheetColumnName(bare) : Number(bare) - 1;
  };
  const start = coordinate(match[1]),
    end = coordinate(match[2]),
    last = axis === 'column' ? MAX_COLUMN_INDEX : MAX_ROW_INDEX;
  return start !== null && end !== null && start <= last && end <= last
    ? { kind: 'lines', axis, start, end }
    : null;
}

/** Reads a reference at a parser or scanner position. Invalid grid coordinates retain their span. */
export function formulaReferenceAt(
  source: string,
  startIndex: number,
  allowQualifier = true,
): ReferenceToken | null {
  const qualifier = allowQualifier ? sheetPrefix(source, startIndex) : null;
  const addressIndex = qualifier?.end ?? startIndex;
  const base = { startIndex, addressIndex, sheetName: qualifier?.name ?? null };
  CELLS.lastIndex = COLUMNS.lastIndex = ROWS.lastIndex = addressIndex;
  const cells = CELLS.exec(source);
  if (cells)
    return { ...base, endIndex: CELLS.lastIndex, coordinates: cellCoordinates(cells[1], cells[2]) };
  const columns = COLUMNS.exec(source);
  if (columns)
    return {
      ...base,
      endIndex: COLUMNS.lastIndex,
      coordinates: lineCoordinates(columns, 'column'),
    };
  const rows = ROWS.exec(source);
  return rows && { ...base, endIndex: ROWS.lastIndex, coordinates: lineCoordinates(rows, 'row') };
}

function bracketEnd(source: string, start: number): number {
  let depth = 1;
  let index = start + 1;
  while (index < source.length && depth) {
    const character = source[index];
    if (character === '"' || character === "'") {
      index = quotedTextEnd(source, index) ?? source.length;
      continue;
    }
    if (character === '[') depth++;
    if (character === ']') depth--;
    index++;
  }
  return formulaReferenceAt(source, index)?.endIndex ?? index;
}

function ignoredEnd(source: string, index: number): number {
  const character = source[index];
  if (character === '"' || character === "'") return quotedTextEnd(source, index) ?? source.length;
  if (character === '[') return bracketEnd(source, index);
  if (source.startsWith('{{', index)) {
    const end = source.indexOf('}}', index + 2);
    return end < 0 ? source.length : end + 2;
  }
  return index + 1;
}

/** Scans partial formulas without treating text, host bindings or external references as cells. */
export function* formulaReferenceTokens(source: string): Generator<ReferenceToken> {
  let index = 0;
  let runEnd = -1;
  while (index < source.length) {
    if (!FORMULA_WORD_CHARACTER.test(source[index - 1] ?? '')) {
      if (index >= runEnd) {
        NAME_RUN.lastIndex = index;
        NAME_RUN.exec(source);
        runEnd = NAME_RUN.lastIndex;
      }
      const qualifier = source[index] === "'" || source[runEnd] === '!';
      const token = formulaReferenceAt(source, index, qualifier);
      if (token) {
        yield token;
        index = token.endIndex;
        continue;
      }
    }
    index = ignoredEnd(source, index);
  }
}
