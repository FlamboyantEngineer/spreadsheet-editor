import type {
  FormulaReferenceHighlight,
  FormulaReferenceTarget,
} from './formula-reference-parser.js';

// The cell, range and whole-line references a formula's text may name, read at one position.

export type ParsedReferenceCandidate = Omit<FormulaReferenceHighlight, 'id' | 'colorIndex'>;

function parseColumnName(value: string): number | null {
  const normalizedValue = value.replace(/\$/g, '').toUpperCase();
  if (!/^[A-Z]+$/.test(normalizedValue)) return null;
  let column = 0;
  for (const character of normalizedValue) {
    column = column * 26 + (character.charCodeAt(0) - 64);
  }
  return column > 0 ? column - 1 : null;
}

function formatColumnName(index: number): string {
  let value = Number(index) + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label || 'A';
}

function formatCellKey(column: number, row: number): string {
  return `${formatColumnName(column)}${row + 1}`;
}

function parseRowNumber(value: string): number | null {
  const normalizedValue = value.replace(/\$/g, '');
  if (!/^[1-9]\d*$/.test(normalizedValue)) return null;
  const row = Number.parseInt(normalizedValue, 10) - 1;
  return Number.isSafeInteger(row) && row >= 0 ? row : null;
}

function parseCellAddressText(value: string): { column: number; row: number } | null {
  const match = /^(\$?[A-Za-z]+)(\$?[1-9]\d*)$/.exec(value.trim());
  if (!match) return null;
  const column = parseColumnName(match[1] || '');
  const row = parseRowNumber(match[2] || '');
  if (column == null || row == null) return null;
  return { column, row };
}

function parseCellAddressAt(
  source: string,
  index: number,
): {
  column: number;
  endIndex: number;
  originalText: string;
  row: number;
} | null {
  const match = /^\$?[A-Za-z]+\$?[1-9]\d*/.exec(source.slice(index));
  if (!match) return null;
  const parsedAddress = parseCellAddressText(match[0]);
  if (!parsedAddress) return null;
  return {
    column: parsedAddress.column,
    endIndex: index + match[0].length,
    originalText: match[0],
    row: parsedAddress.row,
  };
}

type Cell = { column: number; row: number };

function normalizeCellRange(
  start: Cell,
  end: Cell,
): Pick<
  Extract<FormulaReferenceTarget, { kind: 'range' }>,
  'startColumn' | 'startRow' | 'endColumn' | 'endRow'
> {
  return {
    startColumn: Math.min(start.column, end.column),
    startRow: Math.min(start.row, end.row),
    endColumn: Math.max(start.column, end.column),
    endRow: Math.max(start.row, end.row),
  };
}

function buildCellReferenceCandidate(
  source: string,
  index: number,
): ParsedReferenceCandidate | null {
  const start = parseCellAddressAt(source, index);
  if (!start) return null;
  if (source[start.endIndex] === ':') {
    const end = parseCellAddressAt(source, start.endIndex + 1);
    if (end) {
      const target = {
        kind: 'range',
        ...normalizeCellRange(start, end),
      } satisfies FormulaReferenceTarget;
      return {
        originalText: source.slice(index, end.endIndex),
        normalizedKey: `range:${formatCellKey(target.startColumn, target.startRow)}:${formatCellKey(target.endColumn, target.endRow)}`,
        startIndex: index,
        endIndex: end.endIndex,
        target,
      };
    }
  }
  const target = {
    kind: 'cell',
    startColumn: start.column,
    startRow: start.row,
    endColumn: start.column,
    endRow: start.row,
  } satisfies FormulaReferenceTarget;
  return {
    originalText: start.originalText,
    normalizedKey: `cell:${formatCellKey(start.column, start.row)}`,
    startIndex: index,
    endIndex: start.endIndex,
    target,
  };
}

function buildColumnRangeCandidate(source: string, index: number): ParsedReferenceCandidate | null {
  const match = /^(\$?[A-Za-z]+):(\$?[A-Za-z]+)/.exec(source.slice(index));
  if (!match) return null;
  const startColumn = parseColumnName(match[1] || '');
  const endColumn = parseColumnName(match[2] || '');
  if (startColumn == null || endColumn == null) return null;
  const normalizedStart = Math.min(startColumn, endColumn);
  const normalizedEnd = Math.max(startColumn, endColumn);
  const target = {
    kind: 'column-range',
    startColumn: normalizedStart,
    endColumn: normalizedEnd,
  } satisfies FormulaReferenceTarget;
  return {
    originalText: match[0],
    normalizedKey: `column:${formatColumnName(normalizedStart)}:${formatColumnName(normalizedEnd)}`,
    startIndex: index,
    endIndex: index + match[0].length,
    target,
  };
}

function buildRowRangeCandidate(source: string, index: number): ParsedReferenceCandidate | null {
  const match = /^(\$?[1-9]\d*):(\$?[1-9]\d*)/.exec(source.slice(index));
  if (!match) return null;
  const startRow = parseRowNumber(match[1] || '');
  const endRow = parseRowNumber(match[2] || '');
  if (startRow == null || endRow == null) return null;
  const normalizedStart = Math.min(startRow, endRow);
  const normalizedEnd = Math.max(startRow, endRow);
  const target = {
    kind: 'row-range',
    startRow: normalizedStart,
    endRow: normalizedEnd,
  } satisfies FormulaReferenceTarget;
  return {
    originalText: match[0],
    normalizedKey: `row:${normalizedStart + 1}:${normalizedEnd + 1}`,
    startIndex: index,
    endIndex: index + match[0].length,
    target,
  };
}

export function buildReferenceCandidate(
  source: string,
  index: number,
): ParsedReferenceCandidate | null {
  return (
    buildCellReferenceCandidate(source, index) ||
    buildColumnRangeCandidate(source, index) ||
    buildRowRangeCandidate(source, index)
  );
}
