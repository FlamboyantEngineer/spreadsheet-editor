import { FORMULA_WORD_CHARACTER } from './spreadsheet-addressing.js';
export type FormulaReferenceTarget =
  | {
      kind: 'cell';
      startColumn: number;
      startRow: number;
      endColumn: number;
      endRow: number;
    }
  | {
      kind: 'range';
      startColumn: number;
      startRow: number;
      endColumn: number;
      endRow: number;
    }
  | {
      kind: 'column-range';
      startColumn: number;
      endColumn: number;
    }
  | {
      kind: 'row-range';
      startRow: number;
      endRow: number;
    };

export type FormulaReferenceHighlight = {
  id: string;
  originalText: string;
  normalizedKey: string;
  startIndex: number;
  endIndex: number;
  target: FormulaReferenceTarget;
  colorIndex: number;
};

export type FormulaReferenceTokenSegment = {
  color?: string;
  colorIndex?: number;
  referenceKey?: string;
  text: string;
};

type ParsedReferenceCandidate = Omit<FormulaReferenceHighlight, 'id' | 'colorIndex'>;

const FORMULA_REFERENCE_BOUNDARY_PATTERN = FORMULA_WORD_CHARACTER;
const FORMULA_REFERENCE_TOKEN_STOP_PATTERN = /[\s,;()+\-*/^&=<>{}]/;

export const FORMULA_REFERENCE_HIGHLIGHT_COLORS = [
  '#2563eb',
  '#dc2626',
  '#059669',
  '#7c3aed',
  '#d97706',
  '#0891b2',
  '#be123c',
  '#4d7c0f',
] as const;

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

function normalizeCellRange(
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Pick<
  Extract<FormulaReferenceTarget, { kind: 'range' }>,
  'startColumn' | 'startRow' | 'endColumn' | 'endRow'
> {
  return {
    startColumn: Math.min(startColumn, endColumn),
    startRow: Math.min(startRow, endRow),
    endColumn: Math.max(startColumn, endColumn),
    endRow: Math.max(startRow, endRow),
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
        ...normalizeCellRange(start.column, start.row, end.column, end.row),
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

function buildReferenceCandidate(source: string, index: number): ParsedReferenceCandidate | null {
  return (
    buildCellReferenceCandidate(source, index) ||
    buildColumnRangeCandidate(source, index) ||
    buildRowRangeCandidate(source, index)
  );
}

function skipQuotedText(source: string, index: number, quote: string): number {
  let cursor = index + 1;
  while (cursor < source.length) {
    if (source[cursor] === '\\' && cursor + 1 < source.length) {
      cursor += 2;
      continue;
    }
    if (source[cursor] === quote && source[cursor + 1] === quote) {
      cursor += 2;
      continue;
    }
    if (source[cursor] === quote) return cursor + 1;
    cursor += 1;
  }
  return source.length;
}

function skipBracketedText(source: string, index: number): number {
  let depth = 1;
  let cursor = index + 1;
  while (cursor < source.length) {
    const character = source[cursor];
    if (character === '"' || character === "'") {
      cursor = skipQuotedText(source, cursor, character);
      continue;
    }
    if (character === '[') depth += 1;
    if (character === ']') {
      depth -= 1;
      if (depth <= 0) return cursor + 1;
    }
    cursor += 1;
  }
  return source.length;
}

function hasReferenceBoundaryBefore(source: string, index: number): boolean {
  const previous = index > 0 ? source[index - 1] : '';
  return !previous || !FORMULA_REFERENCE_BOUNDARY_PATTERN.test(previous);
}

function hasReferenceBoundaryAfter(source: string, index: number): boolean {
  const next = source[index] || '';
  return !next || !FORMULA_REFERENCE_BOUNDARY_PATTERN.test(next);
}

function hasCrossSheetBangBeforeReference(source: string, index: number): boolean {
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    const character = source[cursor] || '';
    if (character === '!') return true;
    if (FORMULA_REFERENCE_TOKEN_STOP_PATTERN.test(character)) return false;
  }
  return false;
}

function shouldIgnoreCandidate(source: string, candidate: ParsedReferenceCandidate): boolean {
  if (!hasReferenceBoundaryBefore(source, candidate.startIndex)) return true;
  if (!hasReferenceBoundaryAfter(source, candidate.endIndex)) return true;
  if (hasCrossSheetBangBeforeReference(source, candidate.startIndex)) return true;
  const nextCharacter = source[candidate.endIndex] || '';
  if (
    nextCharacter === '!' ||
    nextCharacter === '(' ||
    nextCharacter === '[' ||
    nextCharacter === ']'
  )
    return true;
  return false;
}

export function parseFormulaReferenceHighlights(formulaText: string): FormulaReferenceHighlight[] {
  const source = String(formulaText ?? '');
  const colorByNormalizedKey = new Map<string, number>();
  const references: FormulaReferenceHighlight[] = [];
  let index = 0;

  while (index < source.length) {
    const character = source[index];
    if (character === '"' || character === "'") {
      index = skipQuotedText(source, index, character);
      continue;
    }
    if (character === '[') {
      index = skipBracketedText(source, index);
      continue;
    }

    const candidate = buildReferenceCandidate(source, index);
    if (!candidate) {
      index += 1;
      continue;
    }
    if (shouldIgnoreCandidate(source, candidate)) {
      index = Math.max(candidate.endIndex, index + 1);
      continue;
    }
    if (!colorByNormalizedKey.has(candidate.normalizedKey)) {
      colorByNormalizedKey.set(candidate.normalizedKey, colorByNormalizedKey.size);
    }
    const colorIndex = colorByNormalizedKey.get(candidate.normalizedKey) || 0;
    references.push({
      ...candidate,
      id: `${candidate.normalizedKey}:${candidate.startIndex}-${candidate.endIndex}`,
      colorIndex,
    });
    index = candidate.endIndex;
  }

  return references;
}

export function buildFormulaReferenceTokenSegments(
  formulaText: string,
  references: readonly FormulaReferenceHighlight[],
): FormulaReferenceTokenSegment[] {
  const segments: FormulaReferenceTokenSegment[] = [];
  let cursor = 0;
  references
    .slice()
    .sort((left, right) => left.startIndex - right.startIndex || left.endIndex - right.endIndex)
    .forEach((reference) => {
      if (reference.startIndex < cursor || reference.endIndex <= reference.startIndex) return;
      if (reference.startIndex > cursor) {
        segments.push({ text: formulaText.slice(cursor, reference.startIndex) });
      }
      segments.push({
        text: formulaText.slice(reference.startIndex, reference.endIndex),
        referenceKey: reference.normalizedKey,
        colorIndex: reference.colorIndex,
        color:
          FORMULA_REFERENCE_HIGHLIGHT_COLORS[
            reference.colorIndex % FORMULA_REFERENCE_HIGHLIGHT_COLORS.length
          ] || FORMULA_REFERENCE_HIGHLIGHT_COLORS[0],
      });
      cursor = reference.endIndex;
    });
  if (cursor < formulaText.length) segments.push({ text: formulaText.slice(cursor) });
  return segments;
}
