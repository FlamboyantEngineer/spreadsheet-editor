import { formulaReferenceTokens } from './formula-reference-token.js';
import { requireDefined } from './require-defined.js';
import {
  formatSpreadsheetAddress,
  formatSpreadsheetColumnName,
  MAX_COLUMN_INDEX,
  MAX_ROW_INDEX,
  parseSpreadsheetAddress,
  parseSpreadsheetColumnName,
  parseSpreadsheetRange,
} from './spreadsheet-addressing.js';
import type { StructureChange } from './structure.js';
import { worksheetNameKey } from './workbook.js';

/** Where a line interval lands after a change; null when a deletion removes all of it. */
export function interval(
  start: number,
  end: number,
  change: StructureChange,
): [number, number] | null {
  const reversed = start > end,
    left = Math.min(start, end),
    right = Math.max(start, end);
  let a = left,
    b = right;
  if (change.kind === 'insert') {
    if (a >= change.index) a += change.count;
    if (b >= change.index) b += change.count;
  } else {
    const last = change.index + change.count - 1;
    if (a > last) a -= change.count;
    else if (a >= change.index) a = change.index;
    if (b > last) b -= change.count;
    else if (b >= change.index) b = change.index - 1;
    if (a > b) return null;
  }
  return reversed ? [b, a] : [a, b];
}
type ReferenceMap = (reference: string, sheet: string | null, prefix: string) => string;

/** Scan references while keeping quoted text, anchors, and all authored numeric digits intact. */
export function mapReferences(formula: string, map: ReferenceMap): string {
  if (!formula.trimStart().startsWith('=')) return formula;
  let output = '';
  let cursor = 0;
  for (const {
    startIndex,
    addressIndex,
    endIndex,
    sheetName,
    coordinates,
  } of formulaReferenceTokens(formula)) {
    if (!coordinates) continue;
    const prefix = formula.slice(startIndex, addressIndex);
    const mapped = map(formula.slice(addressIndex, endIndex), sheetName, prefix);
    output += formula.slice(cursor, startIndex) + mapped;
    cursor = endIndex;
  }
  return output + formula.slice(cursor);
}

const RANGE_SEPARATOR = /(\s*:\s*)/;
const lastIndex = (axis: 'column' | 'row') =>
  axis === 'column' ? MAX_COLUMN_INDEX : MAX_ROW_INDEX;

const CELL_PATTERN = /^(\$?)([A-Za-z]+)(\$?)([1-9]\d*)$/;

/** A cell or cell range after a change, keeping its anchors. */
function rewriteCells(
  separator: string | undefined,
  [firstCell, lastCell]: RegExpExecArray[],
  change: StructureChange,
): string {
  const position = (match: RegExpExecArray) =>
    change.axis === 'row'
      ? Number(match[4]) - 1
      : requireDefined(parseSpreadsheetColumnName(match[2]));
  const transformed = interval(position(firstCell), position(lastCell), change);
  if (!transformed || transformed.some((index) => index > lastIndex(change.axis))) return '#REF!';
  const format = (match: RegExpExecArray, value: number) =>
    change.axis === 'row'
      ? `${match[1]}${match[2]}${match[3]}${value + 1}`
      : `${match[1]}${formatSpreadsheetColumnName(value)}${match[3]}${match[4]}`;
  return separator
    ? `${format(firstCell, transformed[0])}${separator}${format(lastCell, transformed[1])}`
    : format(firstCell, transformed[0]);
}

function rewriteReference(reference: string, change: StructureChange): string {
  const [first, separator, last = first] = reference.split(RANGE_SEPARATOR);
  const firstCell = CELL_PATTERN.exec(first),
    lastCell = CELL_PATTERN.exec(last);
  if (firstCell && lastCell) return rewriteCells(separator, [firstCell, lastCell], change);
  const wholeColumns = /^\$?[A-Za-z]+$/.test(first);
  if ((wholeColumns && change.axis !== 'column') || (!wholeColumns && change.axis !== 'row'))
    return reference;
  const position = (token: string) =>
    wholeColumns
      ? requireDefined(parseSpreadsheetColumnName(token.replace('$', '')))
      : Number(token.replace('$', '')) - 1;
  const transformed = interval(position(first), position(last), change);
  if (!transformed || transformed.some((index) => index > lastIndex(change.axis))) return '#REF!';
  const format = (original: string, value: number) =>
    `${original.startsWith('$') ? '$' : ''}${wholeColumns ? formatSpreadsheetColumnName(value) : value + 1}`;
  return `${format(first, transformed[0])}${separator}${format(last, transformed[1])}`;
}

/** The sheet a formula is in, and the sheet a structural change applies to. */
export interface StructureSheets {
  currentSheet: string;
  targetSheet: string;
}

export function rewriteStructureReferences(
  formula: string,
  change: StructureChange,
  { currentSheet, targetSheet }: StructureSheets,
): string {
  return mapReferences(formula, (reference, sheet, prefix) => {
    if (worksheetNameKey(sheet ?? currentSheet) !== worksheetNameKey(targetSheet))
      return prefix + reference;
    const rewritten = rewriteReference(reference, change);
    return rewritten === '#REF!' ? rewritten : prefix + rewritten;
  });
}

export function moveAddress(address: string, change: StructureChange): string | null {
  const cell = parseSpreadsheetAddress(address);
  if (!cell) return null;
  const index = change.axis === 'row' ? cell.y : cell.x,
    next = interval(index, index, change);
  return next
    ? formatSpreadsheetAddress(
        change.axis === 'column' ? next[0] : cell.x,
        change.axis === 'row' ? next[0] : cell.y,
      )
    : null;
}
export function moveRange(range: string, change: StructureChange): string | null {
  const parsed = parseSpreadsheetRange(range);
  if (!parsed) return null;
  const start = change.axis === 'row' ? parsed.start.y : parsed.start.x,
    end = change.axis === 'row' ? parsed.end.y : parsed.end.x;
  const moved = interval(start, end, change);
  if (!moved) return null;
  return `${formatSpreadsheetAddress(change.axis === 'column' ? moved[0] : parsed.start.x, change.axis === 'row' ? moved[0] : parsed.start.y)}:${formatSpreadsheetAddress(change.axis === 'column' ? moved[1] : parsed.end.x, change.axis === 'row' ? moved[1] : parsed.end.y)}`;
}

/** Translate a copied formula without touching literals, sheet names or absolute axes. */
export function translateFormulaReferences(formula: string, dx: number, dy: number): string {
  const axis = (token: string, kind: 'column' | 'row'): string | null => {
    const anchored = token.startsWith('$');
    const bare = token.replace('$', '');
    const before = kind === 'column' ? parseSpreadsheetColumnName(bare) : Number(bare) - 1;
    if (before === null) return null;
    const after = before + (anchored ? 0 : kind === 'column' ? dx : dy);
    if (after < 0 || after > lastIndex(kind)) return null;
    if (after === before) return token;
    return (
      (anchored ? '$' : '') +
      (kind === 'column' ? formatSpreadsheetColumnName(after) : String(after + 1))
    );
  };
  const piece = (token: string) => {
    const cell = /^(\$?[A-Za-z]+)(\$?[1-9]\d*)$/.exec(token);
    if (!cell) return axis(token, /[A-Za-z]/.test(token) ? 'column' : 'row');
    const x = axis(cell[1], 'column'),
      y = axis(cell[2], 'row');
    return x === null || y === null ? null : x + y;
  };
  return mapReferences(formula, (reference, _sheet, prefix) => {
    const [first, separator, last = first] = reference.split(RANGE_SEPARATOR);
    const start = piece(first),
      end = piece(last);
    if (start === null || end === null) return '#REF!';
    return prefix + start + (separator ? separator + end : '');
  });
}
