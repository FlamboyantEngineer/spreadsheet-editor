import { requireDefined } from './require-defined.js';
import {
  formatSpreadsheetAddress,
  formatSpreadsheetColumnName,
  FORMULA_WORD_CHARACTER,
  parseSpreadsheetAddress,
  parseSpreadsheetColumnName,
  parseSpreadsheetRange,
  UNQUOTED_SHEET_NAME,
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
const CELL_OR_LINES = String.raw`(\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![\p{L}\p{N}_.(])`;
/** Sticky patterns: one with an optional sheet prefix, one without for runs that end in no `!`. */
const QUALIFIED_REFERENCE = new RegExp(
    String.raw`(?:(?:'((?:[^']|'')*)'|(${UNQUOTED_SHEET_NAME}))!)?` + CELL_OR_LINES,
    'uy',
  ),
  LOCAL_REFERENCE = new RegExp(`()()${CELL_OR_LINES}`, 'uy'),
  NAME_RUN = /[\p{L}\p{N}_ ]*/uy;

type ReferenceMap = (reference: string, sheet: string | null, prefix: string) => string;

/** The end of the run of name characters starting at `index`. */
function nameRunEnd(formula: string, index: number): number {
  NAME_RUN.lastIndex = index;
  NAME_RUN.exec(formula);
  return Math.max(NAME_RUN.lastIndex, index + 1);
}

/** The reference at `index`; a sheet prefix is tried only where the name run ends in `!`. */
function referenceAt(formula: string, index: number, runEnd: number): RegExpExecArray | null {
  const pattern =
    formula[index] === "'" || formula[runEnd] === '!' ? QUALIFIED_REFERENCE : LOCAL_REFERENCE;
  pattern.lastIndex = index;
  return pattern.exec(formula);
}

function mapped(match: RegExpExecArray, map: ReferenceMap): string {
  const sheet = match[1] ? match[1].replace(/''/g, "'") : match[2] || null;
  const prefix = match[0].slice(0, match[0].length - match[3].length);
  return map(match[3], sheet, prefix);
}

/** The index just past the quote that closes text opened before `index`; doubled quotes stay in. */
function quotedEnd(formula: string, index: number, quote: string): number {
  let cursor = index;
  while (cursor < formula.length) {
    const next = formula[cursor++];
    if (next !== quote) continue;
    if (formula[cursor] === quote) cursor++;
    else break;
  }
  return cursor;
}

/** Scan references while keeping quoted text, anchors, and all authored numeric digits intact. */
export function mapReferences(formula: string, map: ReferenceMap): string {
  if (!formula.trimStart().startsWith('=')) return formula;
  let output = '',
    index = 0,
    runEnd = -1;
  while (index < formula.length) {
    if (!FORMULA_WORD_CHARACTER.test(formula[index - 1] ?? '')) {
      // An unquoted sheet name runs to the end of its run of name characters, so the run's end is
      // found once per run; that keeps the scan linear in the formula length.
      if (index >= runEnd) runEnd = nameRunEnd(formula, index);
      const match = referenceAt(formula, index, runEnd);
      if (match) {
        output += mapped(match, map);
        index += match[0].length;
        continue;
      }
    }
    const character = formula[index++];
    const end =
      character === '"' || character === "'" ? quotedEnd(formula, index, character) : index;
    output += character + formula.slice(index, end);
    index = end;
  }
  return output;
}

const CELL_PATTERN = /^(\$?)([A-Za-z]+)(\$?)([1-9]\d*)$/;

/** A cell or cell range after a change, keeping its anchors. */
function rewriteCells(
  pieces: string[],
  [firstCell, lastCell]: RegExpExecArray[],
  change: StructureChange,
): string {
  const position = (match: RegExpExecArray) =>
    change.axis === 'row'
      ? Number(match[4]) - 1
      : requireDefined(parseSpreadsheetColumnName(match[2]));
  const transformed = interval(position(firstCell), position(lastCell), change);
  if (!transformed) return '#REF!';
  const format = (match: RegExpExecArray, value: number) =>
    change.axis === 'row'
      ? `${match[1]}${match[2]}${match[3]}${value + 1}`
      : `${match[1]}${formatSpreadsheetColumnName(value)}${match[3]}${match[4]}`;
  return pieces.length === 1
    ? format(firstCell, transformed[0])
    : `${format(firstCell, transformed[0])}:${format(lastCell, transformed[1])}`;
}

function rewriteReference(reference: string, change: StructureChange): string {
  const pieces = reference.split(':'),
    first = pieces[0],
    last = pieces[1] ?? first;
  const firstCell = CELL_PATTERN.exec(first),
    lastCell = CELL_PATTERN.exec(last);
  if (firstCell && lastCell) return rewriteCells(pieces, [firstCell, lastCell], change);
  const wholeColumns = /^\$?[A-Za-z]+$/.test(first);
  if ((wholeColumns && change.axis !== 'column') || (!wholeColumns && change.axis !== 'row'))
    return reference;
  const position = (token: string) =>
    wholeColumns
      ? requireDefined(parseSpreadsheetColumnName(token.replace('$', '')))
      : Number(token.replace('$', '')) - 1;
  const transformed = interval(position(first), position(last), change);
  if (!transformed) return '#REF!';
  const format = (original: string, value: number) =>
    `${original.startsWith('$') ? '$' : ''}${wholeColumns ? formatSpreadsheetColumnName(value) : value + 1}`;
  return `${format(first, transformed[0])}:${format(last, transformed[1])}`;
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
    if (after < 0 || after > (kind === 'column' ? 16383 : 1048575)) return null;
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
    const pieces = reference.split(':').map(piece);
    return pieces.some((part) => part === null) ? '#REF!' : prefix + pieces.join(':');
  });
}
