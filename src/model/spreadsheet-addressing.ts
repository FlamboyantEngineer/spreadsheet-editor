/**
 * @file Canonical A1 spreadsheet addressing helpers: column-name <-> index math, `A1` cell-address
 * parse/format and range parse/normalize. Dependency-free, so the model, calculation and browser
 * modules all share it.
 *
 * Coordinate conventions: `x` is the 0-based column index, `y` is the 0-based row index; formatted
 * `A1` rows are 1-based. Parsers accept lowercase input, strip every `$` anchor, and reject row `0`,
 * leading-zero rows, and malformed column names by returning `null`.
 */

/**
 * An unquoted worksheet name inside a formula: a letter in any script or `_`, then letters, digits, `_`
 * or spaces. Use with the `u` flag.
 */
export const UNQUOTED_SHEET_NAME = String.raw`[\p{L}_][\p{L}\p{N}_ ]*`;
/** A character that continues a formula word, so a reference cannot start right after it. */
export const FORMULA_WORD_CHARACTER = /[\p{L}\p{N}_.]/u;

/** Parsed spreadsheet `A1` address coordinates (0-based). */
interface SpreadsheetAddress {
  x: number;
  y: number;
}

/** A parsed, normalized spreadsheet range (0-based, start <= end). */
export interface SpreadsheetRange {
  start: SpreadsheetAddress;
  end: SpreadsheetAddress;
  address: string;
}

/**
 * Parses a spreadsheet column name (`A`, `AB`, ...) into a 0-based index, or
 * `null` when malformed. Input is trimmed and case-normalized.
 * @param columnName - The column name.
 * @returns The 0-based column index, or `null`.
 */
export function parseSpreadsheetColumnName(columnName: unknown): number | null {
  const normalized = String(columnName || '')
    .trim()
    .toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return null;
  let index = 0;
  for (const char of normalized) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1;
}

/**
 * Formats a 0-based column index into a spreadsheet column name. Negative and
 * non-numeric input falls back to `A`.
 * @param index - The 0-based column index.
 * @returns The column name (`A`, `AB`, ...).
 */
export function formatSpreadsheetColumnName(index: unknown): string {
  let value = Number(index) + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label || 'A';
}

/**
 * Formats `(x, y)` cell coordinates into an `A1` spreadsheet address.
 * @param x - The 0-based column index.
 * @param y - The 0-based row index.
 * @returns The `A1` address.
 */
export function formatSpreadsheetAddress(x: number, y: number): string {
  return `${formatSpreadsheetColumnName(x)}${y + 1}`;
}

/**
 * Parses an `A1` spreadsheet address (ignoring `$` anchors) into coordinates,
 * or `null`.
 * @param value - The address.
 * @returns The coordinates, or `null`.
 */
export function parseSpreadsheetAddress(value: unknown): SpreadsheetAddress | null {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(
    String(value || '')
      .trim()
      .replace(/\$/g, '')
      .toUpperCase(),
  );
  if (!match) return null;
  const x = parseSpreadsheetColumnName(match[1]);
  const y = Number.parseInt(match[2], 10) - 1;
  if (x === null || !Number.isInteger(x) || x < 0 || !Number.isInteger(y) || y < 0) return null;
  return { x, y };
}

/**
 * Parses an `A1` (or `A1:B2`) spreadsheet range into normalized coordinates.
 * @param value - The raw range value.
 * @returns The normalized range, or `null` when invalid.
 */
export function parseSpreadsheetRange(value: unknown): SpreadsheetRange | null {
  const normalized = String(value || '')
    .trim()
    .replace(/\$/g, '')
    .toUpperCase();
  if (!normalized || !/^[A-Z]+[1-9]\d*(?::[A-Z]+[1-9]\d*)?$/.test(normalized)) return null;
  const [startValue, endValue = startValue] = normalized.split(':');
  const start = parseSpreadsheetAddress(startValue);
  const end = parseSpreadsheetAddress(endValue);
  if (!start || !end) return null;
  const x1 = Math.min(start.x, end.x);
  const y1 = Math.min(start.y, end.y);
  const x2 = Math.max(start.x, end.x);
  const y2 = Math.max(start.y, end.y);
  return {
    start: { x: x1, y: y1 },
    end: { x: x2, y: y2 },
    address:
      x1 === x2 && y1 === y2
        ? formatSpreadsheetAddress(x1, y1)
        : `${formatSpreadsheetAddress(x1, y1)}:${formatSpreadsheetAddress(x2, y2)}`,
  };
}
