import { MAX_COLUMN_INDEX, MAX_ROW_INDEX } from '../model/spreadsheet-addressing.js';

export const calculationCellKey = (sheetId: string, x: number, y: number) =>
  JSON.stringify([sheetId, x, y]);

const inRange = (value: unknown, max: number): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 && value <= max;

/** The canonical wire key of a cell; alternate spellings cannot be used in dependency maps. */
export function parseCalculationCellKey(key: unknown): [string, number, number] | null {
  if (typeof key !== 'string') return null;
  let tuple: unknown;
  try {
    tuple = JSON.parse(key);
  } catch {
    return null;
  }
  if (!Array.isArray(tuple) || tuple.length !== 3) return null;
  const [sheetId, x, y] = tuple;
  return typeof sheetId === 'string' &&
    inRange(x, MAX_COLUMN_INDEX) &&
    inRange(y, MAX_ROW_INDEX) &&
    calculationCellKey(sheetId, x, y) === key
    ? [sheetId, x, y]
    : null;
}
