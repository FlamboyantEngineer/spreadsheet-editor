import type { AuthoredCellState } from './operations.js';
import { requireDefined } from './require-defined.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { Worksheet } from './workbook.js';

const metadataFields = [
  ['style', 'style'],
  ['formats', 'format'],
  ['inputValues', 'inputValue'],
  ['cellProvenance', 'provenance'],
] as const;

/** Writes authored state without allocating data for cells that carry only presentation. */
export function writeAuthoredCellState(
  sheet: Worksheet,
  address: string,
  state: AuthoredCellState & { materialize?: boolean | undefined },
) {
  const { x, y } = requireDefined(parseSpreadsheetAddress(address));
  if (
    state.materialize ||
    state.value !== null ||
    state.inputValue !== undefined ||
    x < (sheet.data[y]?.length ?? 0)
  ) {
    while (sheet.data.length <= y) sheet.data.push([]);
    while (sheet.data[y].length <= x) sheet.data[y].push(null);
    sheet.data[y][x] = state.value;
  }
  for (const [field, key] of metadataFields) {
    const value = state[key];
    if (value === undefined) delete sheet[field]?.[address];
    else ((sheet[field] ??= {}) as Record<string, unknown>)[address] = structuredClone(value);
  }
}
