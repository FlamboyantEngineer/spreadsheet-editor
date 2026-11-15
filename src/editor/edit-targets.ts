import { requireDefined } from '../model/require-defined.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { copyWorkbook, type Workbook } from '../model/workbook.js';

/** Cell addresses a command wrote, by worksheet id. */
export type CellTargets = Map<string, Iterable<string>>;
export function mergeTargets(a: CellTargets, b: CellTargets): CellTargets {
  const merged: CellTargets = new Map();
  for (const source of [a, b])
    for (const [id, addresses] of source)
      merged.set(id, new Set([...(merged.get(id) ?? []), ...addresses]));
  return merged;
}
export function blockTargets(
  sheetId: string,
  x: number,
  y: number,
  values: unknown[][],
): CellTargets {
  return new Map([
    [
      sheetId,
      values.flatMap((row, dy) => row.map((_cell, dx) => formatSpreadsheetAddress(x + dx, y + dy))),
    ],
  ]);
}
/** A detached, deeply frozen copy of `workbook`, the baseline a host may read (SE-PERF-2.1). */
export function immutableBaseline(workbook: Workbook): Workbook {
  const snapshot = copyWorkbook(workbook),
    pending: object[] = [snapshot];
  while (pending.length) {
    const value = requireDefined(pending.pop());
    if (Object.isFrozen(value)) continue;
    for (const child of Object.values(value))
      if (child && typeof child === 'object') pending.push(child);
    Object.freeze(value);
  }
  return snapshot;
}
