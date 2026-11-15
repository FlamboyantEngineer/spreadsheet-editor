import { calculationCellKey } from '../calculation/publication.js';

export const NO_CELLS: ReadonlySet<string> = new Set();
/** Ready cells less a draft's pending cells, without copying the ready set per keystroke. */
export function withoutPending(
  ready: ReadonlySet<string>,
  pending: ReadonlySet<string>,
): ReadonlySet<string> {
  if (!pending.size) return ready;
  const view: ReadonlySet<string> = {
    get size() {
      return [...view].length;
    },
    has: (key) => ready.has(key) && !pending.has(key),
    *[Symbol.iterator]() {
      for (const key of ready) if (!pending.has(key)) yield key;
    },
    *keys() {
      yield* view;
    },
    *values() {
      yield* view;
    },
    *entries() {
      for (const key of view) yield [key, key];
    },
    forEach(callback, thisArg) {
      for (const key of view) callback.call(thisArg, key, key, view);
    },
  } as ReadonlySet<string>;
  return view;
}
export function readyValues<T>(
  values: readonly (readonly T[])[] | undefined,
  ready: ReadonlySet<string>,
  sheetId: string,
): (T | null)[][] | undefined {
  return values?.map((row, y) =>
    row.map((value, x) => (ready.has(calculationCellKey(sheetId, x, y)) ? value : null)),
  );
}
/**
 * Blank a draft's pending cells; rows without one keep their identity. A row with a pending cell
 * is always a new row, even when that cell is already blank: the grid does not paint a row whose
 * identity is unchanged, so it would miss the cell's change of readiness (SE-PERF-1).
 */
export function withoutPendingValues<T>(
  values: (T | null)[][] | undefined,
  pending: ReadonlySet<string>,
  sheetId: string,
): (T | null)[][] | undefined {
  if (!values || !pending.size) return values;
  let next: (T | null)[][] | undefined;
  for (const key of pending) {
    const [id, x, y] = JSON.parse(key) as [string, number, number];
    if (id !== sheetId || !values[y]) continue;
    next ??= [...values];
    if (next[y] === values[y]) next[y] = [...values[y]];
    if (x < next[y].length) next[y][x] = null;
  }
  return next ?? values;
}
