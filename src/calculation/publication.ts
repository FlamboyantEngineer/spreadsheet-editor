import type { Workbook } from '../model/workbook.js';
import { assertWorkbook } from '../model/workbook.js';
import type { CalculationResult } from './evaluator.js';
import type { JsonValue } from './values.js';
import { assertJsonValue } from './values.js';
import { assertCalculationResult } from './http.js';
import { calculationFormats, worksheetCalculationExtent } from '../model/extent.js';

export type CalculationCellState = 'pending' | 'ready' | 'unavailable';
export interface ExternalCalculationSnapshot {
  inputKey: string;
  revision: number;
  status: CalculationCellState;
  result?: CalculationResult;
  cellStates?: Record<string, CalculationCellState>;
  hostContext?: JsonValue;
  error?: string;
}
export const calculationCellKey = (sheetId: string, x: number, y: number) =>
  JSON.stringify([sheetId, x, y]);

/** Sorted object fields make equivalent JSON inputs share an identity across transports. */
export function orderedJson(value: unknown): string {
  if (value === undefined) return 'null';
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(orderedJson).join(',')}]`;
  return `{${Object.keys(value)
    .sort()
    .filter((key) => (value as Record<string, unknown>)[key] !== undefined)
    .map((key) => `${JSON.stringify(key)}:${orderedJson((value as Record<string, unknown>)[key])}`)
    .join(',')}}`;
}
/** A deterministic freshness token. This is deliberately not an authorization signature. */
export function workbookCalculationKey(workbook: Workbook): string {
  assertWorkbook(workbook);
  return orderedJson(
    workbook.worksheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      data: sheet.data,
      inputValues: sheet.inputValues,
      extent: worksheetCalculationExtent(sheet),
      rowIds: sheet.rowIds,
      columnIds: sheet.columnIds,
      formats: calculationFormats(sheet),
    })),
  );
}
const states = new Set(['pending', 'ready', 'unavailable']);
function assertTransportTree(value: unknown, ancestors = new Set<object>(), depth = 0): void {
  if (depth > 64) throw new TypeError('Calculation publication is nested too deeply.');
  if (
    value === undefined ||
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  )
    return;
  if (typeof value === 'number' && Number.isFinite(value)) return;
  if (typeof value !== 'object' || ancestors.has(value))
    throw new TypeError('Calculation publication must contain finite acyclic JSON.');
  if (!Array.isArray(value) && ![Object.prototype, null].includes(Object.getPrototypeOf(value)))
    throw new TypeError('Calculation publication must contain plain JSON objects.');
  ancestors.add(value);
  for (const child of Object.values(value)) assertTransportTree(child, ancestors, depth + 1);
  ancestors.delete(value);
}
export function assertExternalCalculationSnapshot(
  value: unknown,
  workbook: Workbook,
): asserts value is ExternalCalculationSnapshot {
  if (!value || typeof value !== 'object')
    throw new TypeError('Invalid external calculation snapshot.');
  if (Object.getPrototypeOf(value) !== Object.prototype && Object.getPrototypeOf(value) !== null)
    throw new TypeError('External calculation snapshots must be plain objects.');
  assertTransportTree(value);
  if (new TextEncoder().encode(JSON.stringify(value)).length > 16 * 1024 * 1024)
    throw new RangeError('External calculation snapshot exceeds the transport byte limit.');
  const snapshot = value as ExternalCalculationSnapshot;
  if (snapshot.inputKey !== workbookCalculationKey(workbook))
    throw new TypeError('External calculation input key does not match the workbook.');
  if (
    !Number.isSafeInteger(snapshot.revision) ||
    snapshot.revision < 0 ||
    !states.has(snapshot.status)
  )
    throw new TypeError('Invalid calculation revision or status.');
  if (
    snapshot.error !== undefined &&
    (typeof snapshot.error !== 'string' || snapshot.error.length > 8192)
  )
    throw new TypeError('Invalid external calculation error.');
  if (snapshot.hostContext !== undefined) assertJsonValue(snapshot.hostContext);
  if (snapshot.cellStates !== undefined) {
    if (
      !snapshot.cellStates ||
      typeof snapshot.cellStates !== 'object' ||
      Array.isArray(snapshot.cellStates) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(snapshot.cellStates))
    )
      throw new TypeError('Invalid calculation cell states.');
    if (Object.keys(snapshot.cellStates).length > 250_000)
      throw new RangeError('Too many calculation cell states.');
    for (const [key, state] of Object.entries(snapshot.cellStates)) {
      const tuple: unknown = JSON.parse(key);
      if (
        !Array.isArray(tuple) ||
        tuple.length !== 3 ||
        typeof tuple[0] !== 'string' ||
        !Number.isSafeInteger(tuple[1]) ||
        !Number.isSafeInteger(tuple[2]) ||
        tuple[1] < 0 ||
        tuple[2] < 0 ||
        tuple[1] > 16_383 ||
        tuple[2] > 1_048_575 ||
        calculationCellKey(tuple[0], tuple[1], tuple[2]) !== key ||
        !workbook.worksheets.some((sheet) => sheet.id === tuple[0]) ||
        !states.has(state)
      )
        throw new TypeError('Invalid calculation cell state.');
    }
  }
  if (snapshot.status === 'ready' && !snapshot.result)
    throw new TypeError('Ready calculations require a result.');
  if (snapshot.result) assertCalculationResult(snapshot.result, workbook);
  for (const sheet of workbook.worksheets)
    sheet.data.forEach((row, y) =>
      row.forEach((input, x) => {
        const state = snapshot.cellStates?.[calculationCellKey(sheet.id, x, y)] ?? snapshot.status;
        const result = snapshot.result?.worksheets.find((entry) => entry.id === sheet.id)?.values[
          y
        ]?.[x];
        if (state === 'ready' && input !== null && input !== '' && result == null)
          throw new TypeError('Ready authored cells require calculated values.');
        if (state !== 'ready' && result != null)
          throw new TypeError('Pending or unavailable cells cannot publish cached values.');
      }),
    );
}

interface DependencyIndex {
  /** Cells pending whatever was edited: every cell without recorded dependencies. */
  unrecorded: string[];
  /** Cells that read the used extent of a worksheet, by worksheet id. */
  extentReaders: [string, string[]][];
  /** Reverse dependencies: the cells that read each cell. */
  dependents: Map<string, string[]>;
}
const dependencyIndexes = new WeakMap<CalculationResult, DependencyIndex>();
/** Built once per result, so a keystroke walks only what its draft cell reaches (SE-PERF-1). */
function dependencyIndex(result: CalculationResult): DependencyIndex {
  const cached = dependencyIndexes.get(result);
  if (cached) return cached;
  const recorded = result.dependencies?.cells,
    unrecorded: string[] = [],
    dependents = new Map<string, string[]>();
  for (const sheet of result.worksheets)
    sheet.values.forEach((row, y) =>
      row.forEach((_cell, x) => {
        const key = calculationCellKey(sheet.id, x, y);
        if (!recorded || !Object.hasOwn(recorded, key)) unrecorded.push(key);
      }),
    );
  for (const [key, dependencies] of Object.entries(recorded ?? {}))
    for (const dependency of dependencies) {
      const list = dependents.get(dependency);
      if (list) list.push(key);
      else dependents.set(dependency, [key]);
    }
  const index = {
    unrecorded,
    extentReaders: Object.entries(result.dependencies?.extents ?? {}),
    dependents,
  };
  dependencyIndexes.set(result, index);
  return index;
}

/** Invalidates only formulas whose actually-read dependencies lead to an edited cell. */
export function pendingCalculationCells(
  result: CalculationResult | null,
  edited: readonly string[],
  extentChanges: readonly string[] = [],
): Set<string> {
  const pending = new Set(edited);
  if (!result) return pending;
  const { unrecorded, extentReaders, dependents } = dependencyIndex(result);
  for (const key of unrecorded) pending.add(key);
  if (!result.dependencies) return pending;
  for (const [key, sheets] of extentReaders)
    if (sheets.some((sheet) => extentChanges.includes(sheet))) pending.add(key);
  const queue = [...pending];
  for (let index = 0; index < queue.length; index++)
    for (const dependent of dependents.get(queue[index]) ?? [])
      if (!pending.has(dependent)) {
        pending.add(dependent);
        queue.push(dependent);
      }
  return pending;
}
