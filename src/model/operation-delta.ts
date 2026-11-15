import { orderedJson } from '../calculation/publication.js';
import type { Workbook, Worksheet } from './workbook.js';

export class WorkbookOperationConflict extends Error {
  constructor(public readonly path: string) {
    super(`The workbook changed concurrently at ${path}.`);
    this.name = 'WorkbookOperationConflict';
  }
}

/** Equality by ordered JSON, so key order never makes two authored values differ. */
export const equal = (a: unknown, b: unknown) => orderedJson(a) === orderedJson(b);

/** The state an operation saw before it ran, and the state it left. */
export interface DeltaSides {
  before: unknown;
  after: unknown;
}

type Entity = Record<string, unknown> & { id: string };
type Fields = Record<string, unknown>;
/** One three-way merge: the current value and the operation's sides, at a conflict path. */
interface Merge<Current = unknown, Side = Current> {
  current: Current;
  before: Side;
  after: Side;
  path: string;
}

const object = (value: unknown): value is Fields =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const optionalObject = (value: unknown): value is Fields | undefined =>
  value === undefined || object(value);
const entities = (value: unknown[]): value is Entity[] =>
  value.every((item) => object(item) && typeof item.id === 'string');
const rational = (value: unknown) =>
  object(value) && ('numerator' in value || 'denominator' in value);
const listMerge = (merge: Merge): merge is Merge<unknown[]> =>
  [merge.current, merge.before, merge.after].every((value) => Array.isArray(value));
const objectMerge = (merge: Merge): merge is Merge<Fields, Fields | undefined> =>
  optionalObject(merge.before) && optionalObject(merge.after) && object(merge.current);

/** A list one side lacks merges as an empty list when the current value is a list. */
function withEmptyLists(merge: Merge): Merge {
  if (!Array.isArray(merge.current)) return merge;
  return {
    ...merge,
    before: merge.before === undefined && Array.isArray(merge.after) ? [] : merge.before,
    after: merge.after === undefined && Array.isArray(merge.before) ? [] : merge.after,
  };
}

/** Removes or updates one entity the operation saw, in the merged list. */
function changeEntity(
  result: Entity[],
  { previous, next, path }: { previous: Entity; next: Entity | undefined; path: string },
) {
  const index = result.findIndex((item) => item.id === previous.id);
  const at = `${path}/${previous.id}`;
  if (!next) {
    if (index < 0) return;
    if (!equal(result[index], previous)) throw new WorkbookOperationConflict(at);
    result.splice(index, 1);
  } else if (!equal(previous, next)) {
    if (index < 0) throw new WorkbookOperationConflict(at);
    result[index] = delta(result[index], { before: previous, after: next }, at) as Entity;
  }
}

/** Adds one entity the operation introduced, unless the same entity is already there. */
function addEntity(result: Entity[], next: Entity, path: string) {
  const existing = result.find((item) => item.id === next.id);
  if (existing && !equal(existing, next)) throw new WorkbookOperationConflict(`${path}/${next.id}`);
  if (!existing) result.push(structuredClone(next));
}

function mergeEntities({ current, before, after, path }: Merge<Entity[]>): Entity[] {
  const result = structuredClone(current);
  for (const previous of before)
    changeEntity(result, { previous, next: after.find((item) => item.id === previous.id), path });
  for (const next of after)
    if (!before.some((item) => item.id === next.id)) addEntity(result, next, path);
  return result;
}

function mergeLists({ current, before, after, path }: Merge<unknown[]>): unknown {
  if (entities(before) && entities(after) && entities(current))
    return mergeEntities({ current, before, after, path });
  if ([...before, ...after, ...current].every((item) => typeof item === 'string'))
    return [
      ...new Set([
        ...current.filter((item) => !before.includes(item) || after.includes(item)),
        ...after.filter((item) => !before.includes(item)),
      ]),
    ].sort();
  // Sparse geometry arrays retain unrelated column properties at their stable current indexes.
  if (before.length === after.length && current.length === before.length)
    return current.map((value, index) =>
      delta(value, { before: before[index], after: after[index] }, `${path}/${index}`),
    );
  throw new WorkbookOperationConflict(path);
}

function mergeObjects({ current, before, after, path }: Merge<Fields, Fields | undefined>) {
  const updated = after ?? {};
  const result = structuredClone(current);
  for (const key of new Set([...Object.keys(before ?? {}), ...Object.keys(updated)])) {
    const value = delta(
      current[key],
      { before: before?.[key], after: updated[key] },
      `${path}/${key}`,
    );
    if (value === undefined) delete result[key];
    else result[key] = value;
  }
  return after === undefined && !Object.keys(result).length ? undefined : result;
}

/** Three-way authored deltas: arrays of entities merge by ID; primitive sets by membership. */
export function delta(current: unknown, sides: DeltaSides, path: string): unknown {
  if (equal(sides.before, sides.after)) return structuredClone(current);
  if (equal(current, sides.before) || equal(current, sides.after))
    return structuredClone(sides.after);
  const merge = withEmptyLists({ current, before: sides.before, after: sides.after, path });
  if (rational(merge.before)) throw new WorkbookOperationConflict(path);
  if (listMerge(merge)) return mergeLists(merge);
  if (objectMerge(merge)) return mergeObjects(merge);
  throw new WorkbookOperationConflict(path);
}

export function sheetOf(workbook: Workbook, id: string): Worksheet {
  const sheet = workbook.worksheets.find((sheet) => sheet.id === id);
  if (!sheet) throw new WorkbookOperationConflict(`worksheets/${id}`);
  return sheet;
}
