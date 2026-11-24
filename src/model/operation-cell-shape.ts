import type { Worksheet } from './workbook.js';

const metadataFields = ['style', 'formats', 'inputValues', 'cellProvenance'] as const;
type MetadataField = (typeof metadataFields)[number];
interface Change<T> {
  before: T;
  after: T;
}
interface RowLength extends Change<number> {
  row: number;
  id?: string | undefined;
}
/** Blank data slots and empty maps that authored cell states alone cannot represent. */
export interface CellOperationShape {
  count?: Change<number> | undefined;
  lengths: RowLength[];
  maps: (Change<boolean> & { field: MetadataField })[];
}

export function collectCellShape(
  before: Worksheet,
  after: Worksheet,
  { rows, addresses }: { rows: Iterable<number>; addresses?: ReadonlySet<string> | undefined },
): CellOperationShape | undefined {
  const count =
    before.data.length === after.data.length
      ? undefined
      : { before: before.data.length, after: after.data.length };
  const lengths: RowLength[] = [];
  for (const row of rows) {
    const previous = before.data[row]?.length ?? 0;
    const next = after.data[row]?.length ?? 0;
    if (previous !== next)
      lengths.push({
        row,
        id: after.rowIds?.[row] ?? before.rowIds?.[row],
        before: previous,
        after: next,
      });
  }
  const maps = metadataFields.flatMap((field) => {
    if (Boolean(before[field]) === Boolean(after[field])) return [];
    const keys = Object.keys(before[field] ?? after[field] ?? {});
    if (addresses && keys.length && !keys.some((key) => addresses.has(key))) return [];
    return [{ field, before: Boolean(before[field]), after: Boolean(after[field]) }];
  });
  return count || lengths.length || maps.length ? { count, lengths, maps } : undefined;
}

const swap = <T, Entry extends Change<T>>(change: Entry): Entry => ({
  ...change,
  before: change.after,
  after: change.before,
});

export function invertCellShape(shape: CellOperationShape): CellOperationShape {
  return {
    count: shape.count && swap(shape.count),
    lengths: shape.lengths.map((length) => swap(length)),
    maps: shape.maps.map((map) => swap(map)),
  };
}

function restoreRowLength(
  data: Worksheet['data'],
  { index, after }: { index: number; after: number },
) {
  if (after > 0) while (data.length <= index) data.push([]);
  const row = data[index];
  if (!row) return;
  if (row.length > after && row.slice(after).every((value) => value === null)) row.length = after;
  while (row.length < after) row.push(null);
}

function restoreRowCount(data: Worksheet['data'], count: number) {
  if (data.length > count && data.slice(count).every((row) => !row.length)) data.length = count;
  while (data.length < count) data.push([]);
}

/** Capture the current frame first so undo never trims another writer's stored cells. */
export function restoreCellShape(
  sheet: Worksheet,
  shape: CellOperationShape,
  rowIndex: ReadonlyMap<string, number> | undefined,
): () => void {
  const count = shape.count?.before === sheet.data.length ? shape.count.after : undefined;
  const lengths = shape.lengths.flatMap(({ row, id, before, after }) => {
    const index = id === undefined ? row : rowIndex?.get(id);
    return index !== undefined && (sheet.data[index]?.length ?? 0) === before
      ? [{ index, after }]
      : [];
  });
  return () => {
    for (const length of lengths) restoreRowLength(sheet.data, length);
    if (count !== undefined) restoreRowCount(sheet.data, count);
    for (const { field, after } of shape.maps) {
      if (after) sheet[field] ??= {};
      else if (sheet[field] && !Object.keys(sheet[field]).length) delete sheet[field];
    }
  };
}
