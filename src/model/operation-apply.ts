import { applyCells, lineIndex } from './operation-cells.js';
import { delta, equal, sheetOf, WorkbookOperationConflict } from './operation-delta.js';
import type { LinePreparation, WorkbookOperation, WorksheetGeometry } from './operations.js';
import { requireDefined } from './require-defined.js';
import { changeWorkbookStructure } from './structure.js';
import { assertWorkbook, copyWorkbook, type Workbook, type Worksheet } from './workbook.js';

type Operation<Kind extends WorkbookOperation['kind']> = Extract<WorkbookOperation, { kind: Kind }>;
/** The layout a line command left, so a later geometry or merge patch in its frame applies. */
interface LineLayout {
  beforeGeometry?: WorksheetGeometry;
  geometry: WorksheetGeometry;
  beforeMerges?: Worksheet['mergeCells'];
  merges: Worksheet['mergeCells'];
}
/** The workbook being built and the line layouts recorded so far. */
interface ApplyState {
  next: Workbook;
  lineLayouts: Map<string, LineLayout>;
}

function applyLineFrame(sheet: Worksheet, frame: LinePreparation) {
  const fields = ['minDimensions', 'rowIds', 'columnIds'] as const;
  const next = structuredClone(frame.after);
  for (const field of fields) {
    if (!equal(sheet[field], frame.before[field]))
      throw new WorkbookOperationConflict(`line-extent/${sheet.id}/${field}`);
    if (next[field] === undefined) delete sheet[field];
    else Object.assign(sheet, { [field]: next[field] });
  }
}

/** Replays the line command on the recorded layout, for the geometry and merge patches after it. */
function recordLineLayout(state: ApplyState, sheet: Worksheet, operation: Operation<'lines'>) {
  if (!operation.restoreLayout || !operation.beforeGeometry) return;
  const frame: Workbook = {
    schemaVersion: 1,
    worksheets: [
      {
        id: sheet.id,
        name: sheet.name,
        data: [[]],
        ...structuredClone(operation.beforeGeometry),
        mergeCells: structuredClone(operation.beforeMerges),
      },
    ],
  };
  const transformed = changeWorkbookStructure(frame, sheet.id, {
    axis: operation.axis,
    kind: operation.action,
    index: operation.index,
    count: operation.count,
    ...(operation.action === 'insert' ? { ids: operation.ids } : {}),
  }).worksheets[0];
  state.lineLayouts.set(sheet.id, {
    beforeGeometry: operation.beforeGeometry,
    geometry: {
      minDimensions: transformed.minDimensions,
      columns: transformed.columns,
      rows: transformed.rows,
      freezeColumns: transformed.freezeColumns,
    },
    beforeMerges: operation.beforeMerges,
    merges: transformed.mergeCells,
  });
}

/** Inserts the lines beside their recorded neighbours, else at the recorded index. */
function insertLines(state: ApplyState, ids: string[], operation: Operation<'lines'>) {
  const positions = requireDefined(lineIndex(ids));
  if (operation.ids.some((id) => positions.has(id)))
    throw new WorkbookOperationConflict(`lines/${operation.sheetId}`);
  if (operation.extentOnly) {
    for (const id of operation.ids) ids.push(id);
    return;
  }
  const after = operation.afterId ? positions.get(operation.afterId) : undefined;
  const before = operation.beforeId ? positions.get(operation.beforeId) : undefined;
  const index =
    after ?? (before === undefined ? Math.min(operation.index, ids.length) : before + 1);
  state.next = changeWorkbookStructure(state.next, operation.sheetId, {
    axis: operation.axis,
    kind: 'insert',
    index,
    count: operation.ids.length,
    ids: operation.ids,
  });
}

/** Deletes contiguous runs from the end to avoid copying the workbook for each line. */
function deleteLines(state: ApplyState, ids: string[], operation: Operation<'lines'>) {
  const targets = new Set(operation.ids);
  if (operation.extentOnly) {
    const remaining = ids.filter((id) => !targets.has(id));
    if (!remaining.every((id, index) => id === ids[index]))
      throw new WorkbookOperationConflict(`line-extent/${operation.sheetId}`);
    ids.length = remaining.length;
    return;
  }
  let end = ids.length;
  while (end > 0) {
    if (!targets.has(ids[end - 1])) {
      end--;
      continue;
    }
    let index = end - 1;
    while (index > 0 && targets.has(ids[index - 1])) index--;
    state.next = changeWorkbookStructure(state.next, operation.sheetId, {
      axis: operation.axis,
      kind: 'delete',
      index,
      count: end - index,
    });
    end = index;
  }
}

function applyLines(state: ApplyState, operation: Operation<'lines'>) {
  const sheet = sheetOf(state.next, operation.sheetId);
  if (operation.preparation) applyLineFrame(sheet, operation.preparation);
  const ids = operation.axis === 'row' ? sheet.rowIds : sheet.columnIds;
  const rowCount = sheet.data.length;
  recordLineLayout(state, sheet, operation);
  if (!ids) throw new WorkbookOperationConflict(`lines/${sheet.id}`);
  if (operation.action === 'insert') insertLines(state, ids, operation);
  else deleteLines(state, ids, operation);
  if (operation.dataRows?.before === rowCount) {
    const { data } = sheetOf(state.next, sheet.id);
    while (data.length < operation.dataRows.after) data.push([]);
  }
}

const reviewLists = (value: Pick<Workbook, 'comments' | 'suggestions'>) => ({
  comments: value.comments ?? [],
  suggestions: value.suggestions ?? [],
});

function applyAnnotations({ next }: ApplyState, operation: Operation<'annotations'>) {
  const value = delta(
    reviewLists(next),
    { before: reviewLists(operation.before), after: reviewLists(operation.after) },
    'annotations',
  ) as Pick<Workbook, 'comments' | 'suggestions'>;
  for (const key of ['comments', 'suggestions'] as const) {
    if (operation.after[key] !== undefined || value[key]?.length) next[key] = value[key] as never;
    else if (operation.before[key] !== undefined) delete next[key];
  }
}

function applyMerges(state: ApplyState, operation: Operation<'merges'>) {
  const sheet = sheetOf(state.next, operation.sheetId);
  const layout = state.lineLayouts.get(sheet.id);
  const before =
    layout && equal(operation.before, layout.beforeMerges)
      ? (layout.merges ?? {})
      : operation.before;
  const merges = delta(
    sheet.mergeCells ?? {},
    { before: before ?? {}, after: operation.after ?? {} },
    `merges/${sheet.id}`,
  ) as NonNullable<Worksheet['mergeCells']>;
  if (operation.after || Object.keys(merges).length) sheet.mergeCells = merges;
  else delete sheet.mergeCells;
}

function applyGeometry(state: ApplyState, operation: Operation<'geometry'>) {
  const sheet = sheetOf(state.next, operation.sheetId);
  const layout = state.lineLayouts.get(sheet.id);
  const geometry = {
    minDimensions: sheet.minDimensions,
    columns: sheet.columns,
    rows: sheet.rows,
    freezeColumns: sheet.freezeColumns,
  };
  const before =
    layout && equal(operation.before, layout.beforeGeometry) ? layout.geometry : operation.before;
  const merged = delta(
    geometry,
    { before, after: operation.after },
    `geometry/${sheet.id}`,
  ) as WorksheetGeometry;
  for (const field of ['minDimensions', 'columns', 'rows', 'freezeColumns'] as const) {
    if (merged[field] === undefined) delete sheet[field];
    else Object.assign(sheet, { [field]: merged[field] });
  }
}

function applyRename({ next }: ApplyState, operation: Operation<'sheet-rename'>) {
  const sheet = sheetOf(next, operation.sheetId);
  sheet.name = delta(sheet.name, operation, `worksheets/${sheet.id}/name`) as string;
}

/** Adds the sheet before the first sheet that follows it in the recorded order and still exists. */
function applySheetAdd({ next }: ApplyState, operation: Operation<'sheet-add'>) {
  if (next.worksheets.some((sheet) => sheet.id === operation.sheet.id))
    throw new WorkbookOperationConflict(`worksheets/${operation.sheet.id}`);
  const following = operation.afterOrder
    .slice(operation.afterOrder.indexOf(operation.sheet.id) + 1)
    .find((id) => next.worksheets.some((sheet) => sheet.id === id));
  const index = following
    ? next.worksheets.findIndex((sheet) => sheet.id === following)
    : next.worksheets.length;
  next.worksheets.splice(index, 0, structuredClone(operation.sheet));
}

function applySheetDelete({ next }: ApplyState, operation: Operation<'sheet-delete'>) {
  const sheet = sheetOf(next, operation.sheet.id);
  if (!equal(sheet, operation.sheet)) throw new WorkbookOperationConflict(`worksheets/${sheet.id}`);
  next.worksheets = next.worksheets.filter((item) => item.id !== sheet.id);
}

/** Reorders the sheets both orders know, leaving sheets added since in their places. */
function applyReorder({ next }: ApplyState, operation: Operation<'sheet-reorder'>) {
  const known = next.worksheets
    .filter((sheet) => operation.before.includes(sheet.id))
    .map((sheet) => sheet.id);
  if (
    !equal(
      known,
      operation.before.filter((id) => known.includes(id)),
    )
  )
    throw new WorkbookOperationConflict('worksheet-order');
  const reordered = operation.after
    .filter((id) => known.includes(id))
    .map((id) => sheetOf(next, id));
  next.worksheets = next.worksheets.map((sheet) =>
    known.includes(sheet.id) ? requireDefined(reordered.shift()) : sheet,
  );
}

function applyReplace(state: ApplyState, operation: Operation<'replace'>) {
  if (!equal(state.next, operation.before)) throw new WorkbookOperationConflict('workbook');
  state.next = copyWorkbook(operation.after);
}

function applyName(workbook: Workbook, operation: Operation<'workbook-metadata'>) {
  const { name } = delta({ name: workbook.name }, operation, 'workbook') as Pick<Workbook, 'name'>;
  if (name === undefined) delete workbook.name;
  else workbook.name = name;
}

function applyOperation(state: ApplyState, operation: WorkbookOperation) {
  switch (operation.kind) {
    case 'cells':
      return applyCells(sheetOf(state.next, operation.sheetId), operation);
    case 'lines':
      return applyLines(state, operation);
    case 'line-frame':
      return applyLineFrame(sheetOf(state.next, operation.sheetId), operation);
    case 'annotations':
      return applyAnnotations(state, operation);
    case 'geometry':
      return applyGeometry(state, operation);
    case 'merges':
      return applyMerges(state, operation);
    case 'workbook-metadata':
      return applyName(state.next, operation);
    case 'sheet-rename':
      return applyRename(state, operation);
    case 'sheet-add':
      return applySheetAdd(state, operation);
    case 'sheet-delete':
      return applySheetDelete(state, operation);
    case 'sheet-reorder':
      return applyReorder(state, operation);
    case 'replace':
      return applyReplace(state, operation);
  }
}

/**
 * Apply operations in an already serialized coordinate frame, then validate atomically.
 * Use applyWorkbookTransaction for stale transactions: its baseline checks structural causality.
 * This is a serialized operation application layer, independent of transports and calculators.
 * Throws before changing the caller's workbook; the host commits the returned snapshot once.
 */
export function applyWorkbookOperations(
  current: Workbook,
  operations: readonly WorkbookOperation[],
): Workbook {
  assertWorkbook(current);
  const state: ApplyState = { next: copyWorkbook(current), lineLayouts: new Map() };
  for (const operation of operations) applyOperation(state, operation);
  assertWorkbook(state.next);
  return state.next;
}
