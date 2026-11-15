import { delta, equal, sheetOf, WorkbookOperationConflict } from './operation-delta.js';
import {
  type AuthoredCellState,
  authoredCellState,
  type WorkbookCellChange,
  type WorkbookOperation,
  type WorksheetGeometry,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
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

const CELL_FIELDS = [
  ['style', 'style'],
  ['formats', 'format'],
  ['inputValues', 'inputValue'],
  ['cellProvenance', 'provenance'],
] as const;

function writeState(sheet: Worksheet, address: string, state: AuthoredCellState): void {
  const cell = requireDefined(parseSpreadsheetAddress(address));
  while (sheet.data.length <= cell.y) sheet.data.push([]);
  while (sheet.data[cell.y].length <= cell.x) sheet.data[cell.y].push(null);
  sheet.data[cell.y][cell.x] = state.value;
  for (const [field, key] of CELL_FIELDS) {
    const value = state[key];
    if (value === undefined) delete sheet[field]?.[address];
    else
      Object.assign((sheet[field] ??= {}) as Record<string, unknown>, {
        [address]: structuredClone(value),
      });
  }
}

/** Where a change lands now: its row and column identities win over its recorded address. */
function changeAddress(sheet: Worksheet, change: WorkbookCellChange, sheetId: string): string {
  const position = requireDefined(parseSpreadsheetAddress(change.address));
  const y = change.rowId ? (sheet.rowIds?.indexOf(change.rowId) ?? -1) : position.y;
  const x = change.columnId ? (sheet.columnIds?.indexOf(change.columnId) ?? -1) : position.x;
  if (x < 0 || y < 0) throw new WorkbookOperationConflict(`cells/${sheetId}/${change.address}`);
  return formatSpreadsheetAddress(x, y);
}

function applyCells(state: ApplyState, operation: Operation<'cells'>) {
  const sheet = sheetOf(state.next, operation.sheetId);
  for (const change of operation.changes) {
    const address = changeAddress(sheet, change, operation.sheetId);
    const merged = delta(authoredCellState(sheet, address), change, `cells/${sheet.id}/${address}`);
    writeState(sheet, address, merged as AuthoredCellState);
  }
}

/** Replays the line command on the recorded layout, for the geometry and merge patches after it. */
function recordLineLayout(state: ApplyState, sheet: Worksheet, operation: Operation<'lines'>) {
  if (!operation.beforeGeometry) return;
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
  if (operation.ids.some((id) => ids.includes(id)))
    throw new WorkbookOperationConflict(`lines/${operation.sheetId}`);
  const after = operation.afterId ? ids.indexOf(operation.afterId) : -1;
  const before = operation.beforeId ? ids.indexOf(operation.beforeId) : -1;
  const index =
    after >= 0 ? after : before >= 0 ? before + 1 : Math.min(operation.index, ids.length);
  state.next = changeWorkbookStructure(state.next, operation.sheetId, {
    axis: operation.axis,
    kind: 'insert',
    index,
    count: operation.ids.length,
    ids: operation.ids,
  });
}

/** Deletes the lines that still exist, from the last one back. */
function deleteLines(state: ApplyState, ids: string[], operation: Operation<'lines'>) {
  const positions = operation.ids
    .map((id) => ids.indexOf(id))
    .filter((index) => index >= 0)
    .sort((a, b) => b - a);
  for (const index of positions)
    state.next = changeWorkbookStructure(state.next, operation.sheetId, {
      axis: operation.axis,
      kind: 'delete',
      index,
      count: 1,
    });
}

function applyLines(state: ApplyState, operation: Operation<'lines'>) {
  const sheet = sheetOf(state.next, operation.sheetId);
  const ids = operation.axis === 'row' ? sheet.rowIds : sheet.columnIds;
  recordLineLayout(state, sheet, operation);
  if (!ids) throw new WorkbookOperationConflict(`lines/${sheet.id}`);
  if (operation.action === 'insert') insertLines(state, ids, operation);
  else deleteLines(state, ids, operation);
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
  // A list neither side had stays absent, so an inverse restores the workbook exactly.
  for (const key of ['comments', 'suggestions'] as const)
    if (next[key] !== undefined || operation.after[key] !== undefined || value[key]?.length)
      next[key] = value[key] as never;
}

function applyMerges(state: ApplyState, operation: Operation<'merges'>) {
  const sheet = sheetOf(state.next, operation.sheetId);
  const layout = state.lineLayouts.get(sheet.id);
  const before =
    layout && equal(operation.before, layout.beforeMerges)
      ? (layout.merges ?? {})
      : operation.before;
  sheet.mergeCells = delta(
    sheet.mergeCells ?? {},
    { before, after: operation.after },
    `merges/${sheet.id}`,
  ) as Worksheet['mergeCells'];
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
  Object.assign(sheet, delta(geometry, { before, after: operation.after }, `geometry/${sheet.id}`));
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

function applyOperation(state: ApplyState, operation: WorkbookOperation) {
  switch (operation.kind) {
    case 'cells':
      return applyCells(state, operation);
    case 'lines':
      return applyLines(state, operation);
    case 'annotations':
      return applyAnnotations(state, operation);
    case 'geometry':
      return applyGeometry(state, operation);
    case 'merges':
      return applyMerges(state, operation);
    case 'workbook-metadata':
      return void Object.assign(
        state.next,
        delta({ name: state.next.name }, operation, 'workbook'),
      );
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
