import { authoredAddresses, authoredCellState, equal } from './cell-operations.js';
import {
  createWorkbookEntityId,
  materializeWorksheetExtent,
  type WorkbookIdFactory,
  type WorksheetExtent,
  worksheetExtent,
} from './extent.js';
import type {
  WorkbookCellChange,
  WorkbookCommandResult,
  WorkbookOperation,
  WorksheetGeometry,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import { changeWorkbookStructure, type StructureChange } from './structure.js';
import { copyWorkbook, type Workbook, type Worksheet } from './workbook.js';

type LinesOperation = Extract<WorkbookOperation, { kind: 'lines' }>;

export function annotationOperation(before: Workbook, after: Workbook): WorkbookOperation {
  return {
    kind: 'annotations',
    before: structuredClone({ comments: before.comments, suggestions: before.suggestions }),
    after: structuredClone({ comments: after.comments, suggestions: after.suggestions }),
  };
}
const geometryOf = (sheet: Worksheet): WorksheetGeometry =>
  structuredClone({
    minDimensions: sheet.minDimensions,
    columns: sheet.columns,
    rows: sheet.rows,
    freezeColumns: sheet.freezeColumns,
  });
export function geometryOperation(before: Worksheet, after: Worksheet): WorkbookOperation {
  return {
    kind: 'geometry',
    sheetId: before.id,
    before: geometryOf(before),
    after: geometryOf(after),
  };
}
export function mergeOperation(before: Worksheet, after: Worksheet): WorkbookOperation {
  return {
    kind: 'merges',
    sheetId: before.id,
    before: structuredClone(before.mergeCells ?? {}),
    after: structuredClone(after.mergeCells ?? {}),
  };
}
/** Which worksheet command a sheet operation records, and the worksheet it applies to. */
export interface SheetCommand {
  kind: 'sheet-add' | 'sheet-delete' | 'sheet-rename' | 'sheet-reorder';
  /** The added, deleted or renamed worksheet; a reorder names none. */
  sheetId?: string | undefined;
}
export function sheetOperation(
  { kind, sheetId }: SheetCommand,
  before: Workbook,
  after: Workbook,
): WorkbookOperation {
  const beforeOrder = before.worksheets.map((sheet) => sheet.id),
    afterOrder = after.worksheets.map((sheet) => sheet.id);
  if (kind === 'sheet-reorder') return { kind, before: beforeOrder, after: afterOrder };
  const old = before.worksheets.find((sheet) => sheet.id === sheetId),
    next = after.worksheets.find((sheet) => sheet.id === sheetId);
  if (kind === 'sheet-rename') {
    if (!old || !next) throw new RangeError('Unknown worksheet.');
    return { kind, sheetId: old.id, before: old.name, after: next.name };
  }
  const sheet = kind === 'sheet-add' ? next : old;
  if (!sheet) throw new RangeError('Unknown worksheet.');
  return {
    kind,
    sheet: structuredClone(sheet),
    index: (kind === 'sheet-add' ? afterOrder : beforeOrder).indexOf(sheet.id),
    beforeOrder,
    afterOrder,
  };
}

/** Gives every line of a sheet that tracks identities one, so the command can name its lines. */
function identifyLines(sheet: Worksheet, extent: WorksheetExtent, idFactory: WorkbookIdFactory) {
  if (sheet.rowIds) while (sheet.rowIds.length < extent.rows) sheet.rowIds.push(idFactory('row'));
  if (sheet.columnIds)
    while (sheet.columnIds.length < extent.columns) sheet.columnIds.push(idFactory('column'));
}

/** The ids of the lines a change inserts or deletes, starting identities when the sheet has none. */
function changedLineIds(
  sheet: Worksheet,
  { change, count, extentSize }: { change: StructureChange; count: number; extentSize: number },
  idFactory: WorkbookIdFactory,
) {
  let oldIds = change.axis === 'row' ? sheet.rowIds : sheet.columnIds;
  const ids =
    change.kind === 'insert'
      ? [...(change.ids ?? Array.from({ length: count }, () => idFactory(change.axis)))]
      : (oldIds?.slice(change.index, change.index + count) ?? []);
  if (!oldIds) {
    oldIds = Array.from({ length: extentSize }, () => idFactory(change.axis));
    if (change.axis === 'row') sheet.rowIds = oldIds;
    else sheet.columnIds = oldIds;
    if (change.kind === 'delete') ids.push(...oldIds.slice(change.index, change.index + count));
  }
  return { ids, oldIds };
}

/** The authored cells a deletion removes, so undo can restore them. */
function deletedCells(
  sheet: Worksheet,
  change: StructureChange,
  count: number,
): WorkbookCellChange[] {
  return [...authoredAddresses(sheet)].flatMap((address) => {
    const cell = requireDefined(parseSpreadsheetAddress(address)),
      coordinate = change.axis === 'row' ? cell.y : cell.x;
    if (coordinate < change.index || coordinate >= change.index + count) return [];
    return [
      {
        address,
        rowId: sheet.rowIds?.[cell.y],
        columnId: sheet.columnIds?.[cell.x],
        before: authoredCellState(sheet, address),
        after: { value: null },
      },
    ];
  });
}

/** Where a surviving cell of the changed sheet moves; null for a deleted one. */
function movedAddress(
  address: string,
  { change, count }: { change: StructureChange; count: number },
): string | null {
  const position = requireDefined(parseSpreadsheetAddress(address));
  let coordinate = change.axis === 'row' ? position.y : position.x;
  if (change.kind === 'delete' && coordinate >= change.index && coordinate < change.index + count)
    return null;
  if (coordinate >= change.index) coordinate += change.kind === 'insert' ? count : -count;
  return formatSpreadsheetAddress(
    change.axis === 'column' ? coordinate : position.x,
    change.axis === 'row' ? coordinate : position.y,
  );
}

/** Cell changes besides the move itself, such as rewritten formulas, by structural mapping. */
function survivingChanges(
  previousSheet: Worksheet,
  currentSheet: Worksheet,
  move: { sheetId: string; change: StructureChange; count: number },
): WorkbookCellChange[] {
  const changes: WorkbookCellChange[] = [];
  for (const address of authoredAddresses(previousSheet)) {
    const destination = previousSheet.id === move.sheetId ? movedAddress(address, move) : address;
    if (destination === null) continue;
    const before = authoredCellState(previousSheet, address),
      after = authoredCellState(currentSheet, destination);
    if (equal(before, after)) continue;
    const position = requireDefined(parseSpreadsheetAddress(address));
    changes.push({
      address: destination,
      beforeAddress: address,
      rowId: previousSheet.rowIds?.[position.y],
      columnId: previousSheet.columnIds?.[position.x],
      before,
      after,
    });
  }
  return changes;
}

/** Where a structural command applies, and who makes the ids of lines it inserts. */
export interface StructureTarget {
  sheetId: string;
  idFactory?: WorkbookIdFactory | undefined;
}

/** A structural operation is recorded directly, including on identical blank rows. */
export function changeWorkbookStructureWithOperations(
  workbook: Workbook,
  { sheetId, idFactory = createWorkbookEntityId }: StructureTarget,
  change: StructureChange,
): WorkbookCommandResult {
  const materialized = materializeWorksheetExtent(workbook, sheetId),
    sheet = requireDefined(materialized.worksheets.find((sheet) => sheet.id === sheetId));
  const extent = worksheetExtent(sheet, materialized);
  identifyLines(sheet, extent, idFactory);
  const extentSize = change.axis === 'row' ? extent.rows : extent.columns;
  const count = change.kind === 'delete' ? Math.min(change.count, extentSize - 1) : change.count;
  if (!count) return { workbook: copyWorkbook(workbook), operations: [] };
  const { ids, oldIds } = changedLineIds(sheet, { change, count, extentSize }, idFactory);
  const actual = { ...change, count, ...(change.kind === 'insert' ? { ids } : {}) };
  const next = changeWorkbookStructure(materialized, sheetId, actual);
  const nextSheet = requireDefined(next.worksheets.find((candidate) => candidate.id === sheetId));
  const operation: LinesOperation = {
    kind: 'lines',
    sheetId,
    axis: change.axis,
    action: change.kind,
    index: change.index,
    count,
    ids,
    beforeId: oldIds[change.index - 1],
    afterId: oldIds[change.kind === 'insert' ? change.index : change.index + count],
    beforeGeometry: geometryOf(sheet),
    afterGeometry: geometryOf(nextSheet),
    beforeMerges: structuredClone(sheet.mergeCells ?? {}),
    afterMerges: structuredClone(nextSheet.mergeCells ?? {}),
    ...(change.kind === 'delete' ? { deletedCells: deletedCells(sheet, change, count) } : {}),
  };
  // Compare surviving cells by their structural mapping, so moving a cell is not emitted as a second write.
  const operations: WorkbookOperation[] = [operation];
  for (const previousSheet of workbook.worksheets) {
    const currentSheet = requireDefined(
      next.worksheets.find((candidate) => candidate.id === previousSheet.id),
    );
    const changes = survivingChanges(previousSheet, currentSheet, { sheetId, change, count });
    if (changes.length) operations.push({ kind: 'cells', sheetId: previousSheet.id, changes });
  }
  if (!equal(workbook.comments, next.comments) || !equal(workbook.suggestions, next.suggestions))
    operations.push(annotationOperation(workbook, next));
  return { workbook: next, operations };
}
