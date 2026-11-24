import { invertCellShape } from './operation-cell-shape.js';
import type { WorkbookOperation } from './operations.js';
import { moveReviewReferences } from './review-reference-move.js';

type CellsOperation = Extract<WorkbookOperation, { kind: 'cells' }>;
type LinesOperation = Extract<WorkbookOperation, { kind: 'lines' }>;
type InsertedLines = Map<string, { row: Set<string>; column: Set<string> }>;

/** The ids of every line the operations insert, by sheet and axis. */
function insertedLines(operations: readonly WorkbookOperation[]): InsertedLines {
  const inserted: InsertedLines = new Map();
  for (const operation of operations) {
    if (operation.kind !== 'lines' || operation.extentOnly || operation.action !== 'insert')
      continue;
    let axes = inserted.get(operation.sheetId);
    if (!axes) {
      axes = { row: new Set(), column: new Set() };
      inserted.set(operation.sheetId, axes);
    }
    for (const id of operation.ids) axes[operation.axis].add(id);
  }
  return inserted;
}

/**
 * Removing inserted lines removes their new cells too; do not write to those identities after
 * the inverse line command has removed them.
 */
function invertCells(
  operation: CellsOperation,
  inserted: InsertedLines,
): WorkbookOperation | undefined {
  const axes = inserted.get(operation.sheetId);
  const changes = operation.changes
    .filter(
      (change) => !axes?.row.has(change.rowId ?? '') && !axes?.column.has(change.columnId ?? ''),
    )
    .map(({ inverseBaselineValue, ...change }) => ({
      ...change,
      address: change.beforeAddress ?? change.address,
      beforeAddress: change.address,
      before: {
        ...structuredClone(change.after),
        value: inverseBaselineValue ?? change.after.value,
      },
      after: structuredClone(change.before),
    }));
  const shape = operation.shape && invertCellShape(operation.shape);
  return changes.length || shape ? { ...operation, changes, ...(shape && { shape }) } : undefined;
}

/** The inverse line command, then the cells, geometry and merges it restores. */
function invertLines(operation: LinesOperation): WorkbookOperation[] {
  const { sheetId } = operation;
  const inverse: WorkbookOperation[] = [
    {
      ...operation,
      preparation: undefined,
      restoreLayout: true,
      dataRows: operation.dataRows && {
        before: operation.dataRows.after,
        after: operation.dataRows.before,
      },
      action: operation.action === 'insert' ? 'delete' : 'insert',
      beforeGeometry: operation.afterGeometry,
      afterGeometry: operation.beforeGeometry,
      beforeMerges: operation.afterMerges,
      afterMerges: operation.beforeMerges,
    },
  ];
  if (operation.action === 'delete' && operation.deletedCells?.length)
    inverse.push({
      kind: 'cells',
      sheetId,
      changes: operation.deletedCells.map((change) => ({
        ...change,
        before: { value: null },
        after: structuredClone(change.before),
      })),
    });
  if (operation.beforeGeometry && operation.afterGeometry)
    inverse.push({
      kind: 'geometry',
      sheetId,
      before: structuredClone(operation.afterGeometry),
      after: structuredClone(operation.beforeGeometry),
    });
  if (operation.beforeMerges || operation.afterMerges)
    inverse.push({
      kind: 'merges',
      sheetId,
      before: structuredClone(operation.afterMerges),
      after: structuredClone(operation.beforeMerges),
    });
  return inverse;
}

function invertOther(operation: WorkbookOperation): WorkbookOperation {
  if (operation.kind === 'sheet-add' || operation.kind === 'sheet-delete')
    return {
      ...operation,
      kind: operation.kind === 'sheet-add' ? 'sheet-delete' : 'sheet-add',
      beforeOrder: [...operation.afterOrder],
      afterOrder: [...operation.beforeOrder],
    };
  if (operation.kind === 'annotations') {
    const before = structuredClone(operation.after);
    if (operation.inverseReferenceMove)
      moveReviewReferences(before, operation.inverseReferenceMove);
    return { kind: 'annotations', before, after: structuredClone(operation.before) };
  }
  const swapped = operation as WorkbookOperation & { before: unknown; after: unknown };
  return {
    ...swapped,
    before: structuredClone(swapped.after),
    after: structuredClone(swapped.before),
  } as WorkbookOperation;
}

/** Invert recorded command semantics; structure remains explicit even for blank lines. */
export function invertWorkbookOperations(
  operations: readonly WorkbookOperation[],
): WorkbookOperation[] {
  const inserted = insertedLines(operations);
  let inverse: WorkbookOperation[] = [];
  for (const operation of [...operations].reverse()) {
    if (operation.kind === 'cells') {
      const cells = invertCells(operation, inserted);
      if (cells) inverse.push(cells);
    } else if (operation.kind === 'lines' && !operation.extentOnly) {
      // The inverse line command runs before every inverse recorded so far.
      inverse = [...invertLines(operation), ...inverse];
      if (operation.preparation)
        inverse.push({
          kind: 'line-frame',
          sheetId: operation.sheetId,
          before: operation.preparation.after,
          after: operation.preparation.before,
        });
    } else if (operation.kind === 'lines') {
      inverse.push({ ...operation, action: operation.action === 'insert' ? 'delete' : 'insert' });
    } else inverse.push(invertOther(operation));
  }
  return inverse;
}
