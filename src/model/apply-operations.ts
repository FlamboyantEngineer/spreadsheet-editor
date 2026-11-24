import { createWorkbookEntityId, type WorkbookIdFactory, worksheetExtent } from './extent.js';
import { checkDeletedLines } from './line-conflicts.js';
import { applyWorkbookOperations } from './operation-apply.js';
import { equal, WorkbookOperationConflict } from './operation-delta.js';
import type { WorkbookChangeContext, WorkbookOperation } from './operations.js';
import { requireDefined } from './require-defined.js';
import { assertWorkbook, copyWorkbook, type Workbook } from './workbook.js';

export { applyWorkbookOperations } from './operation-apply.js';
export { WorkbookOperationConflict } from './operation-delta.js';

type Transaction = WorkbookChangeContext;
const formulaInput = (value: unknown) =>
  typeof value === 'string' && value.trimStart().startsWith('=');

/** Formulas and suggestions name sheets, so they need the baseline's sheet names. */
function checkReferences(current: Workbook, transaction: Transaction) {
  const changedReferences = transaction.baseline.worksheets.some(
    (sheet) =>
      current.worksheets.find((candidate) => candidate.id === sheet.id)?.name !== sheet.name,
  );
  const referential = (operation: WorkbookOperation) =>
    (operation.kind === 'cells' &&
      operation.changes.some((change) => formulaInput(change.after.value))) ||
    (operation.kind === 'annotations' &&
      !equal(operation.before.suggestions, operation.after.suggestions));
  if (changedReferences && transaction.operations.some(referential))
    throw new WorkbookOperationConflict('referential-input-baseline');
}

/** Anchored comments and suggestions cannot follow lines that moved since the baseline. */
function checkStructuralAnnotations(
  operation: Extract<WorkbookOperation, { kind: 'annotations' }>,
) {
  for (const comment of operation.after.comments ?? []) {
    const before = operation.before.comments?.find((candidate) => candidate.id === comment.id);
    const moved = !before || before.range !== comment.range || before.sheetId !== comment.sheetId;
    if (comment.range && moved) throw new WorkbookOperationConflict('structural-comment-baseline');
  }
  if (!equal(operation.before.suggestions, operation.after.suggestions))
    throw new WorkbookOperationConflict('structural-suggestion-baseline');
}

/** Positional operations are rejected after a concurrent structure change. */
function checkStructuralOperation(operation: WorkbookOperation) {
  if (['geometry', 'merges', 'replace', 'line-frame'].includes(operation.kind))
    throw new WorkbookOperationConflict('structural-baseline');
  if (
    operation.kind === 'cells' &&
    operation.changes.some(
      (change) => !change.rowId || !change.columnId || formulaInput(change.after.value),
    )
  )
    throw new WorkbookOperationConflict('structural-cell-baseline');
  if (operation.kind === 'annotations') checkStructuralAnnotations(operation);
}

function checkStructure(current: Workbook, transaction: Transaction) {
  const changedStructure = transaction.baseline.worksheets.some((sheet) => {
    const target = current.worksheets.find((candidate) => candidate.id === sheet.id);
    return (
      !target || !equal(sheet.rowIds, target.rowIds) || !equal(sheet.columnIds, target.columnIds)
    );
  });
  if (changedStructure) transaction.operations.forEach(checkStructuralOperation);
}

/**
 * Transaction baseline supplies structural causality. Positional layout/anchor changes
 * are rejected after concurrent structure changes, rather than applied to another entity.
 */
export function applyWorkbookTransaction(
  current: Workbook,
  transaction: WorkbookChangeContext,
  options: { appliedTransactionIds?: ReadonlySet<string> } = {},
): Workbook {
  if (!transaction.transactionId || options.appliedTransactionIds?.has(transaction.transactionId))
    throw new WorkbookOperationConflict('transaction-id');
  assertWorkbook(transaction.baseline);
  if (
    transaction.operations.some((operation) => operation.kind === 'sheet-rename') &&
    !equal(current, transaction.baseline)
  )
    throw new WorkbookOperationConflict('referential-rename-baseline');
  checkDeletedLines(current, transaction);
  checkReferences(current, transaction);
  checkStructure(current, transaction);
  return applyWorkbookOperations(current, transaction.operations);
}

/** Persist this initialized snapshot once before opening concurrent editing sessions. */
export function initializeWorkbookIdentities(
  workbook: Workbook,
  idFactory: WorkbookIdFactory = createWorkbookEntityId,
): Workbook {
  assertWorkbook(workbook);
  const next = copyWorkbook(workbook);
  for (const sheet of next.worksheets) {
    const extent = worksheetExtent(sheet, next);
    sheet.rowIds ??= [];
    sheet.columnIds ??= [];
    while (sheet.rowIds.length < extent.rows) sheet.rowIds.push(idFactory('row'));
    while (sheet.columnIds.length < extent.columns) sheet.columnIds.push(idFactory('column'));
  }
  assertWorkbook(next);
  return next;
}

/** Bounded session deduplication; host durable identity lookup can cover older/restarted work. */
export function createWorkbookOperationBinding(
  host: {
    getSnapshot(): Workbook;
    hasTransaction?(id: string): boolean;
    /** Authorize and persist synchronously in one serialized lane; throw only before mutation. */
    commit(workbook: Workbook, transaction: WorkbookChangeContext): void;
  },
  options: { rememberTransactions?: number } = {},
) {
  const capacity = options.rememberTransactions ?? 10_000;
  if (!Number.isSafeInteger(capacity) || capacity < 1)
    throw new RangeError('Transaction memory capacity must be a positive integer.');
  const accepted = new Set<string>();
  return {
    applyTransaction(_next: Workbook, transaction: WorkbookChangeContext): Workbook {
      if (host.hasTransaction?.(transaction.transactionId))
        throw new WorkbookOperationConflict('transaction-id');
      const next = applyWorkbookTransaction(host.getSnapshot(), transaction, {
        appliedTransactionIds: accepted,
      });
      host.commit(next, transaction);
      accepted.add(transaction.transactionId);
      if (accepted.size > capacity) accepted.delete(requireDefined(accepted.values().next().value));
      return next;
    },
  };
}
