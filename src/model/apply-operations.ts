import { authoredAddresses } from './cell-operations.js';
import { createWorkbookEntityId, type WorkbookIdFactory, worksheetExtent } from './extent.js';
import { applyWorkbookOperations } from './operation-apply.js';
import { equal, sheetOf, WorkbookOperationConflict } from './operation-delta.js';
import {
  type AuthoredCellState,
  authoredCellState,
  type WorkbookChangeContext,
  type WorkbookOperation,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import { assertWorkbook, copyWorkbook, type Workbook, type Worksheet } from './workbook.js';

export { applyWorkbookOperations } from './operation-apply.js';
export { WorkbookOperationConflict } from './operation-delta.js';

type Axis = 'row' | 'column';
type Transaction = WorkbookChangeContext;

/** One line's authored cell at an address, keyed by its row and column identities. */
function lineCell(
  sheet: Worksheet,
  address: string,
  { axis, index }: { axis: Axis; index: number },
): [string, AuthoredCellState] | undefined {
  const point = requireDefined(parseSpreadsheetAddress(address));
  if ((axis === 'row' ? point.y : point.x) !== index) return undefined;
  const state = authoredCellState(sheet, address);
  if ((state.value === null || state.value === '') && Object.keys(state).length === 1)
    return undefined;
  const row = sheet.rowIds?.[point.y],
    column = sheet.columnIds?.[point.x];
  if (!row || !column) throw new WorkbookOperationConflict('line-content-identities');
  return [JSON.stringify([row, column]), state];
}

/** Compare a line's authored cells by identity, including sparse metadata and geometry. */
function lineContent(sheet: Worksheet, axis: Axis, id: string): unknown {
  const index = (axis === 'row' ? sheet.rowIds : sheet.columnIds)?.indexOf(id) ?? -1;
  if (index < 0) return undefined;
  const cells: Record<string, AuthoredCellState> = {};
  for (const address of authoredAddresses(sheet)) {
    const cell = lineCell(sheet, address, { axis, index });
    if (cell) cells[cell[0]] = cell[1];
  }
  return { cells, geometry: axis === 'row' ? sheet.rows?.[String(index)] : sheet.columns?.[index] };
}

const formulaInput = (value: unknown) => typeof value === 'string' && value.startsWith('=');

/** A deleted line someone changed since the baseline is not deleted out from under them. */
function checkDeletedLines(current: Workbook, transaction: Transaction) {
  for (const operation of transaction.operations) {
    if (operation.kind !== 'lines' || operation.action !== 'delete') continue;
    const previous = sheetOf(transaction.baseline, operation.sheetId);
    const latest = sheetOf(current, operation.sheetId);
    for (const id of operation.ids) {
      const target = lineContent(latest, operation.axis, id);
      if (target !== undefined && !equal(lineContent(previous, operation.axis, id), target))
        throw new WorkbookOperationConflict(`modified-line/${id}`);
    }
  }
}

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
  if (operation.kind === 'geometry' || operation.kind === 'merges' || operation.kind === 'replace')
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
    applyTransaction(_next: Workbook, transaction: WorkbookChangeContext): void {
      if (host.hasTransaction?.(transaction.transactionId))
        throw new WorkbookOperationConflict('transaction-id');
      const next = applyWorkbookTransaction(host.getSnapshot(), transaction, {
        appliedTransactionIds: accepted,
      });
      host.commit(next, transaction);
      accepted.add(transaction.transactionId);
      if (accepted.size > capacity) accepted.delete(requireDefined(accepted.values().next().value));
    },
  };
}
