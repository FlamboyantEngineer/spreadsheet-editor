import { orderedJson } from '../calculation/publication.js';
import { assertWorkbook, copyWorkbook, type Workbook, type Worksheet } from './workbook.js';
import {
  authoredCellState,
  type AuthoredCellState,
  type WorkbookChangeContext,
  type WorksheetGeometry,
  type WorkbookOperation,
} from './operations.js';
import { worksheetExtent, createWorkbookEntityId, type WorkbookIdFactory } from './extent.js';
import { changeWorkbookStructure } from './structure.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';

export class WorkbookOperationConflict extends Error {
  constructor(public readonly path: string) {
    super(`The workbook changed concurrently at ${path}.`);
    this.name = 'WorkbookOperationConflict';
  }
}
const equal = (a: unknown, b: unknown) => orderedJson(a) === orderedJson(b);
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const entities = (value: unknown[]): value is (Record<string, unknown> & { id: string })[] =>
  value.every((item) => object(item) && typeof item.id === 'string');

/** Three-way authored deltas: arrays of entities merge by ID; primitive sets by membership. */
function delta(current: unknown, before: unknown, after: unknown, path: string): unknown {
  if (equal(before, after)) return structuredClone(current);
  if (equal(current, before) || equal(current, after)) return structuredClone(after);
  if (before === undefined && Array.isArray(after) && Array.isArray(current)) before = [];
  if (after === undefined && Array.isArray(before) && Array.isArray(current)) after = [];
  if (object(before) && ('numerator' in before || 'denominator' in before))
    throw new WorkbookOperationConflict(path);
  if (Array.isArray(before) && Array.isArray(after) && Array.isArray(current)) {
    if (entities(before) && entities(after) && entities(current)) {
      const result = structuredClone(current);
      for (const previous of before) {
        const next = after.find((item) => item.id === previous.id);
        const index = result.findIndex((item) => item.id === previous.id);
        if (!next) {
          if (index < 0) continue;
          if (!equal(result[index], previous))
            throw new WorkbookOperationConflict(`${path}/${previous.id}`);
          result.splice(index, 1);
        } else if (!equal(previous, next)) {
          if (index < 0) throw new WorkbookOperationConflict(`${path}/${previous.id}`);
          result[index] = delta(
            result[index],
            previous,
            next,
            `${path}/${previous.id}`,
          ) as typeof next;
        }
      }
      for (const next of after) {
        if (before.some((item) => item.id === next.id)) continue;
        const existing = result.find((item) => item.id === next.id);
        if (existing && !equal(existing, next))
          throw new WorkbookOperationConflict(`${path}/${next.id}`);
        if (!existing) result.push(structuredClone(next));
      }
      return result;
    }
    const previous = before,
      updated = after;
    if ([...previous, ...updated, ...current].every((item) => typeof item === 'string'))
      return [
        ...new Set([
          ...current.filter((item) => !previous.includes(item) || updated.includes(item)),
          ...updated.filter((item) => !previous.includes(item)),
        ]),
      ].sort();
  }
  if (
    (object(before) || before === undefined) &&
    (object(after) || after === undefined) &&
    object(current)
  ) {
    const removed = after === undefined;
    const updated = after ?? {};
    const result = structuredClone(current);
    for (const key of new Set([...Object.keys(before ?? {}), ...Object.keys(updated)])) {
      const value = delta(current[key], before?.[key], updated[key], `${path}/${key}`);
      if (value === undefined) delete result[key];
      else result[key] = value;
    }
    return removed && !Object.keys(result).length ? undefined : result;
  }
  // Sparse geometry arrays retain unrelated column properties at their stable current indexes.
  if (
    Array.isArray(before) &&
    Array.isArray(after) &&
    Array.isArray(current) &&
    before.length === after.length &&
    current.length === before.length
  )
    return current.map((value, index) =>
      delta(value, before[index], after[index], `${path}/${index}`),
    );
  throw new WorkbookOperationConflict(path);
}
function sheetOf(workbook: Workbook, id: string): Worksheet {
  const sheet = workbook.worksheets.find((sheet) => sheet.id === id);
  if (!sheet) throw new WorkbookOperationConflict(`worksheets/${id}`);
  return sheet;
}
function writeState(sheet: Worksheet, address: string, state: AuthoredCellState): void {
  const cell = parseSpreadsheetAddress(address)!;
  while (sheet.data.length <= cell.y) sheet.data.push([]);
  while (sheet.data[cell.y].length <= cell.x) sheet.data[cell.y].push(null);
  sheet.data[cell.y][cell.x] = state.value;
  for (const [field, key] of [
    ['style', 'style'],
    ['formats', 'format'],
    ['inputValues', 'inputValue'],
    ['cellProvenance', 'provenance'],
  ] as const) {
    const value = state[key];
    if (value === undefined) delete sheet[field]?.[address];
    else
      Object.assign((sheet[field] ??= {}) as Record<string, unknown>, {
        [address]: structuredClone(value),
      });
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
  let next = copyWorkbook(current);
  const lineLayouts = new Map<
    string,
    {
      beforeGeometry?: WorksheetGeometry;
      geometry: WorksheetGeometry;
      beforeMerges?: Worksheet['mergeCells'];
      merges: Worksheet['mergeCells'];
    }
  >();
  for (const operation of operations) {
    switch (operation.kind) {
      case 'cells': {
        const sheet = sheetOf(next, operation.sheetId);
        for (const change of operation.changes) {
          const position = parseSpreadsheetAddress(change.address)!;
          const y = change.rowId ? (sheet.rowIds?.indexOf(change.rowId) ?? -1) : position.y;
          const x = change.columnId
            ? (sheet.columnIds?.indexOf(change.columnId) ?? -1)
            : position.x;
          if (x < 0 || y < 0)
            throw new WorkbookOperationConflict(`cells/${operation.sheetId}/${change.address}`);
          const address = formatSpreadsheetAddress(x, y);
          writeState(
            sheet,
            address,
            delta(
              authoredCellState(sheet, address),
              change.before,
              change.after,
              `cells/${sheet.id}/${address}`,
            ) as AuthoredCellState,
          );
        }
        break;
      }
      case 'lines': {
        const sheet = sheetOf(next, operation.sheetId);
        const ids = operation.axis === 'row' ? sheet.rowIds : sheet.columnIds;
        if (operation.beforeGeometry) {
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
          lineLayouts.set(sheet.id, {
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
        if (!ids) throw new WorkbookOperationConflict(`lines/${sheet.id}`);
        if (operation.action === 'insert') {
          if (operation.ids.some((id) => ids.includes(id)))
            throw new WorkbookOperationConflict(`lines/${sheet.id}`);
          const after = operation.afterId ? ids.indexOf(operation.afterId) : -1;
          const before = operation.beforeId ? ids.indexOf(operation.beforeId) : -1;
          const index =
            after >= 0 ? after : before >= 0 ? before + 1 : Math.min(operation.index, ids.length);
          next = changeWorkbookStructure(next, sheet.id, {
            axis: operation.axis,
            kind: 'insert',
            index,
            count: operation.ids.length,
            ids: operation.ids,
          });
        } else {
          const positions = operation.ids
            .map((id) => ids.indexOf(id))
            .filter((index) => index >= 0)
            .sort((a, b) => b - a);
          for (const index of positions)
            next = changeWorkbookStructure(next, sheet.id, {
              axis: operation.axis,
              kind: 'delete',
              index,
              count: 1,
            });
        }
        break;
      }
      case 'annotations': {
        const value = delta(
          { comments: next.comments ?? [], suggestions: next.suggestions ?? [] },
          {
            comments: operation.before.comments ?? [],
            suggestions: operation.before.suggestions ?? [],
          },
          {
            comments: operation.after.comments ?? [],
            suggestions: operation.after.suggestions ?? [],
          },
          'annotations',
        ) as Pick<Workbook, 'comments' | 'suggestions'>;
        // A list neither side had stays absent, so an inverse restores the workbook exactly.
        for (const key of ['comments', 'suggestions'] as const)
          if (next[key] !== undefined || operation.after[key] !== undefined || value[key]?.length)
            next[key] = value[key] as never;
        break;
      }
      case 'geometry':
      case 'merges': {
        const sheet = sheetOf(next, operation.sheetId);
        if (operation.kind === 'merges')
          sheet.mergeCells = delta(
            sheet.mergeCells ?? {},
            lineLayouts.has(sheet.id) &&
              equal(operation.before, lineLayouts.get(sheet.id)!.beforeMerges)
              ? (lineLayouts.get(sheet.id)!.merges ?? {})
              : operation.before,
            operation.after,
            `merges/${sheet.id}`,
          ) as Worksheet['mergeCells'];
        else {
          const geometry = {
            minDimensions: sheet.minDimensions,
            columns: sheet.columns,
            rows: sheet.rows,
            freezeColumns: sheet.freezeColumns,
          };
          Object.assign(
            sheet,
            delta(
              geometry,
              lineLayouts.has(sheet.id) &&
                equal(operation.before, lineLayouts.get(sheet.id)!.beforeGeometry)
                ? lineLayouts.get(sheet.id)!.geometry
                : operation.before,
              operation.after,
              `geometry/${sheet.id}`,
            ),
          );
        }
        break;
      }
      case 'workbook-metadata':
        Object.assign(
          next,
          delta({ name: next.name }, operation.before, operation.after, 'workbook'),
        );
        break;
      case 'sheet-rename': {
        const sheet = sheetOf(next, operation.sheetId);
        sheet.name = delta(
          sheet.name,
          operation.before,
          operation.after,
          `worksheets/${sheet.id}/name`,
        ) as string;
        break;
      }
      case 'sheet-add': {
        if (next.worksheets.some((sheet) => sheet.id === operation.sheet.id))
          throw new WorkbookOperationConflict(`worksheets/${operation.sheet.id}`);
        const following = operation.afterOrder
          .slice(operation.afterOrder.indexOf(operation.sheet.id) + 1)
          .find((id) => next.worksheets.some((sheet) => sheet.id === id));
        const index = following
          ? next.worksheets.findIndex((sheet) => sheet.id === following)
          : next.worksheets.length;
        next.worksheets.splice(index, 0, structuredClone(operation.sheet));
        break;
      }
      case 'sheet-delete': {
        const sheet = sheetOf(next, operation.sheet.id);
        if (!equal(sheet, operation.sheet))
          throw new WorkbookOperationConflict(`worksheets/${sheet.id}`);
        next.worksheets = next.worksheets.filter((item) => item.id !== sheet.id);
        break;
      }
      case 'sheet-reorder': {
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
          known.includes(sheet.id) ? reordered.shift()! : sheet,
        );
        break;
      }
      case 'replace':
        if (!equal(next, operation.before)) throw new WorkbookOperationConflict('workbook');
        next = copyWorkbook(operation.after);
        break;
    }
  }
  assertWorkbook(next);
  return next;
}
/** Compare a line's authored cells by identity, including sparse metadata and geometry. */
function lineContent(sheet: Worksheet, axis: 'row' | 'column', id: string): unknown {
  const index = (axis === 'row' ? sheet.rowIds : sheet.columnIds)?.indexOf(id) ?? -1;
  if (index < 0) return undefined;
  const addresses = new Set<string>();
  sheet.data.forEach((row, y) =>
    row.forEach((_value, x) => addresses.add(formatSpreadsheetAddress(x, y))),
  );
  for (const metadata of [sheet.style, sheet.formats, sheet.inputValues, sheet.cellProvenance])
    for (const address of Object.keys(metadata ?? {})) addresses.add(address);
  const cells: Record<string, AuthoredCellState> = {};
  for (const address of addresses) {
    const point = parseSpreadsheetAddress(address)!;
    if ((axis === 'row' ? point.y : point.x) !== index) continue;
    const state = authoredCellState(sheet, address);
    if ((state.value === null || state.value === '') && Object.keys(state).length === 1) continue;
    const row = sheet.rowIds?.[point.y],
      column = sheet.columnIds?.[point.x];
    if (!row || !column) throw new WorkbookOperationConflict('line-content-identities');
    cells[JSON.stringify([row, column])] = state;
  }
  return { cells, geometry: axis === 'row' ? sheet.rows?.[String(index)] : sheet.columns?.[index] };
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
  const changedReferences = transaction.baseline.worksheets.some(
    (sheet) =>
      current.worksheets.find((candidate) => candidate.id === sheet.id)?.name !== sheet.name,
  );
  if (
    changedReferences &&
    transaction.operations.some(
      (operation) =>
        (operation.kind === 'cells' &&
          operation.changes.some(
            (change) =>
              typeof change.after.value === 'string' && change.after.value.startsWith('='),
          )) ||
        (operation.kind === 'annotations' &&
          !equal(operation.before.suggestions, operation.after.suggestions)),
    )
  )
    throw new WorkbookOperationConflict('referential-input-baseline');
  const changedStructure = transaction.baseline.worksheets.some((sheet) => {
    const target = current.worksheets.find((candidate) => candidate.id === sheet.id);
    return (
      !target || !equal(sheet.rowIds, target.rowIds) || !equal(sheet.columnIds, target.columnIds)
    );
  });
  if (changedStructure) {
    for (const operation of transaction.operations) {
      if (
        operation.kind === 'geometry' ||
        operation.kind === 'merges' ||
        operation.kind === 'replace'
      )
        throw new WorkbookOperationConflict('structural-baseline');
      if (
        operation.kind === 'cells' &&
        operation.changes.some(
          (change) =>
            !change.rowId ||
            !change.columnId ||
            (typeof change.after.value === 'string' && change.after.value.startsWith('=')),
        )
      )
        throw new WorkbookOperationConflict('structural-cell-baseline');
      if (operation.kind === 'annotations') {
        for (const comment of operation.after.comments ?? []) {
          const before = operation.before.comments?.find(
            (candidate) => candidate.id === comment.id,
          );
          if (
            comment.range &&
            (!before || before.range !== comment.range || before.sheetId !== comment.sheetId)
          )
            throw new WorkbookOperationConflict('structural-comment-baseline');
        }
        if (!equal(operation.before.suggestions, operation.after.suggestions))
          throw new WorkbookOperationConflict('structural-suggestion-baseline');
      }
    }
  }
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
      if (accepted.size > capacity) accepted.delete(accepted.values().next().value!);
    },
  };
}
