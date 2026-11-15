import { describe, expect, it } from 'vitest';
import { createWorkbook, copyWorkbook } from './workbook.js';
import { materializeWorksheetExtent } from './extent.js';
import {
  applyWorkbookOperations,
  WorkbookOperationConflict,
  applyWorkbookTransaction,
  createWorkbookOperationBinding,
  initializeWorkbookIdentities,
} from './apply-operations.js';
import {
  annotationOperation,
  changeWorkbookStructureWithOperations,
  geometryOperation,
  invertWorkbookOperations,
  sheetOperation,
  writeCellsWithOperations,
} from './operations.js';
import { renameWorksheet } from './structure.js';
import { commandOperations } from '../editor/command-operations.js';
import { addComment, replyToComment, updateComment } from './review.js';

const base = () => {
  const workbook = materializeWorksheetExtent(
    createWorkbook([
      ['1', '2'],
      ['3', '4'],
    ]),
    'sheet-1',
    [2, 2],
  );
  workbook.worksheets[0].rowIds = ['row-1', 'row-2'];
  workbook.worksheets[0].columnIds = ['column-1', 'column-2'];
  return workbook;
};
const stamp = (id: string) => ({ id, time: '2026-01-01T00:00:00Z', actor: { id, name: id } });
describe('SE-HOST-1: atomic serialized operation application', () => {
  it('preserves concurrent cells and per-actor inverse without changing canonical numeric inputs', () => {
    const initial = base();
    initial.worksheets[0].inputValues = {
      A1: { kind: 'number', numerator: '1', denominator: '3', approximate: false },
    };
    const first = writeCellsWithOperations(initial, 'sheet-1', 0, 0, [['first']]);
    const second = writeCellsWithOperations(initial, 'sheet-1', 1, 0, [['second']]);
    const shared = applyWorkbookOperations(
      applyWorkbookOperations(initial, first.operations),
      second.operations,
    );
    expect(shared.worksheets[0].data[0]).toEqual(['first', 'second']);
    const undone = applyWorkbookOperations(shared, invertWorkbookOperations(first.operations));
    expect(undone.worksheets[0].data[0]).toEqual(['1', 'second']);
    expect(undone.worksheets[0].inputValues?.A1?.denominator).toBe('3');
    expect(initial.worksheets[0].data[0]).toEqual(['1', '2']);
  });
  it('rebases same-row inserts, deletes only targeted identities and follows surviving cells', () => {
    const initial = base();
    const a = changeWorkbookStructureWithOperations(initial, 'sheet-1', {
      axis: 'row',
      kind: 'insert',
      index: 1,
      count: 1,
      ids: ['a'],
    });
    const b = changeWorkbookStructureWithOperations(initial, 'sheet-1', {
      axis: 'row',
      kind: 'insert',
      index: 1,
      count: 1,
      ids: ['b'],
    });
    const cell = writeCellsWithOperations(initial, 'sheet-1', 1, 1, [['survives']]);
    let shared = applyWorkbookOperations(
      applyWorkbookOperations(initial, a.operations),
      b.operations,
    );
    shared = applyWorkbookOperations(shared, cell.operations);
    expect(shared.worksheets[0].rowIds?.slice(0, 4)).toEqual([
      initial.worksheets[0].rowIds![0],
      'a',
      'b',
      initial.worksheets[0].rowIds![1],
    ]);
    expect(shared.worksheets[0].data[3][1]).toBe('survives');
    const deletion = changeWorkbookStructureWithOperations(initial, 'sheet-1', {
      axis: 'row',
      kind: 'delete',
      index: 1,
      count: 1,
    });
    expect(() =>
      applyWorkbookTransaction(shared, {
        transactionId: 'stale-delete',
        source: 'user',
        baseline: initial,
        operations: deletion.operations,
      }),
    ).toThrow('modified-line');
    shared = applyWorkbookOperations(shared, deletion.operations);
    expect(shared.worksheets[0].rowIds).toContain('a');
    expect(shared.worksheets[0].rowIds).toContain('b');
    expect(() => applyWorkbookOperations(shared, cell.operations)).toThrow(
      WorkbookOperationConflict,
    );
  });
  it('merges simultaneous same-thread replies and reaction membership; inverse preserves the other actor', () => {
    const initial = addComment(base(), 'sheet-1', 'A1', 'start', stamp('thread'));
    const a = replyToComment(initial, 'thread', 'a', stamp('a'));
    const b = replyToComment(initial, 'thread', 'b', stamp('b'));
    const reactionA = updateComment(a, 'thread', (thread) => {
      thread.entries[0].reactions = { '+1': ['a'] };
    });
    const reactionB = updateComment(b, 'thread', (thread) => {
      thread.entries[0].reactions = { '+1': ['b'] };
    });
    const opA = annotationOperation(initial, reactionA),
      opB = annotationOperation(initial, reactionB);
    const shared = applyWorkbookOperations(applyWorkbookOperations(initial, [opA]), [opB]);
    expect(shared.comments![0].entries.map((entry) => entry.id)).toEqual(['thread', 'a', 'b']);
    expect(shared.comments![0].entries[0].reactions).toEqual({ '+1': ['a', 'b'] });
    const undone = applyWorkbookOperations(shared, invertWorkbookOperations([opA]));
    expect(undone.comments![0].entries.map((entry) => entry.id)).toEqual(['thread', 'b']);
    expect(undone.comments![0].entries[0].reactions).toEqual({ '+1': ['b'] });
  });
  it('merges independent geometry fields and rejects competing fields/order without partial application', () => {
    const initial = base();
    initial.worksheets[0].columns = [{ width: 80 }, { width: 90 }];
    const a = copyWorkbook(initial),
      b = copyWorkbook(initial);
    a.worksheets[0].columns![0].width = 100;
    b.worksheets[0].columns![1].width = 120;
    const shared = applyWorkbookOperations(
      applyWorkbookOperations(initial, [geometryOperation(initial.worksheets[0], a.worksheets[0])]),
      [geometryOperation(initial.worksheets[0], b.worksheets[0])],
    );
    expect(shared.worksheets[0].columns).toEqual([{ width: 100 }, { width: 120 }]);
    const stale = copyWorkbook(initial);
    stale.worksheets[0].columns![0].width = 150;
    const cell = writeCellsWithOperations(initial, 'sheet-1', 0, 0, [['must rollback']]);
    expect(() =>
      applyWorkbookOperations(shared, [
        ...cell.operations,
        geometryOperation(initial.worksheets[0], stale.worksheets[0]),
      ]),
    ).toThrow(WorkbookOperationConflict);
    expect(shared.worksheets[0].data[0][0]).toBe('1');
    const sheets = base();
    sheets.worksheets.push(
      { id: 'two', name: 'Two', data: [['']] },
      { id: 'three', name: 'Three', data: [['']] },
    );
    const reordered = copyWorkbook(sheets);
    reordered.worksheets.reverse();
    const op = sheetOperation('sheet-reorder', sheets, reordered);
    const latest = copyWorkbook(sheets);
    [latest.worksheets[0], latest.worksheets[1]] = [latest.worksheets[1], latest.worksheets[0]];
    expect(() => applyWorkbookOperations(latest, [op])).toThrow(WorkbookOperationConflict);
  });
  it('rejects conflicting suggestion decisions atomically while retaining unrelated decisions', () => {
    const initial = base();
    initial.suggestions = ['a', 'b'].map((id) => ({
      id,
      author: stamp(id).actor,
      createdAt: stamp(id).time,
      status: 'pending',
      changes: [],
      summary: id,
    }));
    const a = copyWorkbook(initial),
      b = copyWorkbook(initial),
      competing = copyWorkbook(initial);
    a.suggestions![0].status = 'accepted';
    b.suggestions![1].status = 'rejected';
    competing.suggestions![0].status = 'rejected';
    const shared = applyWorkbookOperations(
      applyWorkbookOperations(initial, [annotationOperation(initial, a)]),
      [annotationOperation(initial, b)],
    );
    expect(shared.suggestions!.map((suggestion) => suggestion.status)).toEqual([
      'accepted',
      'rejected',
    ]);
    expect(() =>
      applyWorkbookOperations(shared, [annotationOperation(initial, competing)]),
    ).toThrow(WorkbookOperationConflict);
  });
  it('rejects stale positional layout and anchors after structure; retries identified cells safely', () => {
    const initial = base();
    const inserted = changeWorkbookStructureWithOperations(initial, 'sheet-1', {
      axis: 'column',
      kind: 'insert',
      index: 0,
      count: 1,
      ids: ['new-column'],
    });
    const shared = applyWorkbookOperations(initial, inserted.operations);
    const geometry = copyWorkbook(initial);
    geometry.worksheets[0].columns = [{ width: 180 }];
    const safe = writeCellsWithOperations(initial, 'sheet-1', 1, 1, [['safe']]);
    const context = {
      transactionId: 'layout',
      source: 'user' as const,
      baseline: initial,
      operations: [
        ...safe.operations,
        geometryOperation(initial.worksheets[0], geometry.worksheets[0]),
      ],
    };
    expect(() => applyWorkbookTransaction(shared, context)).toThrow('structural-baseline');
    expect(shared.worksheets[0].data[1][2]).toBe('4');
    expect(
      applyWorkbookTransaction(shared, { ...context, operations: safe.operations }).worksheets[0]
        .data[1][2],
    ).toBe('safe');
    const annotated = addComment(initial, 'sheet-1', 'A1', 'position', stamp('new-thread'));
    expect(() =>
      applyWorkbookTransaction(shared, {
        ...context,
        operations: [annotationOperation(initial, annotated)],
      }),
    ).toThrow('structural-comment-baseline');
  });
  it('keeps bounded ID receipts across provider replacement and rejects duplicates before commit', () => {
    let shared = initializeWorkbookIdentities(createWorkbook([['original']]));
    const initial = copyWorkbook(shared);
    const durable = new Set<string>();
    let commits = 0;
    const binding = createWorkbookOperationBinding(
      {
        getSnapshot: () => shared,
        hasTransaction: (id) => durable.has(id),
        commit: (next, transaction) => {
          shared = next;
          durable.add(transaction.transactionId);
          commits++;
        },
      },
      { rememberTransactions: 1 },
    );
    const edit = writeCellsWithOperations(initial, 'sheet-1', 0, 0, [['accepted']]);
    const transaction = {
      transactionId: 'one',
      source: 'user' as const,
      baseline: initial,
      operations: edit.operations,
    };
    binding.applyTransaction(edit.workbook, transaction);
    expect(() => binding.applyTransaction(edit.workbook, transaction)).toThrow('transaction-id');
    // Provider reconnect/replacement leaves this host-owned applier intact.
    expect(commits).toBe(1);
    const changed = writeCellsWithOperations(shared, 'sheet-1', 0, 0, [['newer']]);
    binding.applyTransaction(changed.workbook, {
      ...transaction,
      baseline: shared,
      transactionId: 'two',
      operations: changed.operations,
    });
    expect(() => binding.applyTransaction(edit.workbook, transaction)).toThrow('transaction-id');
    expect(commits).toBe(2);
    expect(shared.worksheets[0].data[0][0]).toBe('newer');
  });
  it('restores deleted-line geometry and merges through semantic inverses', () => {
    const initial = base();
    initial.worksheets[0].rows = { '0': { height: 31 }, '1': { height: 47 } };
    initial.worksheets[0].columns = [{ width: 91 }, { width: 137 }];
    initial.worksheets[0].mergeCells = { A1: [2, 2] };
    for (const axis of ['row', 'column'] as const) {
      const removed = changeWorkbookStructureWithOperations(initial, 'sheet-1', {
        axis,
        kind: 'delete',
        index: 0,
        count: 1,
      });
      const restored = applyWorkbookOperations(
        removed.workbook,
        invertWorkbookOperations(removed.operations),
      );
      expect(restored.worksheets[0].rows).toEqual(initial.worksheets[0].rows);
      expect(restored.worksheets[0].columns).toEqual(initial.worksheets[0].columns);
      expect(restored.worksheets[0].mergeCells).toEqual(initial.worksheets[0].mergeCells);
      expect(restored.worksheets[0].data).toEqual(initial.worksheets[0].data);
    }
  });
  it('rejects a stale rename whose baseline lacks a concurrently introduced reference', () => {
    const initial = base();
    initial.worksheets.push({ id: 'other', name: 'Other', data: [['']] });
    const renamed = renameWorksheet(initial, 'sheet-1', 'Renamed');
    const latest = copyWorkbook(initial);
    latest.worksheets[1].data[0][0] = "='Sheet1'!A1";
    expect(() =>
      applyWorkbookTransaction(latest, {
        transactionId: 'rename',
        source: 'user',
        baseline: initial,
        operations: commandOperations(initial, renamed),
      }),
    ).toThrow('referential-rename-baseline');
    expect(latest.worksheets[0].name).toBe('Sheet1');
    const applied = applyWorkbookTransaction(initial, {
      transactionId: 'rename',
      source: 'user',
      baseline: initial,
      operations: commandOperations(initial, renamed),
    });
    expect(applied.worksheets[0].name).toBe('Renamed');
    const staleFormula = writeCellsWithOperations(initial, 'other', 0, 0, [["='Sheet1'!A1"]]);
    expect(() =>
      applyWorkbookTransaction(applied, {
        transactionId: 'stale-formula',
        source: 'user',
        baseline: initial,
        operations: staleFormula.operations,
      }),
    ).toThrow('referential-input-baseline');
  });
});
