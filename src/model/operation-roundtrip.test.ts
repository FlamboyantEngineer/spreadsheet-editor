import { describe, expect, it, vi } from 'vitest';
import { commandOperations } from '../editor/command-operations.js';
import {
  applyWorkbookOperations,
  applyWorkbookTransaction,
  initializeWorkbookIdentities,
} from './apply-operations.js';
import { formatCells, mergeSelection, writeCells } from './editing.js';
import {
  changeWorkbookStructureWithOperations,
  invertWorkbookOperations,
  type WorkbookCommandResult,
  writeCellsWithOperations,
} from './operations.js';
import {
  addComment,
  decideSuggestionWithOperations,
  proposeWorkbookChange,
  replyToComment,
} from './review.js';
import { deleteWorksheet, type StructureChange } from './structure.js';
import { createWorkbook, type Workbook } from './workbook.js';

const stamp = (id: string) => ({ id, time: '2026-01-01T00:00:00Z', actor: { id, name: id } });
const idFactory = () => {
  let sequence = 0;
  return (kind: string) => `${kind}-${++sequence}`;
};
const sheetId = 'sheet-1';
const range = { x1: 1, y1: 1, x2: 2, y2: 2 };
const format = { properties: { 'font-weight': 'bold' } };

function roundtrip(before: Workbook, command: WorkbookCommandResult) {
  const replay = applyWorkbookOperations(before, command.operations);
  expect(replay).toStrictEqual(command.workbook);
  expect(
    applyWorkbookOperations(replay, invertWorkbookOperations(command.operations)),
  ).toStrictEqual(before);
}

describe('SE-HOST-MODEL-2.1: exact semantic operation round trips', () => {
  it('replays extent growth without changing layout and restores absent blanks and optional maps', () => {
    const makeId = idFactory();
    const before = initializeWorkbookIdentities(createWorkbook(), makeId);
    for (const value of ['new', null]) {
      const command = writeCellsWithOperations(before, [[value]], {
        sheetId,
        x: 2,
        y: 2,
        idFactory: makeId,
      });
      expect(command.workbook.worksheets[0].minDimensions).toBeUndefined();
      roundtrip(before, command);
    }
    const unidentified = createWorkbook();
    const styled = formatCells(unidentified, { sheetId, range }, format);
    expect(styled.worksheets[0].data).toEqual(unidentified.worksheets[0].data);
    roundtrip(unidentified, {
      workbook: styled,
      operations: commandOperations(unidentified, styled),
    });
  });

  it('assigns stable identities to formatted, merged and annotated blank cells', () => {
    const makeId = idFactory();
    const before = initializeWorkbookIdentities(createWorkbook(), makeId);
    for (const next of [
      formatCells(before, { sheetId, range, idFactory: makeId }, format),
      mergeSelection(before, { sheetId, range, idFactory: makeId }),
      addComment(
        before,
        { sheetId, range: 'B2:C3', text: 'Check this', idFactory: makeId },
        stamp('comment'),
      ),
    ]) {
      expect(next.worksheets[0].rowIds).toHaveLength(3);
      expect(next.worksheets[0].columnIds).toHaveLength(3);
      roundtrip(before, { workbook: next, operations: commandOperations(before, next) });
    }
  });

  it.each(['insert', 'delete'] as const)(
    'records preparation to %s a row in the implicit viewport',
    (kind) => {
      const makeId = idFactory();
      const before = initializeWorkbookIdentities(createWorkbook([['1', '=A1']]), makeId);
      const command = changeWorkbookStructureWithOperations(
        before,
        { sheetId, idFactory: makeId },
        { axis: 'row', kind, index: 2, count: 1 },
      );
      expect(command.workbook.worksheets[0].minDimensions).toEqual([
        26,
        kind === 'insert' ? 101 : 99,
      ]);
      roundtrip(before, command);
    },
  );

  it('restores a deleted trailing empty row without manufacturing a blank cell', () => {
    const makeId = idFactory();
    const before = initializeWorkbookIdentities(createWorkbook([['1'], [], []]), makeId);
    const command = changeWorkbookStructureWithOperations(
      before,
      { sheetId, idFactory: makeId },
      { axis: 'row', kind: 'delete', index: 2, count: 1 },
    );
    expect(command.workbook.worksheets[0].data).toEqual([['1'], []]);
    roundtrip(before, command);
  });

  it('restores the clipped endpoints of formula ranges and comment anchors', () => {
    const makeId = idFactory();
    const workbook = createWorkbook([['1'], ['2'], ['3', '=SUM(A1:A3)']]);
    workbook.worksheets[0].minDimensions = [2, 3];
    workbook.worksheets.push({
      id: 'other',
      name: 'Other',
      data: [["=SUM('Sheet1'!A1:A3)", 'old']],
    });
    const before = initializeWorkbookIdentities(
      addComment(workbook, { sheetId, range: 'A1:A3', text: 'Check totals' }, stamp('thread')),
      makeId,
    );
    const command = changeWorkbookStructureWithOperations(
      before,
      { sheetId, idFactory: makeId },
      { axis: 'row', kind: 'delete', index: 0, count: 1 },
    );
    expect(command.workbook.comments?.[0].range).toBe('A1:A2');
    expect(command.workbook.worksheets[0].data[1][1]).toBe('=SUM(A1:A2)');
    roundtrip(before, command);

    const peerReply = { id: 'thread', text: 'Keep my reply' };
    const peerEdit = (value: Workbook) =>
      replyToComment(
        writeCells(value, [['peer']], { sheetId: 'other', x: 1, y: 0 }),
        peerReply,
        stamp('reply'),
      );
    const inverse = invertWorkbookOperations(command.operations);
    expect(
      applyWorkbookTransaction(peerEdit(command.workbook), {
        baseline: command.workbook,
        source: 'undo',
        transactionId: 'undo',
        operations: inverse,
      }),
    ).toStrictEqual(peerEdit(before));
  });

  it('uses a semantic line command when accepting a legacy structural suggestion', () => {
    const makeId = idFactory();
    const before = initializeWorkbookIdentities(createWorkbook([['1'], ['=A1']]), makeId);
    const change = { axis: 'row', kind: 'insert', index: 0, count: 1 } satisfies StructureChange;
    const proposal = proposeWorkbookChange(
      before,
      {
        proposed: changeWorkbookStructureWithOperations(
          before,
          { sheetId, idFactory: makeId, materialize: false },
          change,
        ).workbook,
        operation: { sheetId, change },
        summary: 'Insert a row',
      },
      stamp('suggestion'),
    );
    const withComment = addComment(
      proposal,
      { sheetId, range: 'A2', text: 'Check this' },
      stamp('comment'),
    );
    const decided = decideSuggestionWithOperations(
      withComment,
      { id: 'suggestion', decision: 'accepted', idFactory: makeId },
      stamp('decision'),
    );
    expect(decided.workbook.comments?.[0].range).toBe('A3');
    expect(decided.operations?.[0]).toMatchObject({ kind: 'lines', axis: 'row', action: 'insert' });
    roundtrip(withComment, { workbook: decided.workbook, operations: decided.operations ?? [] });
  });
});

describe('SE-HOST-5.1: transaction conflicts and bulk structure', () => {
  const withIds = (makeId: ReturnType<typeof idFactory>) =>
    initializeWorkbookIdentities(
      {
        ...createWorkbook([
          ['1', ''],
          ['', ''],
        ]),
        worksheets: [
          {
            id: sheetId,
            name: 'Sheet1',
            minDimensions: [2, 2],
            data: [
              ['1', ''],
              ['', ''],
            ],
          },
        ],
      },
      makeId,
    );

  it('rejects a stale formula with leading whitespace and a deletion of a newly merged line', () => {
    const makeId = idFactory();
    const before = withIds(makeId);
    const inserted = changeWorkbookStructureWithOperations(
      before,
      { sheetId, idFactory: makeId },
      { axis: 'row', kind: 'insert', index: 0, count: 1 },
    ).workbook;
    const stale = writeCellsWithOperations(before, [[' =A1']], { sheetId, x: 1, y: 1 });
    expect(() =>
      applyWorkbookTransaction(inserted, {
        baseline: before,
        source: 'user',
        transactionId: 'formula',
        operations: stale.operations,
      }),
    ).toThrow('structural-cell-baseline');

    const merged = mergeSelection(before, { sheetId, range: { x1: 0, y1: 0, x2: 1, y2: 0 } });
    const deleted = changeWorkbookStructureWithOperations(
      before,
      { sheetId, idFactory: makeId },
      { axis: 'row', kind: 'delete', index: 0, count: 1 },
    );
    expect(() =>
      applyWorkbookTransaction(merged, {
        baseline: before,
        source: 'user',
        transactionId: 'delete',
        operations: deleted.operations,
      }),
    ).toThrow('modified-line');
  });

  it('rejects worksheet deletions that miss newer references and preserves unrelated edits', () => {
    const baseline = createWorkbook([['old', 'old']]);
    baseline.worksheets.push({ id: 'data', name: 'Sales Data', data: [['7']] });
    baseline.worksheets.push({ id: 'notes', name: 'Notes', data: [['3']] });
    const deleted = deleteWorksheet(baseline, 'data');
    const transaction = {
      baseline,
      source: 'user' as const,
      transactionId: 'delete',
      operations: commandOperations(baseline, deleted),
    };
    const newFormula = writeCells(baseline, [["  =SUM('sales data'!A1:A2)"]], {
      sheetId,
      x: 0,
      y: 0,
    });
    const newComment = addComment(
      baseline,
      { sheetId: 'data', range: 'A1', text: 'Keep this note' },
      stamp('thread'),
    );
    const newSuggestion = proposeWorkbookChange(
      baseline,
      { proposed: newFormula, summary: 'Use sales data' },
      stamp('suggestion'),
    );
    for (const latest of [newFormula, newComment, newSuggestion])
      expect(() => applyWorkbookTransaction(latest, transaction)).toThrow(
        'referential-delete-baseline',
      );

    const independent = writeCells(baseline, [['=Notes!A1', '="Sales Data!A1"']], {
      sheetId,
      x: 0,
      y: 0,
    });
    expect(applyWorkbookTransaction(independent, transaction)).toStrictEqual(
      writeCells(deleted, [['=Notes!A1', '="Sales Data!A1"']], { sheetId, x: 0, y: 0 }),
    );
  });

  it('performs a contiguous deletion without cloning the workbook once per line', () => {
    const makeId = idFactory();
    const before = initializeWorkbookIdentities(
      createWorkbook(Array.from({ length: 300 }, () => Array<string>(20).fill(''))),
      makeId,
    );
    before.worksheets[0].minDimensions = [20, 300];
    const deleted = changeWorkbookStructureWithOperations(
      before,
      { sheetId, idFactory: makeId },
      { axis: 'row', kind: 'delete', index: 10, count: 100 },
    );
    const clone = vi.spyOn(globalThis, 'structuredClone');
    try {
      const replay = applyWorkbookTransaction(before, {
        transactionId: 'delete',
        source: 'user',
        baseline: before,
        operations: deleted.operations,
      });
      expect(clone.mock.calls.length).toBeLessThan(10);
      expect(replay).toStrictEqual(deleted.workbook);
    } finally {
      clone.mockRestore();
    }
  });
});
