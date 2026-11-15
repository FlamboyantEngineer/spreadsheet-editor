import { describe, expect, it } from 'vitest';
import type { NumericValue } from '../calculation/values.js';
import { createSpreadsheetCopy, pasteSpreadsheetCopy } from './clipboard.js';
import { clearSelection, formatCells, parseDelimited, writeCells } from './editing.js';
import { worksheetExtent } from './extent.js';
import {
  changeWorkbookStructureWithOperations,
  collectCellOperations,
  invertWorkbookOperations,
  writeCellsWithOperations,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { decideSuggestion, previewSuggestions, proposeWorkbookChange } from './review.js';
import { changeWorkbookStructure } from './structure.js';
import { assertWorkbook, createWorkbook } from './workbook.js';

const rational: NumericValue = {
  kind: 'number',
  numerator: '1',
  denominator: '10000000000000000000000000',
  approximate: false,
  provenance: { catalog: 'revision-1' },
};
const stamp = { id: 'proposal', time: '2026-11-11T12:00:00Z', actor: { id: 'alex', name: 'Alex' } };

describe('SE-HOST-MODEL-1.1: canonical authored quantities and stable extent', () => {
  it('roundtrips canonical values and validates nonformula targets and rational metadata', () => {
    const workbook = createWorkbook([['authored in another unit']]);
    workbook.worksheets[0].inputValues = { A1: rational };
    expect(() => assertWorkbook(JSON.parse(JSON.stringify(workbook)))).not.toThrow();
    for (const [address, data] of [
      ['B1', [['x']]],
      ['A1', [['=1+1']]],
    ] as const) {
      const bad = createWorkbook(data.map((row) => [...row]));
      bad.worksheets[0].inputValues = { [address]: rational };
      expect(() => assertWorkbook(bad)).toThrow('corresponding nonformula');
    }
    const bad = structuredClone(workbook);
    requireDefined(bad.worksheets[0].inputValues).A1 = { ...rational, denominator: '0' };
    expect(() => assertWorkbook(bad)).toThrow('numeric');
    expect(() => assertWorkbook(workbook, { maxCells: 1 })).not.toThrow();
  });
  it('ordinary replacement and clear discard canonical values; formatting retains them', () => {
    const workbook = createWorkbook([['tiny']]);
    workbook.worksheets[0].inputValues = { A1: rational };
    expect(
      writeCells(workbook, 'sheet-1', 0, 0, [['2']]).worksheets[0].inputValues?.A1,
    ).toBeUndefined();
    expect(
      writeCells(workbook, 'sheet-1', 0, 0, [['3']]).worksheets[0].inputValues?.A1,
    ).toBeUndefined();
    expect(
      clearSelection(workbook, 'sheet-1', { x1: 0, y1: 0, x2: 0, y2: 0 }).worksheets[0].inputValues
        ?.A1,
    ).toBeUndefined();
    expect(
      formatCells(workbook, 'sheet-1', { x1: 0, y1: 0, x2: 0, y2: 0 }, { 'font-weight': 'bold' })
        .worksheets[0].inputValues?.A1,
    ).toEqual(rational);
    expect(workbook.worksheets[0].inputValues?.A1).toEqual(rational);
  });
  it('counts sparse geometry, merges, formats and annotation anchors in the effective extent', () => {
    const workbook = createWorkbook();
    const sheet = workbook.worksheets[0];
    sheet.columns = Array.from({ length: 5 }, () => ({}));
    sheet.rows = { '8': { height: 30 } };
    sheet.formats = { H12: { kind: 'number' } };
    sheet.mergeCells = { J15: [2, 3] };
    workbook.comments = [
      {
        id: 'thread',
        sheetId: sheet.id,
        range: 'M20',
        state: 'open',
        entries: [{ id: 'post', author: stamp.actor.id, text: 'edge', createdAt: stamp.time }],
      },
    ];
    expect(worksheetExtent(sheet, workbook)).toEqual({ columns: 13, rows: 20 });
    sheet.rowIds = Array.from({ length: 19 }, (_, i) => `r${i}`);
    expect(() => assertWorkbook(workbook)).toThrow('cover the effective extent');
    sheet.rowIds.push('r19');
    expect(() => assertWorkbook(workbook)).not.toThrow();
    sheet.rowIds[1] = 'r0';
    expect(() => assertWorkbook(workbook)).toThrow('identities');
  });
  it('moves IDs/canonical inputs and preserves metadata on a mostly empty edge', () => {
    const workbook = createWorkbook([['1'], [], ['edge']]);
    const sheet = workbook.worksheets[0];
    sheet.minDimensions = [3, 3];
    sheet.rowIds = ['r0', 'r1', 'r2'];
    sheet.columnIds = ['c0', 'c1', 'c2'];
    sheet.inputValues = { A3: rational };
    sheet.formats = { C3: { kind: 'currency', locale: 'de-DE', currency: 'EUR' } };
    const inserted = changeWorkbookStructure(workbook, sheet.id, {
      axis: 'row',
      kind: 'insert',
      index: 2,
      count: 1,
      ids: ['fresh'],
    });
    expect(inserted.worksheets[0].minDimensions).toEqual([3, 4]);
    expect(inserted.worksheets[0].rowIds).toEqual(['r0', 'r1', 'fresh', 'r2']);
    expect(inserted.worksheets[0].inputValues?.A4).toEqual(rational);
    expect(inserted.worksheets[0].formats?.C4?.currency).toBe('EUR');
    const deleted = changeWorkbookStructure(inserted, sheet.id, {
      axis: 'row',
      kind: 'delete',
      index: 2,
      count: 1,
    });
    expect(deleted.worksheets[0].rowIds).toEqual(sheet.rowIds);
    expect(deleted.worksheets[0].minDimensions).toEqual([3, 3]);
    expect(deleted.worksheets[0].inputValues?.A3).toEqual(rational);
  });
  it('rejects invalid locale/currency and out-of-budget sparse extents', () => {
    const workbook = createWorkbook();
    workbook.worksheets[0].formats = { A1: { kind: 'currency', currency: 'ZZZ' } };
    expect(() => assertWorkbook(workbook)).toThrow('currency');
    workbook.worksheets[0].formats = { A1: { kind: 'currency', locale: 'xx-YY' } };
    expect(() => assertWorkbook(workbook)).toThrow('locale');
    workbook.worksheets[0].formats = { XFD100: { kind: 'number' } };
    expect(() => assertWorkbook(workbook)).toThrow('extent');
  });
});

describe('SE-HOST-MODEL-1.1: structure commands on a tall, narrow sheet', () => {
  it('inserts rows and columns without padding the sheet past the cell limit', () => {
    const workbook = createWorkbook(
      Array.from({ length: 10_000 }, (_, y) => Array.from({ length: 20 }, () => String(y))),
    );
    const ids = (kind: string) => `${kind}-${Math.random()}`;
    for (const [axis, size] of [
      ['row', { columns: 20, rows: 10_001 }],
      ['column', { columns: 21, rows: 10_000 }],
    ] as const) {
      const { workbook: next } = changeWorkbookStructureWithOperations(
        workbook,
        'sheet-1',
        { axis, kind: 'insert', index: 1, count: 1 },
        ids,
      );
      expect(worksheetExtent(next.worksheets[0], next)).toEqual(size);
    }
    const small = changeWorkbookStructureWithOperations(
      createWorkbook([['1']]),
      'sheet-1',
      { axis: 'row', kind: 'insert', index: 0, count: 1 },
      ids,
    ).workbook;
    expect(worksheetExtent(small.worksheets[0], small)).toEqual({ columns: 26, rows: 101 });
  });
});

describe('SE-HOST-MODEL-2.1: semantic transactions', () => {
  it('SE-PERF-2.1: compares cells in place and reports only those that differ', () => {
    const before = createWorkbook([
      ['1', '2', '3'],
      ['4', '5', '6'],
    ]);
    before.worksheets[0].style = { A1: 'color: red', C2: 'color: blue' };
    before.worksheets[0].cellProvenance = {
      B1: { actor: { id: 'alex', name: 'Alex' }, source: 'human' },
    };
    const after = structuredClone(before),
      sheet = after.worksheets[0];
    expect(collectCellOperations(before, after)).toEqual([]);
    sheet.data[1][0] = '40';
    sheet.style = { A1: 'color: red', B2: 'color: green' };
    sheet.cellProvenance = { B1: { actor: { id: 'sam', name: 'Sam' }, source: 'human' } };
    sheet.inputValues = { C1: rational };
    const [operation] = collectCellOperations(before, after);
    expect(operation.kind === 'cells' && operation.changes.map((change) => change.address)).toEqual(
      ['A2', 'C2', 'B2', 'C1', 'B1'],
    );
    const targeted = collectCellOperations(before, after, new Map([['sheet-1', ['A2', 'A1']]]));
    expect(targeted).toEqual([
      {
        kind: 'cells',
        sheetId: 'sheet-1',
        changes: [expect.objectContaining({ address: 'A2', after: { value: '40' } })],
      },
    ]);
  });
  it('uses the host identity factory when ordinary writes extend either axis', () => {
    const workbook = createWorkbook([['original']]);
    workbook.worksheets[0].rowIds = ['row-original'];
    workbook.worksheets[0].columnIds = ['column-original'];
    let sequence = 0;
    const factory = (kind: string) => `${kind}-host-${++sequence}`;
    const result = writeCellsWithOperations(workbook, 'sheet-1', 1, 1, [['new']], factory);
    expect(result.workbook.worksheets[0].rowIds).toEqual(['row-original', 'row-host-1']);
    expect(result.workbook.worksheets[0].columnIds).toEqual(['column-original', 'column-host-2']);
    expect(result.operations).toMatchObject([
      { kind: 'lines', axis: 'row', ids: ['row-host-1'] },
      { kind: 'lines', axis: 'column', ids: ['column-host-2'] },
      {
        kind: 'cells',
        changes: [{ address: 'B2', rowId: 'row-host-1', columnId: 'column-host-2' }],
      },
    ]);
    const single = writeCells(workbook, 'sheet-1', 0, 1, [['new']], factory);
    expect(single.worksheets[0].rowIds).toEqual(['row-original', 'row-host-3']);
    expect(workbook.worksheets[0].rowIds).toEqual(['row-original']);
    expect(workbook.worksheets[0].columnIds).toEqual(['column-original']);
  });
  it('records a blank final-line insert from its command and retains explicitly generated IDs', () => {
    const workbook = createWorkbook();
    let sequence = 0;
    const result = changeWorkbookStructureWithOperations(
      workbook,
      'sheet-1',
      { axis: 'row', kind: 'insert', index: 99, count: 1 },
      () => `new-${sequence++}`,
    );
    expect(result.workbook.worksheets[0].minDimensions).toEqual([26, 101]);
    expect(result.operations[0]).toMatchObject({
      kind: 'lines',
      axis: 'row',
      action: 'insert',
      index: 99,
      count: 1,
      ids: ['new-0'],
    });
    expect(result.operations).toHaveLength(1);
    expect(workbook.worksheets[0].minDimensions).toBeUndefined();
  });
  it('includes cross-sheet reference changes in the same transaction without duplicate moved-cell writes', () => {
    const workbook = createWorkbook([['1'], ['=A1']]);
    workbook.worksheets[0].minDimensions = [1, 2];
    workbook.worksheets[0].rowIds = ['r1', 'r2'];
    workbook.worksheets.push({ id: 'other', name: 'Other', data: [['=Sheet1!A2']] });
    const result = changeWorkbookStructureWithOperations(workbook, 'sheet-1', {
      axis: 'row',
      kind: 'insert',
      index: 1,
      count: 1,
      ids: ['fresh'],
    });
    expect(result.operations).toHaveLength(2);
    expect(result.operations[0]).toMatchObject({
      kind: 'lines',
      ids: ['fresh'],
      beforeId: 'r1',
      afterId: 'r2',
    });
    expect(result.operations[1]).toMatchObject({
      kind: 'cells',
      sheetId: 'other',
      changes: [{ address: 'A1', before: { value: '=Sheet1!A2' }, after: { value: '=Sheet1!A3' } }],
    });
  });
  it('records bounded targeted cell metadata before and after, including canonical replacement', () => {
    const workbook = createWorkbook([['tiny']]);
    workbook.worksheets[0].inputValues = { A1: rational };
    const result = writeCellsWithOperations(workbook, 'sheet-1', 0, 0, [['replacement']]);
    expect(result.operations).toMatchObject([
      {
        kind: 'cells',
        sheetId: 'sheet-1',
        changes: [
          {
            address: 'A1',
            before: { value: 'tiny', inputValue: rational },
            after: { value: 'replacement' },
          },
        ],
      },
    ]);
    expect(
      'inputValue' in
        ((result.operations[0].kind === 'cells' && result.operations[0].changes[0].after) || {}),
    ).toBe(false);
  });
});

describe('SE-HOST-MODEL-2.1: inverse command metadata and bounded foreign input', () => {
  it('retains removed authored metadata for a semantic line undo', () => {
    const workbook = createWorkbook([['first'], ['tiny'], ['last']]);
    const sheet = workbook.worksheets[0];
    sheet.minDimensions = [1, 3];
    sheet.rowIds = ['r0', 'r1', 'r2'];
    sheet.inputValues = { A2: rational };
    sheet.rows = { '1': { height: 45 } };
    const deleted = changeWorkbookStructureWithOperations(workbook, sheet.id, {
      axis: 'row',
      kind: 'delete',
      index: 1,
      count: 1,
    });
    const inverse = invertWorkbookOperations(deleted.operations);
    expect(inverse[0]).toMatchObject({
      kind: 'lines',
      action: 'insert',
      ids: ['r1'],
      beforeId: 'r0',
      afterId: 'r2',
    });
    expect(inverse).toContainEqual(
      expect.objectContaining({
        kind: 'cells',
        changes: expect.arrayContaining([
          expect.objectContaining({
            address: 'A2',
            rowId: 'r1',
            after: expect.objectContaining({ value: 'tiny', inputValue: rational }),
          }),
        ]),
      }),
    );
    expect(inverse).toContainEqual(
      expect.objectContaining({
        kind: 'geometry',
        after: expect.objectContaining({ rows: { '1': { height: 45 } } }),
      }),
    );
  });
  it('rejects unbounded foreign fields and columns before creating a workbook', () => {
    expect(() =>
      parseDelimited(Array.from({ length: 600 }, () => '\u754c'.repeat(10000)).join('\n')),
    ).toThrow('byte limit');
    expect(() => parseDelimited('a'.repeat(10001))).toThrow('text limit');
    expect(() => parseDelimited(Array.from({ length: 16385 }, () => 'x').join('\t'))).toThrow(
      'cell limit',
    );
  });
});

describe('SE-HOST-MODEL-3.1 and SE-REVIEW-8.1: authored clipboard and canonical review baselines', () => {
  it('tiles exact multiples with independent reference offsets and copies canonical metadata', () => {
    const workbook = createWorkbook([['1', '=A1+$A1+A$1+$A$1']]);
    const sheet = workbook.worksheets[0];
    sheet.inputValues = { A1: rational };
    sheet.formats = { A1: { kind: 'number' } };
    sheet.style = { A1: 'font-weight:bold' };
    const copy = createSpreadsheetCopy(
      workbook,
      sheet.id,
      { x1: 0, y1: 0, x2: 1, y2: 0 },
      'displayed\tresult',
    );
    const pasted = pasteSpreadsheetCopy(workbook, sheet.id, 2, 1, copy, {
      x1: 2,
      y1: 1,
      x2: 5,
      y2: 2,
    }).worksheets[0];
    expect(pasted.data[1].slice(2, 6)).toEqual(['1', '=C2+$A2+C$1+$A$1', '1', '=E2+$A2+E$1+$A$1']);
    expect(pasted.data[2][3]).toBe('=C3+$A3+C$1+$A$1');
    expect(pasted.inputValues?.C2).toEqual(rational);
    expect(pasted.inputValues?.E3).toEqual(rational);
    expect(pasted.style?.E3).toBe('font-weight:bold');
    expect(pasted.formats?.E3?.kind).toBe('number');
    expect(pasted.inputValues?.E3).not.toBe(rational);
  });
  it('pastes one whole block into smaller/nonmultiple selections and moves a cut only once', () => {
    const workbook = createWorkbook([['1', '=A1']]);
    const selection = { x1: 0, y1: 0, x2: 1, y2: 0 },
      copy = createSpreadsheetCopy(workbook, 'sheet-1', selection, '1\t1');
    const pasted = pasteSpreadsheetCopy(workbook, 'sheet-1', 2, 1, copy, {
      x1: 2,
      y1: 1,
      x2: 4,
      y2: 2,
    });
    expect(pasted.worksheets[0].data[1].slice(2)).toEqual(['1', '=C2']);
    expect(pasted.worksheets[0].data[2]).toBeUndefined();
    const cut = createSpreadsheetCopy(workbook, 'sheet-1', selection, '1\t1', true);
    const moved = pasteSpreadsheetCopy(workbook, 'sheet-1', 1, 0, cut, {
      x1: 1,
      y1: 0,
      x2: 4,
      y2: 1,
    });
    expect(moved.worksheets[0].data[0]).toEqual([null, '1', '=A1']);
    expect(moved.worksheets[0].data[1]).toBeUndefined();
    expect(workbook.worksheets[0].data[0]).toEqual(['1', '=A1']);
    expect(() =>
      pasteSpreadsheetCopy(writeCells(workbook, 'sheet-1', 0, 0, [['peer']]), 'sheet-1', 2, 0, cut),
    ).toThrow('cut source');
  });
  it('keeps canonical suggestion baselines through preview/acceptance and detects canonical-only conflicts', () => {
    const workbook = createWorkbook([['tiny']]);
    workbook.worksheets[0].inputValues = { A1: rational };
    const candidate = writeCells(workbook, 'sheet-1', 0, 0, [['2']]);
    const proposal = proposeWorkbookChange(workbook, candidate, 'Replace', stamp);
    expect(proposal.suggestions?.[0].changes[0].beforeInputValue).toEqual(rational);
    expect(previewSuggestions(proposal, 'alex').worksheets[0].inputValues?.A1).toBeUndefined();
    expect(
      decideSuggestion(proposal, 'proposal', 'accepted', { ...stamp, id: 'decision' }).worksheets[0]
        .inputValues?.A1,
    ).toBeUndefined();
    const conflicted = structuredClone(proposal);
    requireDefined(conflicted.worksheets[0].inputValues).A1 = { ...rational, numerator: '2' };
    expect(() =>
      decideSuggestion(conflicted, 'proposal', 'accepted', { ...stamp, id: 'decision' }),
    ).toThrow('changed since');
  });
});
