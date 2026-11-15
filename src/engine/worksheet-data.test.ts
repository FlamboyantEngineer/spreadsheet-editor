import { describe, expect, it, vi } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import {
  captureWorksheetCellChanges,
  gridWorksheetExtent,
  syncWorksheetData,
} from './worksheet-data.js';

describe('SE-UI-1.1: logical vendor worksheet dimensions', () => {
  it('rejects an oversized vendor batch or far-away write before extending authored data', () => {
    const sheet = createWorkbook([['1']]).worksheets[0];
    const original = structuredClone(sheet);
    expect(() =>
      captureWorksheetCellChanges(sheet, [{ x: 0, y: 1_048_575, value: 'new' }]),
    ).toThrow('cell limit');
    expect(() =>
      captureWorksheetCellChanges(sheet, Array(250_001).fill({ x: 0, y: 0, value: '' })),
    ).toThrow('cell limit');
    sheet.minDimensions = [100, 1000];
    expect(() => captureWorksheetCellChanges(sheet, [{ x: 250, y: 0, value: 'new' }])).toThrow(
      'cell limit',
    );
    expect(sheet.data).toEqual(original.data);
    expect(captureWorksheetCellChanges(sheet, [{ x: 0, y: 0, value: '2' }]).data).toEqual([['2']]);
  });
  it('captures a native batch without authoring viewport padding or stale canonical values', () => {
    const sheet = createWorkbook([['1', '2', '3']]).worksheets[0];
    sheet.inputValues = {
      A1: { kind: 'number', numerator: '1', denominator: '1', approximate: false },
      C1: { kind: 'number', numerator: '3', denominator: '1', approximate: false },
    };
    const original = structuredClone(sheet);
    const next = captureWorksheetCellChanges(sheet, [
      { x: 0, y: 0, value: '' },
      { x: 1, y: 0, value: '' },
      { x: 2, y: 0, value: '3' },
      { x: 25, y: 99, value: '' },
    ]);
    expect(next.data).toEqual([['', '', '3']]);
    expect(next.inputValues).toEqual({ C1: requireDefined(original.inputValues).C1 });
    expect(sheet).toEqual(original);
  });

  it('counts blank metadata edges and retains the initial viewport only while minima are implicit', () => {
    const sheet = createWorkbook().worksheets[0];
    expect(gridWorksheetExtent(sheet)).toEqual([26, 100]);
    sheet.minDimensions = [2, 2];
    sheet.style = { Z100: 'font-weight:bold' };
    sheet.rows = { '100': { height: 32 } };
    sheet.mergeCells = { AA101: [2, 2] };
    expect(gridWorksheetExtent(sheet)).toEqual([28, 102]);
    expect(gridWorksheetExtent(sheet, { columns: 30, rows: 150 })).toEqual([30, 150]);
    expect(sheet.minDimensions).toEqual([2, 2]);
  });

  it('updates minima and column structures before rebuilding body data, without mutating authored inputs', () => {
    const sheet = createWorkbook([['authored']]).worksheets[0];
    sheet.minDimensions = [27, 101];
    sheet.mergeCells = { A1: [2, 1] };
    const original = structuredClone(sheet);
    const options = {
      columns: Array.from({ length: 26 }, () => ({})),
      minDimensions: [26, 100],
      allowInsertColumn: false,
      allowDeleteColumn: false,
      mergeCells: { A1: [2, 1] },
    };
    const insertColumn = vi.fn(() => {
      expect(options.allowInsertColumn).toBe(true);
      expect(options.mergeCells).toEqual({}); // no interactive merge-destruction dialog
    });
    const setData = vi.fn((data: unknown[][]) => {
      expect(options.minDimensions).toEqual([27, 101]);
      expect(options.columns).toHaveLength(27);
      data[0][0] = 'vendor-mutated-copy';
    });
    const engine = {
      options,
      cols: Array(26),
      insertColumn,
      setData,
    } as unknown as JSpreadsheetWorksheetInstance;
    syncWorksheetData(engine, sheet);
    expect(insertColumn).toHaveBeenCalledExactlyOnceWith(1);
    expect(setData).toHaveBeenCalledTimes(1);
    expect(sheet).toEqual(original);
    expect(options.allowInsertColumn).toBe(false);
    expect(options.allowDeleteColumn).toBe(false);
    expect(options.mergeCells).toEqual({ A1: [2, 1] });
  });

  it('shrinks trailing blank columns and rows on delete, undo and controlled replacement', () => {
    const sheet = createWorkbook([[]]).worksheets[0];
    sheet.minDimensions = [25, 99];
    const options = { columns: Array(27).fill({}), minDimensions: [27, 101] };
    const deleteColumn = vi.fn();
    const setData = vi.fn();
    const engine = {
      options,
      cols: Array(27),
      deleteColumn,
      setData,
    } as unknown as JSpreadsheetWorksheetInstance;
    syncWorksheetData(engine, sheet);
    expect(deleteColumn).toHaveBeenCalledExactlyOnceWith(25, 2);
    expect(options.minDimensions).toEqual([25, 99]);
    expect(options.columns).toHaveLength(25);
  });
});
