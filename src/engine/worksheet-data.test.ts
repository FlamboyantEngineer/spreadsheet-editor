import { describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { gridWorksheetExtent, syncWorksheetData } from './worksheet-data.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

describe('SE-UI-1: logical vendor worksheet dimensions', () => {
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
