import { describe, expect, it } from 'vitest';
import { createWorkbook } from './workbook.js';
import { worksheetExtent } from './extent.js';
import { borderCells, clearSelection, formatCells, mergeSelection } from './editing.js';

const range = { x1: 0, y1: 0, x2: 5, y2: 5 };

describe('SE-CHROME-1: borders share the cell style spelling', () => {
  it('leaves no empty style entries and keeps the extent of an unstyled range', () => {
    const workbook = createWorkbook([['1']]);
    const cleared = borderCells(workbook, 'sheet-1', range, 'none');
    expect(cleared.worksheets[0].style).toEqual({});
    expect(worksheetExtent(cleared.worksheets[0], cleared)).toEqual({ columns: 1, rows: 1 });
    const removed = borderCells(
      borderCells(workbook, 'sheet-1', range, 'all'),
      'sheet-1',
      range,
      'all',
    );
    expect(removed.worksheets[0].style).toEqual({});
  });

  it('writes the same style text as formatting for the same cell state', () => {
    const workbook = createWorkbook([['1']]);
    const cell = { x1: 0, y1: 0, x2: 0, y2: 0 };
    const bordered = borderCells(
      formatCells(workbook, 'sheet-1', cell, { 'font-weight': 'bold' }),
      'sheet-1',
      cell,
      'top',
    );
    const formatted = formatCells(workbook, 'sheet-1', cell, {
      'font-weight': 'bold',
      'border-top': '1px solid currentColor',
    });
    expect(bordered.worksheets[0].style).toEqual(formatted.worksheets[0].style);
    expect(bordered.worksheets[0].style!.A1).toBe(
      'font-weight: bold; border-top: 1px solid currentColor',
    );
  });
});

describe('SE-CHROME-2: commands name an unknown worksheet', () => {
  it('throws a range error instead of failing inside the command', () => {
    const workbook = createWorkbook([['1']]);
    for (const command of [
      () => clearSelection(workbook, 'missing', range),
      () => borderCells(workbook, 'missing', range, 'all'),
      () => mergeSelection(workbook, 'missing', range),
    ])
      expect(command).toThrow(new RangeError('Unknown worksheet.'));
  });
});
