import { describe, expect, it } from 'vitest';
import { borderCells, clearSelection, formatCells, mergeSelection } from './editing.js';
import { worksheetExtent } from './extent.js';
import { requireDefined } from './require-defined.js';
import { createWorkbook } from './workbook.js';

const range = { x1: 0, y1: 0, x2: 5, y2: 5 };

describe('SE-CHROME-1.1: borders share the cell style spelling', () => {
  it('leaves no empty style entries and keeps the extent of an unstyled range', () => {
    const workbook = createWorkbook([['1']]);
    const cleared = borderCells(workbook, { sheetId: 'sheet-1', range }, 'none');
    expect(cleared.worksheets[0].style).toEqual({});
    expect(worksheetExtent(cleared.worksheets[0], cleared)).toEqual({ columns: 1, rows: 1 });
    const removed = borderCells(
      borderCells(workbook, { sheetId: 'sheet-1', range }, 'all'),
      { sheetId: 'sheet-1', range },
      'all',
    );
    expect(removed.worksheets[0].style).toEqual({});
  });

  it('writes the same style text as formatting for the same cell state', () => {
    const workbook = createWorkbook([['1']]);
    const cell = { x1: 0, y1: 0, x2: 0, y2: 0 };
    const bordered = borderCells(
      formatCells(
        workbook,
        { sheetId: 'sheet-1', range: cell },
        { properties: { 'font-weight': 'bold' } },
      ),
      { sheetId: 'sheet-1', range: cell },
      'top',
    );
    const formatted = formatCells(
      workbook,
      { sheetId: 'sheet-1', range: cell },
      {
        properties: {
          'font-weight': 'bold',
          'border-top': '1px solid currentColor',
        },
      },
    );
    expect(bordered.worksheets[0].style).toEqual(formatted.worksheets[0].style);
    expect(requireDefined(bordered.worksheets[0].style).A1).toBe(
      'font-weight: bold; border-top: 1px solid currentColor',
    );
  });
});

describe('SE-CHROME-2.1: commands name an unknown worksheet', () => {
  it('throws a range error instead of failing inside the command', () => {
    const workbook = createWorkbook([['1']]);
    for (const command of [
      () => clearSelection(workbook, { sheetId: 'missing', range }),
      () => borderCells(workbook, { sheetId: 'missing', range }, 'all'),
      () => mergeSelection(workbook, { sheetId: 'missing', range }),
    ])
      expect(command).toThrow(new RangeError('Unknown worksheet.'));
  });
});
