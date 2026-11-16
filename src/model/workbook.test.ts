import { describe, expect, it } from 'vitest';
import { writeCells } from './editing.js';
import { parseFormula } from './formula-parser.js';
import { assertWorkbook, createWorkbook } from './workbook.js';

describe('SE-CALC-2.1: portable workbook and parser', () => {
  it('edits without mutating source input', () => {
    const original = createWorkbook([['1E-25']]);
    const next = writeCells(original, [['=1-A1']], { sheetId: 'sheet-1', x: 1, y: 0 });
    expect(original.worksheets[0].data).toEqual([['1E-25']]);
    expect(next.worksheets[0].data).toEqual([['1E-25', '=1-A1']]);
  });
  it('rejects malformed cells, duplicate names, and oversized workbooks', () => {
    expect(() => assertWorkbook(createWorkbook([[Infinity]]))).toThrow('cell');
    const duplicate = createWorkbook();
    duplicate.worksheets.push({ id: 'other', name: 'sheet1', data: [] });
    expect(() => assertWorkbook(duplicate)).toThrow('unique');
    expect(() => assertWorkbook(createWorkbook([['a', 'b']]), { maxCells: 1 })).toThrow('limit');
  });
  it('rejects geometry that would exceed the bounded grid area', () => {
    const workbook = createWorkbook();
    workbook.worksheets[0].minDimensions = [16384, 1048576];
    expect(() => assertWorkbook(workbook)).toThrow('dimensions');
    delete workbook.worksheets[0].minDimensions;
    workbook.worksheets[0].mergeCells = { A1: [250000, 250000] };
    expect(() => assertWorkbook(workbook)).toThrow('merged');
  });
  it('extracts scientific literals, absolute addresses, and quoted sheet ranges', () => {
    expect(parseFormula('1E-25')).toEqual({ kind: 'decimal', decimal: '1E-25' });
    expect(parseFormula("'Other sheet'!$A$1:B2")).toMatchObject({
      kind: 'reference',
      sheetName: 'Other sheet',
      start: { x: 0, y: 0 },
      end: { x: 1, y: 1 },
    });
    expect(() => parseFormula('1;process.exit()')).toThrow();
  });
});
