import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { formatWorksheetCell } from './cell-format.js';
import { calculateWorkbook } from './evaluator.js';

describe('SE-EDIT-2: number formats preserve exact calculations', () => {
  it('formats a near-one value as percent without altering its comparison result', () => {
    const result = calculateWorkbook(createWorkbook([['=1-1E-25', '=A1=1']]));
    const value = result.worksheets[0].values[0][0];
    expect(formatWorksheetCell(value, { kind: 'percent', decimals: 2 })).toBe('100.00%');
    expect(result.worksheets[0].values[0][1]).toBe(false);
    expect(formatWorksheetCell(value)).toBe('0.9999999999999999999999999');
  });
  it('formats tiny scientific, fraction and high-decimal values without binary floating point', () => {
    const value = calculateWorkbook(createWorkbook([['=1E-25']])).worksheets[0].values[0][0];
    expect(formatWorksheetCell(value, { kind: 'scientific', decimals: 2 })).toBe('1.00e-25');
    expect(formatWorksheetCell(value, { kind: 'number', decimals: 25 })).toBe(
      '0.0000000000000000000000001',
    );
    expect(formatWorksheetCell(value, { kind: 'fraction' })).toBe('1/10000000000000000000000000');
  });
});
