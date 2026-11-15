import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook } from './evaluator.js';
import { formatCellValue } from './format.js';

const evaluate = (formula: string) =>
  calculateWorkbook(createWorkbook([[formula]])).worksheets[0].values[0][0];

describe('SE-CALC-4: calendar month shifts stay inside supported serial bounds', () => {
  it.each([
    ['=EDATE(0,0)', '0'],
    ['=EOMONTH(0,0)', '0'],
    ['=EDATE(0,1)', '31'],
    ['=EDATE(31,-1)', '0'],
    ['=EOMONTH(1,-1)', '0'],
    ['=EDATE(DATE(9999,12,31),0)', '2958465'],
    ['=EOMONTH(DATE(9999,12,1),0)', '2958465'],
    ['=EDATE(DATE(9999,11,30),1)', '2958464'],
  ])('keeps the boundary result of %s as %s', (formula, expected) => {
    expect(formatCellValue(evaluate(formula))).toBe(expected);
  });

  it.each([
    '=EDATE(1,-1)',
    '=EOMONTH(1,-2)',
    '=EDATE(DATE(9999,12,1),1)',
    '=EOMONTH(DATE(9999,12,1),1)',
  ])('rejects a month shift outside the serial domain: %s', (formula) => {
    expect(evaluate(formula)).toMatchObject({ code: '#NUM!' });
  });
});
