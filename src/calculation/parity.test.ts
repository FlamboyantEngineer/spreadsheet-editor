import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook } from './evaluator.js';
import { formatCellValue } from './format.js';

const value = (formula: string) =>
  calculateWorkbook(createWorkbook([[formula]])).worksheets[0].values[0][0];
describe('SE-CR-2.6f: bounded worksheet evaluation', () => {
  it('resumes a 1,200-cell chain and evaluates shared host calls once', () => {
    const data = Array.from({ length: 1200 }, (_, i) => [i === 1199 ? '=TICK()' : `=A${i + 2}+1`]);
    data[0].push('=A1+A2');
    let calls = 0;
    const result = calculateWorkbook(
      createWorkbook(data),
      {},
      {
        callFunction: (name) =>
          name === 'TICK'
            ? (calls++, { kind: 'number', numerator: '1', denominator: '1', approximate: false })
            : undefined,
      },
    );
    expect(formatCellValue(result.worksheets[0].values[0][0])).toBe('1200');
    expect(formatCellValue(result.worksheets[0].values[0][1])).toBe('2399');
    expect(calls).toBe(1);
  });
  it('retains lazy branches, error inspections and literal string types', () => {
    for (const formula of ['=IFNA(NA(),7)', '=CHOOSE(2,1/0,7)', '=IFS(FALSE,1/0,TRUE,7)'])
      expect(formatCellValue(value(formula))).toBe('7');
    expect(value('=ISNA(NA())')).toBe(true);
    expect(value('=ISERR(NA())')).toBe(false);
    expect(value('=IFNA(1/0,7)')).toMatchObject({ code: '#DIV/0!' });
    expect(value('="001 "')).toBe('001 ');
    expect(value('=1<"0"')).toBe(true);
  });
  it('coerces aggregate literals and ignores text and logical references', () => {
    expect(formatCellValue(value('=COUNT(TRUE,"2",NA(),"x")'))).toBe('2');
    expect(formatCellValue(value('=MIN(TRUE,"2")'))).toBe('1');
    expect(formatCellValue(value('=COUNTA("text",NA())'))).toBe('2');
    const result = calculateWorkbook(
      createWorkbook([[true, "'2", '=COUNT(A1:B1)', '=MIN(A1:B1)']]),
    );
    expect(formatCellValue(result.worksheets[0].values[0][2])).toBe('0');
    expect(formatCellValue(result.worksheets[0].values[0][3])).toBe('0');
  });
  it('bounds cumulative reference reads before allocation and cannot hide exhaustion', () => {
    const result = calculateWorkbook(createWorkbook([['=IFERROR(SUM(B1:B6000)+SUM(C1:C6000),0)']]));
    expect(result.worksheets[0].values[0][0]).toMatchObject({ code: '#LIMIT!' });
    expect(value('=IF(FALSE,SUM(A1:XFD1048576),7)')).toMatchObject({ numerator: '7' });
  });
});
