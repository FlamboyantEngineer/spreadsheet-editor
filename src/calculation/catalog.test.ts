import { readFileSync } from 'node:fs';
import { describe, it, expect } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook, SUPPORTED_FUNCTIONS } from './evaluator.js';
import { formatCellValue } from './format.js';

const evaluate = (formula: string) =>
  calculateWorkbook(createWorkbook([[formula]])).worksheets[0].values[0][0];
describe('SE-CR-2.6f: enabled function catalog', () => {
  it('covers every supported worksheet name and alias at the reviewed commit', () => {
    const { worksheetFunctions } = JSON.parse(
      readFileSync(new URL('../../specs/FUNCTION-CATALOG.json', import.meta.url), 'utf8'),
    ) as { worksheetFunctions: string[] };
    expect(worksheetFunctions.filter((name) => !SUPPORTED_FUNCTIONS.includes(name))).toEqual([]);
    expect(SUPPORTED_FUNCTIONS.includes('ACCRINT')).toBe(false);
    expect(evaluate('=ACCRINT(1,2,3)')).toMatchObject({ code: '#NAME?' });
  });
  it.each([
    ['=COUNTIF(A2:A4,">1")', '2'],
    ['=SUMIF(A2:A4,">1",B2:B4)', '50'],
    ['=SUMIFS(B2:B4,A2:A4,">1",A2:A4,"<3")', '20'],
    ['=AVERAGEIF(A2:A4,">1",B2:B4)', '25'],
    ['=MAXIFS(B2:B4,A2:A4,"<=2")', '20'],
    ['=SUMPRODUCT(A2:A4,B2:B4)', '140'],
    ['=MEDIAN(A2:A4)', '2'],
    ['=LARGE(B2:B4,2)', '20'],
    ['=RANK.AVG(2,A2:A4)', '2'],
    ['=INDEX(B2:B4,2)', '20'],
    ['=MATCH(2,A2:A4,0)', '2'],
    ['=VLOOKUP(2,A2:B4,2,FALSE)', '20'],
    ['=LOOKUP(2.5,A2:A4,B2:B4)', '20'],
  ])('evaluates %s against ranges and typed values', (formula, expected) => {
    const result = calculateWorkbook(
      createWorkbook([[formula], ['1', '10'], ['2', '20'], ['3', '30']]),
    );
    expect(formatCellValue(result.worksheets[0].values[0][0])).toBe(expected);
  });
  it.each([
    ['=SWITCH(2,1,1/0,2,7,0)', '7'],
    ['=CEILING.MATH(-2.4)', '-2'],
    ['=CEILING.PRECISE(2.4,)', '0'],
    ['=FLOOR.MATH(-2.4,1,1)', '-2'],
    ['=MROUND(7,2)', '8'],
    ['=EVEN(3)', '4'],
    ['=ODD(-2)', '-3'],
    ['=CONVERT(1,"in","mm")', '25.4'],
    ['=PROPER("hELLO WORLD")', 'Hello World'],
    ['=SUBSTITUTE("a-a-a","a","b",2)', 'a-b-a'],
    ['=REPLACE("abc",2,1,"Z")', 'aZc'],
    ['=SEARCH("b?","aBcde")', '2'],
    ['=TEXTJOIN("-",TRUE,"a","","b")', 'a-b'],
    ['=UNICODE("é")', '233'],
    ['=AVG(1,2,3)', '2'],
    ['=TEXT(0.125,"0.0%")', '12.5%'],
  ])('evaluates %s', (formula, expected) => {
    expect(formatCellValue(evaluate(formula))).toBe(expected);
  });
  it('preserves tiny nonzero inputs in approximate expm1 and log1p', () => {
    for (const formula of ['=EXPM1(1E-1000)', '=LOG1P(1E-1000)']) {
      const value = evaluate(formula);
      expect(value).toMatchObject({ kind: 'number', approximate: true });
      expect(formatCellValue(value)).toBe('≈1e-1000');
    }
  });
});
