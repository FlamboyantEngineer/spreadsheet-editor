import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook, type FormulaHost } from './evaluator.js';
import { createArithmetic } from './rational.js';
import { type CellValue, isNumeric } from './values.js';
const corpus = JSON.parse(
  readFileSync(new URL('../../specs/NUMERIC-CORPUS.json', import.meta.url), 'utf8'),
) as { cases: Record<string, [string, string][]> };
const arithmetic = createArithmetic();
function evaluate(formula: string, data: (string | boolean | null)[][] = []): CellValue {
  const workbook = createWorkbook(data.length ? data : [[null]]);
  workbook.worksheets.push({
    id: 'case',
    name: 'Case',
    data: [[formula.startsWith('=') ? formula : '=' + formula]],
  });
  // Unqualified fixture references belong to the fixture sheet.
  workbook.worksheets[1].data[0][0] = String(workbook.worksheets[1].data[0][0]).replace(
    /(?<![A-Za-z0-9_!])([A-Z]+\d+(?::[A-Z]+\d+)?)(?![A-Za-z0-9_(])/g,
    'Sheet1!$1',
  );
  return calculateWorkbook(workbook).worksheets[1].values[0][0];
}
function expectDecimal(value: CellValue, expected: string) {
  expect(isNumeric(value), JSON.stringify(value)).toBe(true);
  if (!isNumeric(value)) return;
  const parsed = arithmetic.parse(expected);
  expect({ numerator: value.numerator, denominator: value.denominator }).toEqual({
    numerator: parsed.numerator,
    denominator: parsed.denominator,
  });
}
describe('SE-CALC-5.1: portable pinned behavior corpus', () => {
  for (const family of ['basic', 'calendar', 'dateOverflow', 'quantities'])
    it.each(corpus.cases[family])(`${family} %s`, (formula, expected) =>
      expectDecimal(evaluate(formula), expected),
    );
  it.each(corpus.cases.text)('text %s', (formula, expected) => {
    const value = evaluate(formula);
    if (
      ['CODE', 'UNICODE', 'FIND', 'SEARCH', 'COUNTIF'].some((name) =>
        formula.startsWith(name + '('),
      )
    )
      expectDecimal(value, expected);
    else expect(value).toBe(expected);
  });
  it.each(corpus.cases.restored)('aggregate/lookup %s', (formula, expected) =>
    expectDecimal(
      evaluate(formula, [
        ['.1', 'yes', '=4 mm'],
        ['.2', 'yes', '=6 mm'],
        ['.3', 'no', '=8 mm'],
      ]),
      expected,
    ),
  );
  it.each(corpus.cases.index)('INDEX shape %s', (formula, expected) =>
    expectDecimal(
      evaluate(formula, [
        ['10', '20', '30', '40', '50'],
        ['1', '2', '3', '4', '5'],
        ['6', '7', '8', '9', '10'],
      ]),
      expected,
    ),
  );
  it('distinguishes blank lookup results from invalid indices and matches text patterns only against text', () => {
    const data = [
      ['1', null, 'item'],
      ['2', null, null],
    ];
    expectDecimal(evaluate('VLOOKUP(1,A1:B1,2,FALSE)', data), '0');
    expect(evaluate('VLOOKUP(1,A1:B1,3,FALSE)', data)).toMatchObject({ code: '#REF!' });
    expectDecimal(evaluate('LOOKUP(2,A1:A2,B1:B2)', data), '0');
    expect(evaluate('MATCH("1",A1:A2,0)', data)).toMatchObject({ code: '#N/A' });
    expectDecimal(evaluate('COUNTIF(A1:C2,"*")', data), '1');
    expectDecimal(evaluate('COUNTIF(A1:C2,"i?em")', data), '1');
  });
  it.each(['SUM(INDEX(A1:A3,0))', 'SUM(INDEX(A1:C3,0,2))'])(
    'preserves quantity rows %s',
    (formula) => {
      const value = evaluate(formula, [
        ['=1 mm', '=1 mm', '=1 mm'],
        ['=2 mm', '=2 mm', '=2 mm'],
        ['=3 mm', '=3 mm', '=3 mm'],
      ]);
      expectDecimal(value, '.006');
      expect(value).toMatchObject({ dimension: [1, 0, 0, 0, 0, 0, 0] });
    },
  );
  it('retains exact rational cancellation as explicit finite-kernel differences', () => {
    expectDecimal(evaluate('INT(1/3*3)'), '1');
    expect(evaluate('1/3*3=1')).toBe(true);
  });
  it('retains lazy branches, budget errors and explicit unsupported/errors', () => {
    for (const [formula, expected] of [
      ['TEXT(1,"unsupported")', '#VALUE!'],
      ['TEXT(1 m,"0")', '#VALUE!'],
      ['ATAN2(0,0)', '#DIV/0!'],
      ['INDEX(1,4)', '#REF!'],
      ['CHOOSE(0,1,2)', '#VALUE!'],
      ['T(1/0)', '#DIV/0!'],
      ['FOO(1)', '#NAME?'],
    ])
      expect(evaluate(formula)).toMatchObject({ code: expected });
    expectDecimal(evaluate('IF(FALSE,SUM(A1:XFD1048576),7)'), '7');
  });
  it('captures formula cell contexts and metadata without converting host values', () => {
    const seen: unknown[] = [];
    const host: FormulaHost = {
      hostContext: { pin: { snapshot: '1' } },
      readName(name, context) {
        seen.push([name, context]);
        return { ...arithmetic.parse('9007199254740993'), provenance: { constants: '2022' } };
      },
      callFunction(name, args, context) {
        seen.push([name, args, context]);
        return arithmetic.parse('1');
      },
    };
    const result = calculateWorkbook(createWorkbook([['=Foo+BAR()']]), {}, host);
    expectDecimal(result.worksheets[0].values[0][0], '9007199254740994');
    expect(result.worksheets[0].values[0][0]).toMatchObject({ provenance: { constants: '2022' } });
    expect(seen).toEqual([
      ['Foo', { sheetId: 'sheet-1', x: 0, y: 0, formula: '=Foo+BAR()' }],
      ['BAR', [], { sheetId: 'sheet-1', x: 0, y: 0, formula: '=Foo+BAR()' }],
    ]);
    expect(result.context?.hostContext).toEqual({ pin: { snapshot: '1' } });
  });
});
