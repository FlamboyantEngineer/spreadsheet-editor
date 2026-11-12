import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook } from './evaluator.js';
import { createArithmetic } from './rational.js';
import { formatWorksheetCell } from './cell-format.js';
import { makeQuantity, quantityInUnit } from './units.js';

const numeric = createArithmetic();
const evaluate = (formula: string) =>
  calculateWorkbook(createWorkbook([[formula]])).worksheets[0].values[0][0];

describe('SE-CALC-4: dimensionless function arguments and compatible coordinates', () => {
  it.each([
    '=LEFT("abc",1 m)',
    '=ROUND(1,1 rad)',
    '=REPT("x",1 bit)',
    '=CHOOSE(1 rad,"first","second")',
    '=DATE(1 rad,1,1)',
    '=TIME(1 bit,0,0)',
    '=DAYS(2 rad,1 rad)',
    '=YEAR(1 rad)',
    '=MONTH(1 bit)',
    '=EDATE(1 rad,0)',
    '=WEEKDAY(1 bit)',
    '=RADIANS(1 rad)',
    '=DEGREES(1 bit)',
    '=CONVERT(1 m,"m","cm")',
    '=AVERAGE(0,1 m)',
    '=AVERAGE(1 m,0)',
    '=AVERAGE(0,1 rad)',
  ])('rejects quantities in a scalar argument: %s', (formula) => {
    expect(evaluate(formula)).toMatchObject({ code: '#VALUE!' });
  });

  it('evaluates ATAN2 with compatible coordinate quantities and rejects incompatible ones', () => {
    const result = evaluate('=ATAN2(1 m,100 cm)');
    expect(formatWorksheetCell(result, { kind: 'number', decimals: 6 })).toBe('≈0.785398');
    expect(evaluate('=ATAN2(1 m,1 s)')).toMatchObject({ code: '#VALUE!' });
    expect(evaluate('=ATAN2(0 m,0 m)')).toMatchObject({ code: '#DIV/0!' });
    const resultWithPins = calculateWorkbook(
      createWorkbook([['=ATAN2(x,y)']]),
      {},
      {
        readName: (name) => ({
          ...makeQuantity(numeric.parse('1'), 'm'),
          provenance: { [name]: 'v1' },
        }),
      },
    );
    expect(resultWithPins.worksheets[0].values[0][0]).toMatchObject({
      provenance: { x: 'v1', y: 'v1' },
    });
  });

  it.each([
    ['=CONVERT(100,"cel","fah")', '212'],
    ['=CONVERT(273.15,"kel","cel")', '0'],
    ['=CONVERT(1,"ozm","kg")', '0.028349523125'],
    ['=CONVERT(1,"meter","mm")', '1000'],
    ['=CONVERT(1,"ang","m")', '1e-10'],
    ['=CONVERT(1,"l","m^3")', '0.001'],
    ['=CONVERT(1,"lt","m^3")', '0.001'],
  ])('recognizes the Excel-specific conversion alias in %s', (formula, expected) => {
    expect(formatWorksheetCell(evaluate(formula))).toBe(expected);
  });

  it('formats canonical and authored non-SI quantities without discarding their physical kind', () => {
    const workbook = createWorkbook([['1', '1 rad', '=1 rad', '8']]);
    workbook.worksheets[0].formats = {
      A1: { kind: 'unit', unit: 'rad', expectedDimension: 'Angle' },
      B1: { kind: 'unit', unit: 'rad' },
      C1: { kind: 'unit', unit: 'rad' },
      D1: { kind: 'unit', unit: 'B', expectedDimension: 'Information' },
    };
    workbook.worksheets[0].inputValues = { D1: makeQuantity(numeric.parse('8'), 'bit') };
    const values = calculateWorkbook(workbook).worksheets[0].values[0];
    values.slice(0, 3).forEach((value) => {
      expect(value).toMatchObject({ numerator: '1', quantityKind: 'angle' });
      expect(formatWorksheetCell(value, { kind: 'unit', unit: 'rad' })).toBe('1 rad');
    });
    expect(formatWorksheetCell(values[3], { kind: 'unit', unit: 'B' })).toBe('1 B');
    expect(() => quantityInUnit(numeric.parse('1'), 'rad')).toThrow('#VALUE!');
    expect(() => quantityInUnit(makeQuantity(numeric.parse('1'), 'bit'), 'rad')).toThrow('#VALUE!');
  });
});
