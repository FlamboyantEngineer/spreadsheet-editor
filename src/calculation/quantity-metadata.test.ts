import { describe, expect, it } from 'vitest';
import { assertWorkbook, createWorkbook } from '../model/workbook.js';
import { formatExactDecimal, formatWorksheetCell } from './cell-format.js';
import { calculateWorkbook } from './evaluator.js';
import { assertCalculationResult, createHttpCalculator, handleCalculationRequest } from './http.js';
import { createArithmetic } from './rational.js';
import {
  BINARY_PREFIXES,
  createQuantityArithmetic,
  makeQuantity,
  quantityInUnit,
  resolveUnit,
  SI_DECIMAL_PREFIXES,
  UNIT_DEFINITIONS,
  validateUnitDimension,
} from './units.js';
import { assertJsonValue, assertNumericValue, type NumericValue } from './values.js';
const numeric = createArithmetic(),
  math = createQuantityArithmetic();
const evaluate = (formula: string) =>
  calculateWorkbook(createWorkbook([[formula]])).worksheets[0].values[0][0];
describe('SE-CALC-4.1: complete quantity metadata and registry', () => {
  it('audits every reviewed alias and prefix without binary conversions', () => {
    for (const [category, data] of Object.entries(UNIT_DEFINITIONS))
      for (const unit of data.units)
        for (const alias of new Set([...unit.input, unit.display])) {
          if (category === 'temperature' && ['C', 'F'].includes(alias)) continue;
          expect(
            () => makeQuantity(numeric.parse('1'), alias, { wordSizeBits: '32' }),
            alias,
          ).not.toThrow();
        }
    for (const prefix of SI_DECIMAL_PREFIXES)
      for (const alias of prefix.input) {
        const value = makeQuantity(numeric.parse('1'), alias + 'm');
        expect(value).toMatchObject(
          numeric.power(numeric.parse('10'), numeric.parse(String(prefix.exponent))),
        );
      }
    for (const prefix of BINARY_PREFIXES)
      for (const alias of prefix.input) {
        const value = makeQuantity(numeric.parse('1'), alias + 'B');
        expect(value).toMatchObject({
          ...numeric.multiply(
            numeric.parse('8'),
            numeric.power(numeric.parse('2'), numeric.parse(String(prefix.exponent))),
          ),
          quantityKind: 'information',
        });
      }
    expect(() => resolveUnit('w')).toThrow('#VALUE!');
    expect(resolveUnit('w', { wordSizeBits: '32' }).factor.numerator).toBe('32');
    expect(() => calculateWorkbook(createWorkbook(), { wordSizeBits: '0' })).toThrow(
      'wordSizeBits',
    );
    expect(resolveUnit('C').dimension).toEqual([0, 0, 1, 1, 0, 0, 0]);
    expect(resolveUnit('F').dimension).toEqual([-2, -1, 4, 2, 0, 0, 0]);
  });
  it('applies absolute/difference rules, affine scales, signed literals and averages', () => {
    const difference = evaluate('=30 °C-25 °C');
    expect(difference).toMatchObject({ numerator: '5', temperature: 'difference' });
    expect(formatWorksheetCell(difference, { kind: 'unit', unit: '°C' })).toBe('5 °C');
    expect(formatWorksheetCell(difference, { kind: 'unit', unit: '°F' })).toBe('9 °F');
    expect(evaluate('=AVERAGE(0 °C,100 °C)')).toMatchObject({
      ...numeric.parse('323.15'),
      dimension: [0, 0, 0, 0, 1, 0, 0],
      temperature: 'absolute',
    });
    for (const formula of [
      '=20 °C+10 °C',
      '=20 °C+5 K',
      '=SUM(20 °C,5 °C)',
      '=AVERAGE(20 °C,20 °C-10 °C)',
      '=-(AVERAGE(0 °C,100 °C))',
      '=(20 °C-10 °C)-20 °C',
      '=(20 °C-10 °C)=20 K',
    ])
      expect(evaluate(formula), formula).toMatchObject({ code: '#VALUE!' });
    for (const formula of ['=-40 °C', '=-(40 °C)', '=-40 °F'])
      expect(evaluate(formula), formula).toMatchObject(numeric.parse('233.15'));
    expect(evaluate('=9 °F/s')).toMatchObject({
      ...numeric.parse('5'),
      dimension: [0, 0, -1, 0, 1, 0, 0],
    });
    expect(evaluate('=(20 °C-10 °C)+0 °C')).toMatchObject({
      ...numeric.parse('283.15'),
      temperature: 'absolute',
    });
    expect(evaluate('=STDEV.P(0 °C,100 °C)')).toMatchObject({
      ...numeric.parse('50'),
      temperature: 'difference',
    });
    const legacy = { ...numeric.parse('273.15'), dimension: [0, 0, 0, 0, 1, 0, 0] as const };
    expect(quantityInUnit(legacy, '°F')).toMatchObject(numeric.parse('32'));
  });
  it('preserves physical kinds through scalars and rejects incompatible explicit kinds', () => {
    const energy = makeQuantity(numeric.parse('1'), 'J'),
      torque = makeQuantity(numeric.parse('1'), 'N*m', {}, 'torque');
    expect(() => math.add(energy, torque)).toThrow('#VALUE!');
    expect(() => makeQuantity(numeric.parse('1'), 'J', {}, 'torque')).toThrow('#VALUE!');
    expect(math.multiply(energy, numeric.parse('2')).quantityKind).toBe('energy');
    expect(math.multiply(numeric.parse('2'), energy).quantityKind).toBe('energy');
    expect(math.divide(energy, numeric.parse('2')).quantityKind).toBe('energy');
    expect(math.negate(energy).quantityKind).toBe('energy');
    expect(() => math.add(makeQuantity(numeric.parse('1'), 'rad'), numeric.parse('1'))).toThrow(
      '#VALUE!',
    );
    expect(() =>
      math.add(makeQuantity(numeric.parse('1'), 'b'), makeQuantity(numeric.parse('1'), 'rad')),
    ).toThrow('#VALUE!');
    const squared = math.multiply(
      makeQuantity(numeric.parse('2'), 'rad'),
      makeQuantity(numeric.parse('2'), 'rad'),
    );
    expect(squared.quantityKind).toBe('angle^2');
    expect(math.sqrt(squared).quantityKind).toBe('angle');
    expect(() => math.add(squared, numeric.parse('4'))).toThrow('#VALUE!');
    expect(() => validateUnitDimension('rad', 'dimensionless')).toThrow('#VALUE!');
    expect(() => validateUnitDimension('mm', 'Length')).not.toThrow();
    expect(() => validateUnitDimension('mm', 'length')).not.toThrow();
    expect(() => validateUnitDimension('m/s', 'm/s')).not.toThrow();
  });
  it('unions pinned provenance and rejects conflicting revisions before arithmetic', () => {
    const a = { ...numeric.parse('2'), provenance: { constants: '2022' } },
      b = { ...numeric.parse('3'), provenance: { catalog: 'v1' } };
    expect(math.multiply(a, b)).toMatchObject({
      provenance: { constants: '2022', catalog: 'v1' },
    });
    expect(numeric.add(a, b)).toMatchObject({ provenance: { constants: '2022', catalog: 'v1' } });
    expect(() => math.multiply(a, { ...b, provenance: { constants: '2018' } })).toThrow('#VALUE!');
    expect(evaluate('=1E-25')).toMatchObject({ numerator: '1', approximate: false });
  });
  it('validates all numeric metadata and bounded opaque JSON at ingress', () => {
    const value: NumericValue = { ...numeric.parse('2'), provenance: { constants: '2022' } };
    expect(() => assertNumericValue(value)).not.toThrow();
    for (const invalid of [
      { ...value, temperature: 'absolute' },
      { ...value, provenance: { constants: 2 } },
      { ...value, dimension: [0, 0, 0, 0, 1, 0] },
      { ...value, approximate: true },
    ])
      expect(() => assertNumericValue(invalid)).toThrow();
    expect(() => assertJsonValue({ x: new Date() })).toThrow();
    expect(() => assertJsonValue({ n: Infinity })).toThrow();
    let deep: unknown = {};
    for (let i = 0; i < 33; i++) deep = { deep };
    expect(() => assertJsonValue(deep)).toThrow('depth');
    const workbook = createWorkbook([['4'], ['=A1*2']]);
    workbook.worksheets[0].inputValues = {
      A1: {
        ...numeric.parse('0.04'),
        dimension: [1, 0, 0, 0, 0, 0, 0],
        provenance: { constants: '2022' },
      },
    };
    workbook.worksheets[0].formats = { A1: { kind: 'unit', unit: 'mm', inputUnit: 'mm' } };
    assertWorkbook(workbook);
    const result = calculateWorkbook(workbook);
    expect(result.worksheets[0].values[1][0]).toMatchObject({
      ...numeric.parse('0.08'),
      provenance: { constants: '2022' },
    });
    assertCalculationResult(result, workbook);
    expect(() =>
      assertCalculationResult(
        {
          ...result,
          context: {
            operands: {
              bad: { formula: '=A', name: 'A', value: { ...value, provenance: { x: 1 } } },
            },
          },
        },
        workbook,
      ),
    ).toThrow('provenance');
  });
  it('records only actually-read cell and extent dependencies under lazy evaluation', () => {
    const workbook = createWorkbook([
      ['=IF(TRUE,B1,C1)', '=SUM(A2:A3)', '=SUM(A:A)'],
      ['1'],
      ['2'],
    ]);
    const result = calculateWorkbook(workbook);
    expect(result.dependencies?.cells['["sheet-1",0,0]']).toEqual(['["sheet-1",1,0]']);
    expect(result.dependencies?.extents?.['["sheet-1",2,0]']).toEqual(['sheet-1']);
    expect(result.dependencies?.cells['["sheet-1",1,0]']).toEqual([
      '["sheet-1",0,1]',
      '["sheet-1",0,2]',
    ]);
  });
  it('round-trips canonical metadata, word context, pins and dependencies over HTTP', async () => {
    const workbook = createWorkbook([['1', '=A1*2', '=1 w']]);
    workbook.worksheets[0].inputValues = {
      A1: {
        ...numeric.parse('273.15'),
        dimension: [0, 0, 0, 0, 1, 0, 0],
        temperature: 'absolute',
        provenance: { constant: 'r1' },
      },
    };
    const provider = createHttpCalculator('http://localhost/calculate', {
      fetch: (async (url, init) =>
        handleCalculationRequest(new Request(url, init), {
          host: { hostContext: { revision: 'v1' } },
        })) as typeof fetch,
    });
    const result = await provider.calculate(workbook, { wordSizeBits: '32' });
    expect(result.worksheets[0].values[0][0]).toMatchObject({
      temperature: 'absolute',
      provenance: { constant: 'r1' },
    });
    expect(result.worksheets[0].values[0][2]).toMatchObject({
      numerator: '32',
      quantityKind: 'information',
    });
    expect(result.context?.hostContext).toEqual({ revision: 'v1' });
  });
});
describe('SE-CALC-6.1: exact localized formatting', () => {
  it('formats USD/EUR, large integers, percentages and leap day without Number', () => {
    expect(
      formatWorksheetCell(numeric.parse('9007199254740993'), {
        kind: 'currency',
        currency: 'USD',
        locale: 'en-US',
        decimals: 0,
      }),
    ).toBe('$9,007,199,254,740,993');
    expect(
      formatWorksheetCell(numeric.parse('-1234.567'), {
        kind: 'currency',
        currency: 'EUR',
        locale: 'de-DE',
        decimals: 2,
      }),
    ).toBe('-1.234,57 €');
    expect(
      formatWorksheetCell(numeric.parse('.125'), { kind: 'percent', locale: 'de-DE', decimals: 1 }),
    ).toBe('12,5 %');
    expect(formatWorksheetCell(numeric.parse('60'), { kind: 'date', locale: 'de-DE' })).toBe(
      '29. Feb. 1900',
    );
    expect(formatWorksheetCell(numeric.parse('60'), { kind: 'date' })).toBe('Feb 29, 1900');
    const workbook = createWorkbook([['1']]);
    for (const format of [
      { kind: 'number', locale: 'not-a-real-locale' },
      { kind: 'currency', currency: 'ZZZ' },
    ]) {
      workbook.worksheets[0].formats = { A1: format as never };
      expect(() => assertWorkbook(workbook)).toThrow();
    }
  });
  it('provides an exact Intl-parts fallback with the same currency and percent output', () => {
    for (const [text, locale, options] of [
      [
        '9007199254740993',
        'en-US',
        { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 },
      ],
      [
        '-1234.57',
        'de-DE',
        { style: 'currency', currency: 'EUR', minimumFractionDigits: 2, maximumFractionDigits: 2 },
      ],
      ['.125', 'de-DE', { style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 }],
    ] as const)
      expect(formatExactDecimal(text, locale, options, true)).toBe(
        formatExactDecimal(text, locale, options),
      );
  });
});
