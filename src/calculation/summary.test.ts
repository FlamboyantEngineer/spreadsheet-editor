import { describe, expect, it } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { formatCellValue } from './format.js';
import { createArithmetic } from './rational.js';
import { summarizeCellValues } from './summary.js';
import { createQuantityArithmetic, makeQuantity } from './units.js';

const arithmetic = createArithmetic(),
  quantities = createQuantityArithmetic();

describe('SE-CALC-4.1: exact and dimensionally valid selection statistics', () => {
  it('counts nonempty values and keeps rational statistics exact', () => {
    const third = arithmetic.divide(arithmetic.parse('1'), arithmetic.parse('3'));
    const summary = summarizeCellValues([third, third, third, null, '', 'label', false]);
    expect(summary.count).toBe(5);
    expect(summary.sum).toMatchObject(arithmetic.parse('1'));
    expect(summary.average).toEqual(third);
    expect(summarizeCellValues([null, '', 'label'])).toEqual({
      count: 1,
      sum: arithmetic.parse('0'),
      average: null,
    });
  });

  it('preserves compatible units and declines incompatible aggregate values', () => {
    const meter = makeQuantity(arithmetic.parse('1'), 'm');
    const centimeter = makeQuantity(arithmetic.parse('100'), 'cm');
    const summary = summarizeCellValues([meter, centimeter]);
    expect(formatCellValue(requireDefined(summary.sum))).toBe('2 m');
    expect(formatCellValue(requireDefined(summary.average))).toBe('1 m');
    expect(summarizeCellValues([meter, makeQuantity(arithmetic.parse('1'), 's')])).toEqual({
      count: 2,
      sum: null,
      average: null,
    });
    expect(summarizeCellValues([arithmetic.parse('0'), meter]).average).toBeNull();
    expect(
      summarizeCellValues([
        makeQuantity(arithmetic.parse('1'), 'J'),
        makeQuantity(arithmetic.parse('1'), 'N*m', {}, 'torque'),
      ]),
    ).toMatchObject({ sum: null, average: null });
  });

  it('averages absolute temperatures independently of their unavailable sum', () => {
    const freezing = makeQuantity(arithmetic.parse('0'), '°C'),
      boiling = makeQuantity(arithmetic.parse('100'), '°C');
    const summary = summarizeCellValues([freezing, boiling]);
    expect(summary.sum).toBeNull();
    expect(summary.average).toMatchObject({
      ...arithmetic.parse('323.15'),
      temperature: 'absolute',
    });
    const difference = quantities.subtract(boiling, freezing);
    expect(summarizeCellValues([difference, difference])).toMatchObject({
      sum: { numerator: '200', temperature: 'difference' },
      average: { numerator: '100', temperature: 'difference' },
    });
    expect(summarizeCellValues([freezing, difference]).average).toBeNull();
  });

  it('unions provenance and refuses incompatible revisions', () => {
    const first = { ...arithmetic.parse('1'), provenance: { constants: 'v1' } };
    expect(summarizeCellValues([first, arithmetic.parse('2')])).toMatchObject({
      sum: { provenance: { constants: 'v1' } },
      average: { provenance: { constants: 'v1' } },
    });
    expect(
      summarizeCellValues([first, { ...arithmetic.parse('2'), provenance: { constants: 'v2' } }]),
    ).toMatchObject({ count: 2, sum: null, average: null });
  });
});
