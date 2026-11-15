import { describe, expect, it } from 'vitest';
import { formatCells } from '../model/editing.js';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import { formatWorksheetCell } from './cell-format.js';
import { calculateWorkbook } from './evaluator.js';
import { formatCellValue } from './format.js';

describe('SE-CALC-4.1: exact engineering quantities', () => {
  it('converts tiny authored inputs to SI exactly and preserves their input unit when display changes', () => {
    const workbook = createWorkbook([['1E-25', '=A1*2', '=A1+A2'], ['3 mm']]);
    workbook.worksheets[0].formats = {
      A1: { kind: 'unit', unit: 'mm', inputUnit: 'mm' },
      A2: { kind: 'unit', unit: 'mm' },
    };
    const result = calculateWorkbook(workbook);
    expect(result.worksheets[0].values[0][0]).toMatchObject({
      numerator: '1',
      denominator: '10000000000000000000000000000',
      dimension: [1, 0, 0, 0, 0, 0, 0],
    });
    expect(
      formatWorksheetCell(result.worksheets[0].values[0][1], { kind: 'unit', unit: 'mm' }),
    ).toBe('2e-25 mm');
    const changed = formatCells(
      workbook,
      'sheet-1',
      { x1: 0, x2: 0, y1: 0, y2: 0 },
      {},
      { kind: 'unit', unit: 'm' },
    );
    expect(requireDefined(changed.worksheets[0].formats).A1.inputUnit).toBe('mm');
    expect(calculateWorkbook(changed).worksheets[0].values[0][0]).toEqual(
      result.worksheets[0].values[0][0],
    );
    expect(
      formatWorksheetCell(result.worksheets[0].values[1][0], { kind: 'unit', unit: 'm' }),
    ).toBe('0.003 m');
  });
  it('checks dimensions, combines derived units, and leaves unformatted digit-led labels as text', () => {
    const result = calculateWorkbook(
      createWorkbook([
        ['=10 mm + 2 cm', '=1 kg * 3 m/s^2', '=1 m + 2 s', '10 mm', '=ROUND(1 m,2)'],
      ]),
    );
    const values = result.worksheets[0].values[0];
    expect(formatWorksheetCell(values[0], { kind: 'unit', unit: 'mm' })).toBe('30 mm');
    expect(formatWorksheetCell(values[1], { kind: 'unit', unit: 'N' })).toBe('3 N');
    expect(values[2]).toMatchObject({ code: '#VALUE!' });
    expect(values[3]).toBe('10 mm');
    expect(values[4]).toMatchObject({ code: '#VALUE!' });
  });
});

describe('SE-CALC-4.1: calendar serial formats', () => {
  it('renders calendar formats in UTC and preserves Excel day 60', () => {
    const result = calculateWorkbook(
      createWorkbook([['46290', '60', '=DATE(2026,9,25)', '=TIME(12,30,0)']]),
    );
    const values = result.worksheets[0].values[0];
    expect(formatWorksheetCell(values[0], { kind: 'date' })).toBe('Sep 25, 2026');
    expect(formatWorksheetCell(values[1], { kind: 'date' })).toBe('Feb 29, 1900');
    expect(formatCellValue(values[2])).toBe('46290');
    expect(formatWorksheetCell(values[3], { kind: 'time' })).toBe('12:30:00');
    expect(result.worksheets[0].values[0][0]).toMatchObject({ numerator: '46290' });
  });
});

it('SE-CALC-5.1: captures a viewer-zone calendar clock once per pass', () => {
  let calls = 0;
  const workbook = createWorkbook([['=TODAY()', '=NOW()', '=TODAY()']]);
  const host = {
    now: () => {
      calls++;
      return new Date('2026-09-25T00:30:00Z');
    },
  };
  const result = calculateWorkbook(workbook, { timeZone: 'America/Los_Angeles' }, host);
  expect(calls).toBe(1);
  expect(formatWorksheetCell(result.worksheets[0].values[0][0], { kind: 'date' })).toBe(
    'Sep 24, 2026',
  );
  expect(formatWorksheetCell(result.worksheets[0].values[0][1], { kind: 'datetime' })).toBe(
    'Sep 24, 2026 17:30:00',
  );
});
it('SE-CALC-4.1: negative affine literals apply their sign before SI conversion and invalid units fail predictably', () => {
  const result = calculateWorkbook(createWorkbook([['=-1 °C', '=1 constructor']]));
  expect(formatWorksheetCell(result.worksheets[0].values[0][0], { kind: 'unit', unit: '°C' })).toBe(
    '-1 °C',
  );
  expect(result.worksheets[0].values[0][1]).toMatchObject({ code: '#VALUE!' });
});
