import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import { calculateWorkbook } from './evaluator.js';
import { formatCellValue } from './format.js';
import { createArithmetic } from './rational.js';
import { isNumeric, type NumericValue } from './values.js';

const calculate = (formula: string) =>
  calculateWorkbook(createWorkbook([[formula]])).worksheets[0].values[0][0];
const exact = (numerator: string, denominator = '1') => ({
  kind: 'number',
  numerator,
  denominator,
  approximate: false,
});

describe('SE-CALC-1: exact arithmetic', () => {
  it('reads a formula whose value is empty as empty, not as an error', () => {
    const [values] = calculateWorkbook(
      createWorkbook([['=B1', '=C1', null, '=B1&"x"']]),
    ).worksheets;
    expect(values.values[0][0]).not.toBe('#ERROR!');
    expect(values.values[0][3]).toBe('x');
  });
  it('preserves 1E-25, near-one values, and their cancellation', () => {
    const workbook = createWorkbook([['1E-25', '=1-A1', '=1-B1', '=C1=A1', '=B1=1']]);
    const [small, nearOne, recovered, same, rounded] =
      calculateWorkbook(workbook).worksheets[0].values[0];
    expect(small).toEqual(exact('1', '10000000000000000000000000'));
    expect(nearOne).toEqual(exact('9999999999999999999999999', '10000000000000000000000000'));
    expect(recovered).toEqual(small);
    expect(same).toBe(true);
    expect(rounded).toBe(false);
    expect(formatCellValue(nearOne)).toBe('0.9999999999999999999999999');
    expect(formatCellValue(small)).toBe('1e-25');
  });
  it('cancels far beyond the decimal approximation precision', () => {
    expect(calculate('=(1-(1-1E-250))/1E-250')).toEqual(exact('1'));
    expect(calculate('=1/3*3')).toEqual(exact('1'));
    expect(calculate('=0.1+0.2')).toEqual(exact('3', '10'));
    expect(calculate('=9007199254740993+1')).toEqual(exact('9007199254740994'));
  });
  it('rounds only when explicitly requested, or when formatting a copy', () => {
    const third = calculate('=1/3');
    expect(third).toEqual(exact('1', '3'));
    expect(formatCellValue(third, { significantDigits: 3 })).toBe('0.333');
    expect(third).toEqual(exact('1', '3'));
    expect(calculate('=ROUND(1/3,2)')).toEqual(exact('33', '100'));
    expect(calculate('=ROUND(-2.5,0)')).toEqual(exact('-3'));
    expect(calculate('=INT(-1/3)')).toEqual(exact('-1'));
    expect(calculate('=MOD(-5,3)')).toEqual(exact('1'));
  });
  it('keeps integer powers and perfect square roots exact', () => {
    expect(calculate('=2^-3')).toEqual(exact('1', '8'));
    expect(calculate('=SQRT(4/9)')).toEqual(exact('2', '3'));
    expect(calculate('=25%')).toEqual(exact('1', '4'));
  });
  it('marks approximations and propagates the mark to dependents', () => {
    const values = calculateWorkbook(createWorkbook([['=SQRT(2)', '=A1*A1']]), { precision: 64 })
      .worksheets[0].values[0] as NumericValue[];
    expect(values.every((value) => value.approximate && value.precision === 64)).toBe(true);
    expect(formatCellValue(values[0], { significantDigits: 16 })).toBe('≈1.414213562373095');
    expect(formatCellValue(calculate('=SIN(1)'))).toMatch(/^≈0\.841470/);
    expect(calculate('=SQRT(-1)')).toEqual({ kind: 'error', code: '#NUM!' });
  });
  it('isolates arithmetic precision between consumers', () => {
    const a = createArithmetic({ precision: 32 }),
      b = createArithmetic({ precision: 100 });
    expect(a.sqrt(a.parse('2')).numerator.length).toBeLessThan(
      b.sqrt(b.parse('2')).numerator.length,
    );
    expect(a.sqrt(a.parse('2')).precision).toBe(32);
  });
});

describe('SE-CALC-2: text order does not follow the runtime locale', () => {
  afterEach(() => vi.restoreAllMocks());
  it('compares text with one fixed collation', () => {
    const localeCompare = vi.spyOn(String.prototype, 'localeCompare');
    expect(calculate('=IF("Äpfel"<"Birnen",1,0)')).toEqual(exact('1'));
    expect(calculate('=IF("a"="A",1,0)')).toEqual(exact('1'));
    expect(localeCompare).not.toHaveBeenCalled();
  });
});

describe('SE-CALC-2: workbook formulas', () => {
  it('reads other sheets, named values, ranges, and absolute references', () => {
    const workbook: Workbook = {
      schemaVersion: 1,
      worksheets: [
        { id: 'first', name: 'First sheet', data: [['1E-25'], ['2E-25']] },
        {
          id: 'second',
          name: 'Second',
          data: [["=SUM('First sheet'!$A$1:A2)/3", "=SUM('First sheet'!A:A)", '=rate*2']],
        },
      ],
    };
    const values = calculateWorkbook(
      workbook,
      {},
      { readName: (name) => (name === 'rate' ? createArithmetic().parse('1') : undefined) },
    ).worksheets[1].values[0];
    expect(values[0]).toEqual(exact('1', '10000000000000000000000000'));
    expect(values[1]).toEqual(exact('3', '10000000000000000000000000'));
    expect(values[2]).toEqual(exact('2'));
  });
  it('evaluates IF lazily and exposes typed formula errors', () => {
    expect(calculate('=IF(FALSE,1/0,7)')).toEqual(exact('7'));
    expect(calculate('=IFERROR(1/0,9)')).toEqual(exact('9'));
    expect(calculate('=ISERROR(1/0)')).toBe(true);
    expect(calculate('=NoSuchSheet!A1')).toEqual({ kind: 'error', code: '#REF!' });
    expect(calculate('=MISSING(2)')).toEqual({ kind: 'error', code: '#NAME?' });
    expect(calculate('=2+')).toEqual({ kind: 'error', code: '#ERROR!' });
    expect(calculate('=1/0')).toEqual({ kind: 'error', code: '#DIV/0!' });
  });
  it('detects cross-sheet cycles and bounds expensive ranges', () => {
    const workbook: Workbook = {
      schemaVersion: 1,
      worksheets: [
        { id: 'a', name: 'A', data: [['=B!A1']] },
        { id: 'b', name: 'B', data: [['=A!A1']] },
      ],
    };
    expect(calculateWorkbook(workbook).worksheets[0].values[0][0]).toEqual({
      kind: 'error',
      code: '#CYCLE!',
    });
    expect(calculate('=SUM(A1:XFD1048576)')).toEqual({ kind: 'error', code: '#LIMIT!' });
    expect(calculate('=1e99999999')).toEqual({ kind: 'error', code: '#LIMIT!' });
  });
  it('supports aggregate, statistical, text, and host-defined functions', () => {
    expect(calculate('=AVERAGE(1,2,3)')).toEqual(exact('2'));
    expect(calculate('=VAR.S(1,2,3)')).toEqual(exact('1'));
    expect(calculate('=STDEV.P(1,1)')).toEqual(exact('0'));
    expect(calculate('=UPPER(LEFT("example",3))&"!"')).toBe('EXA!');
    expect(
      calculateWorkbook(
        createWorkbook([['=DOUBLE(3)']]),
        {},
        {
          callFunction: (name, args) =>
            name === 'DOUBLE' && isNumeric(args[0])
              ? createArithmetic().multiply(args[0], createArithmetic().parse('2'))
              : undefined,
        },
      ).worksheets[0].values[0][0],
    ).toEqual(exact('6'));
  });
});

describe('SE-4: evaluation cost does not depend on reading order', () => {
  it('sums a range placed above its data about as fast as below it', () => {
    const values = Array.from({ length: 9_999 }, (_, y) => [String(y + 1)]);
    const above = createWorkbook([['=SUM(A2:A10000)'], ...values]),
      below = createWorkbook([...values, ['=SUM(A1:A9999)']]);
    const time = (workbook: Workbook) => {
      const started = performance.now(),
        result = calculateWorkbook(workbook);
      return { elapsed: performance.now() - started, result };
    };
    time(below);
    const first = time(below),
      second = time(above);
    expect(second.result.worksheets[0].values[0][0]).toEqual(
      first.result.worksheets[0].values[9_999][0],
    );
    // A forward read once threw an Error per cell: about seven times slower than reading back.
    expect(second.elapsed).toBeLessThan(Math.max(3 * first.elapsed, 50));
    expect(second.elapsed).toBeLessThan(1_000);
  });
});
