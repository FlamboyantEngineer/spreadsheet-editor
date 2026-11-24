import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { createWorkbook } from '../model/workbook.js';
import { createDocumentInsertion } from './insertion.js';

describe('SE-DOC-1.1: document insertion modes', () => {
  it('creates an ordinary table from calculated text, with no live formulas', () => {
    const workbook = createWorkbook([['=1-1e-25', '<img src=x onerror=alert(1)>']]);
    const insertion = createDocumentInsertion(workbook, calculateWorkbook(workbook), {
      mode: 'table',
    });
    expect(insertion.kind).toBe('table');
    if (insertion.kind !== 'table') throw new Error('Expected a table.');
    expect(insertion.rows[0][0].text).toBe('0.9999999999999999999999999');
    expect(Object.keys(insertion).sort()).toEqual([
      'html',
      'kind',
      'range',
      'rows',
      'schemaVersion',
      'sheetId',
    ]);
    expect(insertion.html).toContain('&lt;img');
    expect(insertion.html).not.toContain('<img');
    expect(JSON.stringify(insertion.rows)).not.toContain('=1-');
  });
  it('keeps cross-sheet dependencies and isolates the live insertion snapshot', () => {
    const workbook = createWorkbook([['=Inputs!A1/3']]);
    workbook.worksheets.push({ id: 'inputs', name: 'Inputs', data: [['1']] });
    const insertion = createDocumentInsertion(workbook, null, { mode: 'spreadsheet', range: 'A1' });
    if (insertion.kind !== 'spreadsheet') throw new Error('Expected live insertion.');
    workbook.worksheets[1].data[0][0] = '9';
    expect(insertion.workbook.worksheets).toHaveLength(2);
    expect(calculateWorkbook(insertion.workbook).worksheets[0].values[0][0]).toMatchObject({
      numerator: '1',
      denominator: '3',
    });
    expect(insertion.range).toBe('A1');
  });
  it('preserves merge spans and selected bounds in ordinary tables', () => {
    const workbook = createWorkbook([
      ['Merged', '', 'Outside'],
      ['Next', 'Cell'],
    ]);
    workbook.worksheets[0].mergeCells = { A1: [2, 1] };
    const insertion = createDocumentInsertion(workbook, calculateWorkbook(workbook), {
      mode: 'table',
      range: 'A1:B2',
    });
    if (insertion.kind !== 'table') throw new Error('Expected table.');
    expect(insertion.rows[0]).toEqual([{ text: 'Merged', colspan: 2, rowspan: 1 }]);
    expect(insertion.rows[1]).toHaveLength(2);
    expect(insertion.html).not.toContain('Outside');
  });
  it('uses the original anchor value and format when a selection clips a merge', () => {
    const workbook = createWorkbook([['=1/8'], ['', 'Covered'], ['']]);
    workbook.worksheets[0].mergeCells = { A1: [3, 3] };
    workbook.worksheets[0].formats = { A1: { kind: 'percent', decimals: 2 } };
    const insertion = createDocumentInsertion(workbook, calculateWorkbook(workbook), {
      mode: 'table',
      range: 'C3:B2',
    });
    if (insertion.kind !== 'table') throw new Error('Expected table.');
    expect(insertion.range).toBe('B2:C3');
    expect(insertion.rows).toEqual([[{ text: '12.50%', colspan: 2, rowspan: 2 }], []]);
    expect(insertion.html).toContain('<td colspan="2" rowspan="2">12.50%</td>');
  });
  it('captures an explicit cell near the end of a tall valid worksheet', () => {
    const workbook = createWorkbook(Array.from({ length: 140_000 }, () => []));
    workbook.worksheets[0].data[139_999] = ['Last'];
    const range = 'A140000';
    const table = createDocumentInsertion(workbook, calculateWorkbook(workbook), {
      mode: 'table',
      range,
    });
    const live = createDocumentInsertion(workbook, null, { mode: 'spreadsheet', range });
    expect(table).toMatchObject({ range, rows: [[{ text: 'Last' }]] });
    expect(live).toMatchObject({ range });
    expect(() => createDocumentInsertion(workbook, null, { mode: 'spreadsheet' })).toThrow(
      'Invalid insertion range.',
    );
  });
  it('rejects a normal table without computed values', () => {
    expect(() => createDocumentInsertion(createWorkbook(), null, { mode: 'table' })).toThrow(
      'current calculation',
    );
  });
  it('rejects results with mismatched dimensions', () => {
    const result = calculateWorkbook(createWorkbook([['=1']]));
    expect(() =>
      createDocumentInsertion(createWorkbook([['=1', '=2']]), result, { mode: 'table' }),
    ).toThrow('Result does not match the row.');
  });
});
