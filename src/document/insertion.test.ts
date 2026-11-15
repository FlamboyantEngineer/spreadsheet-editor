import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { createWorkbook } from '../model/workbook.js';
import { createDocumentInsertion } from './insertion.js';

describe('SE-DOC-1: document insertion modes', () => {
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
  it('rejects a normal table without computed values', () => {
    expect(() => createDocumentInsertion(createWorkbook(), null, { mode: 'table' })).toThrow(
      'current calculation',
    );
  });
});
