import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import {
  type ExternalCalculationSnapshot,
  workbookCalculationKey,
} from '../calculation/publication.js';
import { createWorkbook } from '../model/workbook.js';
import { createLinkedSpreadsheetInsertion, updateLinkedSpreadsheetInsertion } from './linked.js';

describe('SE-DOC-3.1: linked authoritative captures', () => {
  it('captures detached exact results and refreshes through source/revision checks', () => {
    const workbook = createWorkbook([['=1/3']]);
    const publication: ExternalCalculationSnapshot = {
      inputKey: workbookCalculationKey(workbook),
      revision: 1,
      status: 'ready',
      result: calculateWorkbook(workbook),
    };
    const captured = createLinkedSpreadsheetInsertion(workbook, publication, {
      source: { id: 'source', revision: 'r1' },
      range: 'A1',
    });
    const printed = captured.table.html;
    workbook.worksheets[0].data[0][0] = '=2/3';
    expect(captured.table.html).toBe(printed);
    expect(captured.workbook.worksheets[0].data[0][0]).toBe('=1/3');
    const next: ExternalCalculationSnapshot = {
      inputKey: workbookCalculationKey(workbook),
      revision: 2,
      status: 'ready',
      result: calculateWorkbook(workbook),
    };
    expect(() =>
      updateLinkedSpreadsheetInsertion(captured, {
        workbook,
        publication: next,
        source: { id: 'source', revision: 'r2' },
        expectedRevision: 'stale',
      }),
    ).toThrow('captured revision');
    expect(() =>
      updateLinkedSpreadsheetInsertion(captured, {
        workbook,
        publication: next,
        source: { id: 'source', revision: 'r1' },
        expectedRevision: 'r1',
      }),
    ).toThrow('different captured data');
    const refreshed = updateLinkedSpreadsheetInsertion(captured, {
      workbook,
      publication: next,
      source: { id: 'source', revision: 'r2' },
      expectedRevision: 'r1',
    });
    expect(refreshed.source.revision).toBe('r2');
    expect(refreshed.table.html).not.toBe(printed);
    expect(captured.source.revision).toBe('r1');
    expect(() =>
      createLinkedSpreadsheetInsertion(workbook, publication, {
        source: { id: 'source', revision: 'r1' },
      }),
    ).toThrow();
    expect(() =>
      createLinkedSpreadsheetInsertion(
        workbook,
        { inputKey: next.inputKey, revision: 2, status: 'pending' },
        { source: { id: 'source', revision: 'r2' } },
      ),
    ).toThrow('ready authoritative');
  });
});
