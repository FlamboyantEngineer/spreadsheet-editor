import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { parseDelimited, serializeDelimited } from './delimited.js';
import { writeCells } from './editing.js';
import { rewriteStructureReferences } from './formula-references.js';
import { addComment, proposeWorkbookChange } from './review.js';
import {
  changeWorkbookStructure,
  deleteWorksheet,
  renameWorksheet,
  validateWorksheetName,
} from './structure.js';
import { createWorkbook } from './workbook.js';

describe('SE-CHROME-2.1: structure and exact formula text', () => {
  it('rewrites relative and absolute references and preserves strings and scientific literals', () => {
    const formula = '=SUM($A$1:A3)+Other!B2+\'Sheet one\'!B2+1E-25+IF(TRUE,"A2",LOG10(100))';
    expect(
      rewriteStructureReferences(
        formula,
        {
          axis: 'row',
          kind: 'insert',
          index: 1,
          count: 1,
        },
        { currentSheet: 'Sheet one', targetSheet: 'Sheet one' },
      ),
    ).toBe('=SUM($A$1:A4)+Other!B2+\'Sheet one\'!B3+1E-25+IF(TRUE,"A2",LOG10(100))');
    expect(
      rewriteStructureReferences(
        '=SUM(A1:A3)+A2',
        {
          axis: 'row',
          kind: 'delete',
          index: 0,
          count: 1,
        },
        { currentSheet: 'Sheet1', targetSheet: 'Sheet1' },
      ),
    ).toBe('=SUM(A1:A2)+A1');
    expect(
      rewriteStructureReferences(
        '=A2',
        {
          axis: 'row',
          kind: 'delete',
          index: 1,
          count: 1,
        },
        { currentSheet: 'Sheet1', targetSheet: 'Sheet1' },
      ),
    ).toBe('=#REF!');
  });
  it('rewrites whole rows/columns and quoted sheet names with escaped apostrophes', () => {
    expect(
      rewriteStructureReferences(
        "=SUM('O''Brien'!$A:$C)+SUM(1:3)",
        {
          axis: 'column',
          kind: 'insert',
          index: 1,
          count: 2,
        },
        { currentSheet: 'Sheet1', targetSheet: "O'Brien" },
      ),
    ).toBe("=SUM('O''Brien'!$A:$E)+SUM(1:3)");
    expect(
      rewriteStructureReferences(
        '=SUM($1:$3)',
        {
          axis: 'row',
          kind: 'delete',
          index: 1,
          count: 1,
        },
        { currentSheet: 'Sheet1', targetSheet: 'Sheet1' },
      ),
    ).toBe('=SUM($1:$2)');
  });
  it('moves data, formulas on other sheets, styles, merge spans, and comment anchors together', () => {
    let source = createWorkbook([
      ['1E-25', 'Label'],
      ['=A1', ''],
    ]);
    source.worksheets[0].style = { A2: 'font-weight: bold' };
    source.worksheets[0].mergeCells = { A2: [2, 1] };
    source.worksheets.push({ id: 'other', name: 'Other', data: [['=Sheet1!A2']] });
    source = addComment(
      source,
      { sheetId: 'sheet-1', range: 'A2', text: 'Move this' },
      {
        id: 'comment',
        actor: { id: 'alex', name: 'Alex' },
        time: '2026-11-10T12:00:00Z',
      },
    );
    const next = changeWorkbookStructure(source, 'sheet-1', {
      axis: 'row',
      kind: 'insert',
      index: 1,
      count: 1,
    });
    expect(next.worksheets[0].data[2][0]).toBe('=A1');
    expect(next.worksheets[1].data[0][0]).toBe('=Sheet1!A3');
    expect(next.worksheets[0].style?.A3).toBe('font-weight: bold');
    expect(next.worksheets[0].mergeCells?.A3).toEqual([2, 1]);
    expect(next.comments?.[0].range).toBe('A3');
    expect(source.comments?.[0].range).toBe('A2');
  });
  it('marks deleted comment anchors and suggestions instead of silently attaching them elsewhere', () => {
    let source = createWorkbook([['1'], ['2'], ['3']]);
    source = addComment(
      source,
      { sheetId: 'sheet-1', range: 'A2', text: 'Deleted' },
      {
        id: 'comment',
        actor: { id: 'alex', name: 'Alex' },
        time: '2026-11-10T12:00:00Z',
      },
    );
    source = proposeWorkbookChange(
      source,
      {
        proposed: writeCells(source, [['4']], { sheetId: 'sheet-1', x: 0, y: 1 }),
        summary: 'Edit A2',
      },
      { id: 'suggestion', actor: { id: 'alex', name: 'Alex' }, time: '2026-11-10T12:00:00Z' },
    );
    const next = changeWorkbookStructure(source, 'sheet-1', {
      axis: 'row',
      kind: 'delete',
      index: 1,
      count: 1,
    });
    expect(next.comments?.[0].orphaned).toBe(true);
    expect(next.suggestions?.[0].outdated).toBe(true);
  });
  it('renames/deletes sheet-qualified references and validates sheet names', () => {
    const workbook = createWorkbook([['=Inputs!A1+1E-25', '="Inputs!A1"']]);
    workbook.worksheets.push({ id: 'inputs', name: 'Inputs', data: [['3']] });
    const renamed = renameWorksheet(workbook, 'inputs', "O'Brien");
    expect(renamed.worksheets[0].data[0]).toEqual(["='O''Brien'!A1+1E-25", '="Inputs!A1"']);
    expect(deleteWorksheet(renamed, 'inputs').worksheets[0].data[0][0]).toBe('=#REF!+1E-25');
    expect(() => renameWorksheet(workbook, 'inputs', 'Bad:Name')).toThrow('unique worksheet');
  });
  it('SE-CHROME-2.1: shares normalized worksheet validation with inline renaming', () => {
    const workbook = createWorkbook([["='  Input   Data  '!$A$1+1E-25"]]);
    workbook.worksheets.push({ id: 'inputs', name: '  Input   Data  ', data: [['3']] });
    const original = structuredClone(workbook);

    expect(validateWorksheetName(workbook, 'inputs', '  Input\n Data  ')).toBe('Input Data');
    expect(validateWorksheetName(workbook, 'inputs', ' x '.repeat(128))).toHaveLength(255);
    expect(validateWorksheetName(workbook, 'inputs', 'x'.repeat(256))).toHaveLength(256);
    const renamed = renameWorksheet(workbook, 'inputs', '  Year\n  One  ');
    expect(renamed.worksheets[1]).toMatchObject({ id: 'inputs', name: 'Year One' });
    expect(renamed.worksheets[0].data[0][0]).toBe("='Year One'!$A$1+1E-25");
    expect(workbook).toEqual(original);
  });
  it.each(['', ' \t\n ', 'sheet1', ' Sheet1 ', 'x'.repeat(257), ...'[]:*?/\\!'])(
    'SE-CHROME-2.1: refuses invalid inline worksheet name %j in validation and mutation',
    (name) => {
      const workbook = createWorkbook();
      workbook.worksheets.push({ id: 'inputs', name: 'Inputs', data: [['']] });
      expect(() => validateWorksheetName(workbook, 'inputs', name)).toThrow(TypeError);
      expect(() => renameWorksheet(workbook, 'inputs', name)).toThrow(TypeError);
      expect(workbook.worksheets[1].name).toBe('Inputs');
    },
  );
  it('SE-CHROME-2.1: rejects stale worksheet ids and case/whitespace-equivalent names', () => {
    const workbook = createWorkbook();
    workbook.worksheets.push({ id: 'inputs', name: ' Input   Data ', data: [['']] });
    expect(() => validateWorksheetName(workbook, 'sheet-1', '  input\n data ')).toThrow(TypeError);
    expect(() => validateWorksheetName(workbook, 'missing', 'New name')).toThrow(TypeError);
    expect(() => renameWorksheet(workbook, 'missing', 'New name')).toThrow(TypeError);
  });
  it('round-trips multiline CSV and clipboard text without converting numeric strings', () => {
    const rows = [['1E-25', '0.9999999999999999999999999', 'line\none', 'a,b', '"quoted"']];
    expect(parseDelimited(serializeDelimited(rows, ','), ',')).toEqual(rows);
    expect(parseDelimited('1E-25\t=1-A1\r\n')).toEqual([['1E-25', '=1-A1']]);
  });
  it('SE-CHROME-2.1: reads, rewrites, renames and deletes sheet names outside ASCII', () => {
    const workbook = createWorkbook([['=Données!A2*2', '=Δ_1!A1']]);
    workbook.worksheets.push(
      { id: 'data', name: 'Données', data: [['x'], ['21']] },
      { id: 'delta', name: 'Δ_1', data: [['5']] },
    );
    expect(calculateWorkbook(workbook).worksheets[0].values[0]).toEqual([
      expect.objectContaining({ numerator: '42', denominator: '1' }),
      expect.objectContaining({ numerator: '5', denominator: '1' }),
    ]);
    const insert = { axis: 'row', kind: 'insert', index: 0, count: 1 } as const;
    expect(
      rewriteStructureReferences('=Données!A2*2', insert, {
        currentSheet: 'Sheet1',
        targetSheet: 'Données',
      }),
    ).toBe('=Données!A3*2');
    expect(
      rewriteStructureReferences('=Données!A2*2', insert, {
        currentSheet: 'Sheet1',
        targetSheet: 'es',
      }),
    ).toBe('=Données!A2*2');
    expect(changeWorkbookStructure(workbook, 'data', insert).worksheets[0].data[0][0]).toBe(
      '=Données!A3*2',
    );
    expect(deleteWorksheet(workbook, 'data').worksheets[0].data[0][0]).toBe('=#REF!*2');
    expect(renameWorksheet(workbook, 'data', 'Résultats').worksheets[0].data[0][0]).toBe(
      "='Résultats'!A2*2",
    );
  });
  it('SE-CHROME-2.1: compares sheet names the same way when renaming, restructuring and deleting', () => {
    const workbook = createWorkbook([["='  Input   Data '!A1", '=input data!A1']]);
    workbook.worksheets.push({ id: 'inputs', name: 'Input Data', data: [['3']] });
    const insert = { axis: 'row', kind: 'insert', index: 0, count: 1 } as const;
    expect(changeWorkbookStructure(workbook, 'inputs', insert).worksheets[0].data[0]).toEqual([
      "='  Input   Data '!A2",
      '=input data!A2',
    ]);
    expect(deleteWorksheet(workbook, 'inputs').worksheets[0].data[0]).toEqual(['=#REF!', '=#REF!']);
  });
  it('SE-CHROME-2.1: renaming a sheet rewrites formulas held in pending suggestions', () => {
    const workbook = createWorkbook([['1']]);
    workbook.worksheets.push({ id: 'inputs', name: 'Inputs', data: [['3']] });
    const proposed = writeCells(workbook, [['=Inputs!A1*2']], { sheetId: 'sheet-1', x: 0, y: 0 });
    const suggested = proposeWorkbookChange(
      workbook,
      { proposed, summary: 'Use inputs' },
      {
        id: 'suggestion',
        actor: { id: 'alex', name: 'Alex' },
        time: '2026-11-10T12:00:00Z',
      },
    );
    const renamed = renameWorksheet(suggested, 'inputs', 'Sources');
    expect(renamed.suggestions?.[0].changes[0].after).toBe("='Sources'!A1*2");
  });
  it('SE-CHROME-2.1: scans long formulas in linear time', () => {
    const insert = { axis: 'row', kind: 'insert', index: 0, count: 1 } as const;
    const references =
      '=' + Array.from({ length: 20_000 }, (_, index) => `A${index + 1}`).join('+');
    // A long run of words and spaces once made every word start rescan the rest of the run.
    const words = '=' + 'ab '.repeat(40_000) + 'A1';
    const started = performance.now();
    const rewritten = rewriteStructureReferences(references, insert, {
      currentSheet: 'Sheet1',
      targetSheet: 'Sheet1',
    });
    expect(rewritten.startsWith('=A2+A3+')).toBe(true);
    expect(rewritten.endsWith('+A20001')).toBe(true);
    expect(
      rewriteStructureReferences(words, insert, { currentSheet: 'Sheet1', targetSheet: 'Sheet1' }),
    ).toBe('=' + 'ab '.repeat(40_000) + 'A2');
    // Generous: the quadratic scan took about 5 seconds here.
    expect(performance.now() - started).toBeLessThan(1_000);
  });
});
