import { describe, expect, it } from 'vitest';
import { parseFormula } from './formula-parser.js';
import { parseFormulaReferenceHighlights } from './formula-reference-parser.js';
import { rewriteStructureReferences, translateFormulaReferences } from './formula-references.js';
import {
  formatSpreadsheetColumnName,
  parseSpreadsheetColumnName,
} from './spreadsheet-addressing.js';
import { renameWorksheet } from './structure.js';
import { createWorkbook } from './workbook.js';

describe('SE-CHROME-2.1, SE-CALC-2.1: formula reference tokens', () => {
  it('preserves sheet qualification, whitespace, absolute axes and host bindings when moving references', () => {
    const formula = "='Other sheet'! $A$1 : A3 + {{A1}} + A2";
    const insert = { axis: 'row', kind: 'insert', index: 1, count: 1 } as const;
    expect(
      rewriteStructureReferences(formula, insert, { currentSheet: 'Here', targetSheet: 'Here' }),
    ).toBe("='Other sheet'! $A$1 : A3 + {{A1}} + A3");
    expect(
      rewriteStructureReferences(formula, insert, {
        currentSheet: 'Here',
        targetSheet: 'Other sheet',
      }),
    ).toBe("='Other sheet'! $A$1 : A4 + {{A1}} + A2");
    expect(translateFormulaReferences(formula, 2, 3)).toBe(
      "='Other sheet'! $A$1 : C6 + {{A1}} + C5",
    );
  });

  it('renames a qualified reference without changing the spacing after its qualifier', () => {
    const workbook = createWorkbook([["='Other sheet'! \t$A$1 : A3"]]);
    workbook.worksheets.push({ id: 'other', name: 'Other sheet', data: [['1']] });
    expect(renameWorksheet(workbook, 'other', 'Data').worksheets[0].data[0][0]).toBe(
      "='Data'! \t$A$1 : A3",
    );
  });

  it('finds range endpoints consistently in drafts and calculation', () => {
    const text = 'SUM(B2 : $A$1, "C3", {{D4}}, LOG10 (10), Other! E5, A7';
    expect(
      parseFormulaReferenceHighlights('=' + text).map((reference) => reference.originalText),
    ).toEqual(['B2 : $A$1', 'A7']);
    expect(parseFormula(text + ')')).toMatchObject({
      kind: 'call',
      args: [
        { kind: 'reference', start: { x: 1, y: 1 }, end: { x: 0, y: 0 } },
        { kind: 'literal', value: 'C3' },
        { kind: 'name', name: 'D4' },
        { kind: 'call', name: 'LOG10' },
        { kind: 'reference', sheetName: 'Other' },
        { kind: 'reference' },
      ],
    });
  });

  it('rejects addresses beyond the grid without formatting overflowing column indices', () => {
    const longColumn = 'A'.repeat(300);
    expect(parseSpreadsheetColumnName(longColumn)).toBeNull();
    expect(formatSpreadsheetColumnName(Infinity)).toBe('A');
    expect(parseFormula(longColumn + '1')).toEqual({ kind: 'error', code: '#REF!' });
    expect(
      parseFormulaReferenceHighlights('=' + longColumn + '1 + A1').map(
        (reference) => reference.originalText,
      ),
    ).toEqual(['A1']);
    expect(translateFormulaReferences('=XFD1', 1, 0)).toBe('=#REF!');
  });
});
