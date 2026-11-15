import { describe, expect, it } from 'vitest';
import { createSpreadsheetCopy, pasteSpreadsheetCopy } from './clipboard.js';
import { translateFormulaReferences } from './structure.js';
import { createWorkbook } from './workbook.js';

describe('SE-HOST-MODEL-3.1: authored clipboard references', () => {
  it('shifts relative axes and leaves quoted text, absolute axes and scientific numbers intact', () => {
    expect(translateFormulaReferences('=SUM(A1:B3)+$C4+D$5+$E$6+1E-25+"A1"', 2, 3)).toBe(
      '=SUM(C4:D6)+$C7+F$5+$E$6+1E-25+"A1"',
    );
    expect(translateFormulaReferences("='Other sheet'!A1+B:B+2:4", 1, 2)).toBe(
      "='Other sheet'!B3+C:C+4:6",
    );
    expect(translateFormulaReferences('=A1+$B$2', -1, -1)).toBe('=#REF!+$B$2');
  });
  it('copies formulas relative to their original target and cuts without translation', () => {
    const workbook = createWorkbook([['1', '=A1*2']]);
    const selection = { x1: 1, x2: 1, y1: 0, y2: 0 };
    const copy = createSpreadsheetCopy(
      workbook,
      { sheetId: 'sheet-1', range: selection },
      { text: '2' },
    );
    expect(
      pasteSpreadsheetCopy(workbook, copy, { sheetId: 'sheet-1', x: 2, y: 1 }).worksheets[0]
        .data[1][2],
    ).toBe('=B2*2');
    const cut = createSpreadsheetCopy(
      workbook,
      { sheetId: 'sheet-1', range: selection },
      { text: '2', cut: true },
    );
    expect(
      pasteSpreadsheetCopy(workbook, cut, { sheetId: 'sheet-1', x: 2, y: 1 }).worksheets[0]
        .data[1][2],
    ).toBe('=A1*2');
    expect(workbook.worksheets[0].data).toEqual([['1', '=A1*2']]);
  });
  it('SE-HOST-MODEL-3.1: uses the host factory when pasted data extends tracked lines', () => {
    const workbook = createWorkbook([['1']]);
    workbook.worksheets[0].rowIds = ['row-1'];
    workbook.worksheets[0].columnIds = ['column-1'];
    const copy = createSpreadsheetCopy(
      workbook,
      { sheetId: 'sheet-1', range: { x1: 0, x2: 0, y1: 0, y2: 0 } },
      { text: '1' },
    );
    const next = pasteSpreadsheetCopy(workbook, copy, {
      sheetId: 'sheet-1',
      x: 1,
      y: 1,
      idFactory: (kind) => `host-${kind}`,
    });
    expect(next.worksheets[0].rowIds).toEqual(['row-1', 'host-row']);
    expect(next.worksheets[0].columnIds).toEqual(['column-1', 'host-column']);
  });
});
