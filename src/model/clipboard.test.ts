import { describe, expect, it } from 'vitest';
import { createWorkbook } from './workbook.js';
import { translateFormulaReferences } from './structure.js';
import { createSpreadsheetCopy, pasteSpreadsheetCopy } from './clipboard.js';

describe('SE-CR-2.3: authored clipboard references', () => {
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
    const copy = createSpreadsheetCopy(workbook, 'sheet-1', selection, '2');
    expect(pasteSpreadsheetCopy(workbook, 'sheet-1', 2, 1, copy).worksheets[0].data[1][2]).toBe(
      '=B2*2',
    );
    const cut = createSpreadsheetCopy(workbook, 'sheet-1', selection, '2', true);
    expect(pasteSpreadsheetCopy(workbook, 'sheet-1', 2, 1, cut).worksheets[0].data[1][2]).toBe(
      '=A1*2',
    );
    expect(workbook.worksheets[0].data).toEqual([['1', '=A1*2']]);
  });
});
