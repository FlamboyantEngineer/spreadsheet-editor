import { describe, expect, it } from 'vitest';
import { commandOperations } from '../editor/command-operations.js';
import { applyWorkbookOperations, applyWorkbookTransaction } from './apply-operations.js';
import { materializeWorksheetExtent } from './extent.js';
import {
  changeWorkbookStructureWithOperations,
  invertWorkbookOperations,
  type WorkbookOperation,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { addComment } from './review.js';
import { deleteWorksheet, renameWorksheet, type StructureChange } from './structure.js';
import { appendWorksheet, copyWorkbook, createWorkbook, type Workbook } from './workbook.js';

const stamp = { id: 'thread-1', time: '2026-01-01T00:00:00Z', actor: { id: 'ana', name: 'Ana' } };

/** A worksheet with every structure-sensitive part present, plus a second worksheet. */
function rich(): Workbook {
  let workbook = materializeWorksheetExtent(
    createWorkbook([
      ['1', '2', '3', '4'],
      ['5', '=A1+B2', 'text', '8'],
      ['9', '10', '=SUM(A1:A3)', '12'],
      ['13', '14', '15', '16'],
    ]),
    'sheet-1',
    [4, 4],
  );
  const sheet = workbook.worksheets[0];
  sheet.rowIds = ['row-1', 'row-2', 'row-3', 'row-4'];
  sheet.columnIds = ['column-1', 'column-2', 'column-3', 'column-4'];
  sheet.rows = { '0': { height: 31 }, '1': { height: 47 }, '3': { height: 19 } };
  sheet.columns = [{ width: 91 }, { width: 137 }, { width: 64 }, { width: 120 }];
  sheet.mergeCells = { C3: [2, 2] };
  sheet.formats = { B2: { kind: 'number', decimals: 2 }, D4: { kind: 'general' } };
  sheet.style = { A2: 'font-weight: bold', C1: 'color: red' };
  sheet.freezeColumns = 1;
  workbook = addComment(workbook, 'sheet-1', 'B3', 'Check this', stamp);
  workbook = appendWorksheet(workbook);
  workbook.worksheets[1].id = 'sheet-2';
  workbook.worksheets[1].data = [["='Sheet1'!A1", 'x']];
  return workbook;
}

function lines(workbook: Workbook, change: StructureChange) {
  const ids = Array.from({ length: change.count }, (_, index) => `new-${change.axis}-${index}`);
  let next = 0;
  return changeWorkbookStructureWithOperations(workbook, 'sheet-1', change, () => ids[next++]);
}
function sheetChange(before: Workbook, after: Workbook) {
  return { workbook: after, operations: commandOperations(before, after) };
}

const cases: [
  string,
  (workbook: Workbook) => { workbook: Workbook; operations: WorkbookOperation[] },
][] = [
  ['insert a row', (w) => lines(w, { axis: 'row', kind: 'insert', index: 1, count: 1 })],
  [
    'insert two rows inside a merge',
    (w) => lines(w, { axis: 'row', kind: 'insert', index: 3, count: 2 }),
  ],
  ['delete a row', (w) => lines(w, { axis: 'row', kind: 'delete', index: 1, count: 1 })],
  [
    'delete rows through a merge',
    (w) => lines(w, { axis: 'row', kind: 'delete', index: 2, count: 2 }),
  ],
  ['insert a column', (w) => lines(w, { axis: 'column', kind: 'insert', index: 0, count: 1 })],
  ['delete a column', (w) => lines(w, { axis: 'column', kind: 'delete', index: 1, count: 1 })],
  [
    'delete columns through a merge',
    (w) => lines(w, { axis: 'column', kind: 'delete', index: 2, count: 2 }),
  ],
  [
    'add a worksheet',
    (w) => {
      const next = appendWorksheet(w);
      requireDefined(next.worksheets.at(-1)).id = 'sheet-3';
      return sheetChange(w, next);
    },
  ],
  ['remove a worksheet', (w) => sheetChange(w, deleteWorksheet(w, 'sheet-2'))],
  ['rename a worksheet', (w) => sheetChange(w, renameWorksheet(w, 'sheet-1', 'Data'))],
  [
    'reorder worksheets',
    (w) => {
      const next = copyWorkbook(w);
      next.worksheets.reverse();
      return sheetChange(w, next);
    },
  ],
];

describe('SE-HOST-5.1: structure operations and their inverses', () => {
  it.each(cases)('%s, then its inverse, restores the workbook exactly', (_name, change) => {
    const initial = rich();
    const changed = change(copyWorkbook(initial));
    expect(changed.workbook).not.toStrictEqual(initial);
    const applied = applyWorkbookOperations(copyWorkbook(initial), changed.operations);
    expect(applied).toStrictEqual(changed.workbook);
    expect(
      applyWorkbookOperations(applied, invertWorkbookOperations(changed.operations)),
    ).toStrictEqual(initial);
  });

  // A rename's inverse renames back and restores the formulas the rename rewrote. A formula
  // written after the rename is not in its operations, so the transaction layer rejects the
  // stale inverse instead of leaving that formula pointing at a missing worksheet.
  const renamedThenReferenced = () => {
    const initial = rich();
    const renamed = sheetChange(initial, renameWorksheet(initial, 'sheet-1', 'Data'));
    const withFormula = copyWorkbook(renamed.workbook);
    withFormula.worksheets[1].data[0][1] = '=Data!B2';
    return { renamed, withFormula, inverse: invertWorkbookOperations(renamed.operations) };
  };
  it('rejects a rename inverse after a formula names the new worksheet', () => {
    const { renamed, withFormula, inverse } = renamedThenReferenced();
    expect(() =>
      applyWorkbookTransaction(withFormula, {
        transactionId: 'undo-rename',
        source: 'undo',
        baseline: renamed.workbook,
        operations: inverse,
      }),
    ).toThrow('referential-rename-baseline');
  });
  // Expected failure: plain operation application does not rewrite references the operations
  // do not name (specs/HOST-INTEGRATION.md, SE-HOST-5.1).
  it.fails(
    'rewrites a later formula when the rename inverse is applied without a transaction',
    () => {
      const { withFormula, inverse } = renamedThenReferenced();
      const restored = applyWorkbookOperations(withFormula, inverse);
      expect(restored.worksheets[0].name).toBe('Sheet1');
      expect(restored.worksheets[1].data[0]).toEqual(["='Sheet1'!A1", "='Sheet1'!B2"]);
    },
  );
});
