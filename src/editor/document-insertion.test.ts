// @vitest-environment jsdom
import { act, renderHook } from '@testing-library/react';
import { expect, it, vi } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { calculationCellKey, workbookCalculationKey } from '../calculation/publication.js';
import { createWorkbook } from '../model/workbook.js';
import { useSpreadsheetEditor } from './use-editor.js';

it('SE-DOC-1.1: a clipped table requires the original merge anchor to be ready', () => {
  const workbook = createWorkbook([['Anchor', null, null]]);
  const sheet = workbook.worksheets[0];
  sheet.mergeCells = { A1: [3, 1] };
  const values = calculateWorkbook(workbook);
  values.worksheets[0].values[0][0] = null;
  const onInsert = vi.fn();
  const { result } = renderHook(() =>
    useSpreadsheetEditor({
      value: workbook,
      calculation: {
        mode: 'external',
        snapshot: {
          inputKey: workbookCalculationKey(workbook),
          revision: 1,
          status: 'ready',
          result: values,
          cellStates: { [calculationCellKey(sheet.id, 0, 0)]: 'unavailable' },
        },
      },
      onInsert,
    }),
  );
  act(() => {
    result.current.select({ x1: 1, y1: 0, x2: 2, y2: 0 });
  });
  expect(() => result.current.insert('table')).toThrow('current ready result');
  expect(onInsert).not.toHaveBeenCalled();
});
