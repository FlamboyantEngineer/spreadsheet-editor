// @vitest-environment jsdom
import { act, renderHook } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createMemoryHost } from '../host/memory-host.js';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import { useSpreadsheetEditor } from './use-editor.js';

const identity = createMemoryHost({ reader: { id: 'alex', name: 'Alex' } }).identity;

describe('SE-9: undo keeps operations, not workbook copies', () => {
  it('undoes and redoes every kind of step back to exactly the same workbook', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([
          ['1', '=A1*2'],
          ['3', 'text'],
        ]),
        identity,
      }),
    );
    const e = () => result.current;
    const select =
      (x1: number, y1: number, x2 = x1, y2 = y1) =>
      () =>
        void e().select({ x1, y1, x2, y2 });
    // Each step is a list of actions, each acted on its own render.
    const steps: [string, ...(() => unknown)[]][] = [
      [
        'edit beyond the extent',
        select(3, 4),
        () => e().setDraftValue('far'),
        () => e().applyFormula(),
      ],
      ['format', select(0, 0, 1, 1), () => e().format({ 'font-weight': 'bold' })],
      ['border', () => e().border('outside')],
      ['merge', () => e().merge()],
      ['unmerge', () => e().merge(true)],
      ['comment', () => e().beginComment(), () => e().postComment('Check this')],
      ['paste', () => e().pasteCells([['7', '8']], 0, 2)],
      ['clear', select(0, 1), () => e().clear()],
      ['rename', () => e().renameSheet('sheet-1', 'Data')],
      ['add sheet', () => e().sheetCommand('add')],
      ['reorder', () => e().sheetCommand('right', 'sheet-1')],
      ['insert row', select(0, 0), () => e().structure('row', 'insert')],
      [
        'replace',
        () => {
          e().setShowFind(true);
          e().setQuery('text');
        },
        () => e().replace(true, 'words'),
      ],
    ];
    const states: Workbook[] = [e().workbook];
    for (const [, ...actions] of steps) {
      for (const action of actions) act(() => void action());
      states.push(e().workbook);
    }
    expect(e().history.past).toHaveLength(steps.length);
    for (let index = steps.length - 1; index >= 0; index--) {
      act(() => e().undo());
      expect(e().workbook, `undo ${steps[index][0]}`).toStrictEqual(states[index]);
    }
    for (let index = 1; index <= steps.length; index++) {
      act(() => e().redo());
      expect(e().workbook, `redo ${steps[index - 1][0]}`).toStrictEqual(states[index]);
    }
    const kept = e().history.past.map((entry) => 'snapshot' in entry);
    expect(kept).toEqual(steps.map(([name]) => ['add sheet', 'insert row'].includes(name)));
  });
});
