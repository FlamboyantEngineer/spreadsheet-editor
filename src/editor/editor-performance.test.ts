// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';

/** Numbers in every column, and a formula in every tenth cell that reads its row's first cell. */
function workbookOf(rows: number, columns: number): Workbook {
  return createWorkbook(
    Array.from({ length: rows }, (_, y) =>
      Array.from({ length: columns }, (_, x) =>
        x === columns - 1 ? `=A${y + 1}+${x}` : String(y * columns + x),
      ),
    ),
  );
}

function elapsed(run: () => void): number {
  const started = performance.now();
  run();
  return performance.now() - started;
}

async function measure(rows: number, columns: number) {
  const { result } = renderHook(
    (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
    // A host's inline calculation prop is a new object on every render.
    { initialProps: { defaultValue: workbookOf(rows, columns) } as SpreadsheetEditorProps },
  );
  const ready = () =>
    waitFor(() => expect(result.current.calculation.status).toBe('ready'), {
      timeout: 60_000,
      interval: 1,
    });
  await ready();
  const firstKeystroke = elapsed(() => act(() => result.current.setDraftValue('4')));
  const keystroke = elapsed(() => act(() => result.current.setDraftValue('42')));
  const started = performance.now();
  const commit = elapsed(() => act(() => void result.current.applyFormula()));
  await ready();
  const edit = performance.now() - started;
  expect(result.current.workbook.worksheets[0].data[0][0]).toBe('42');
  return { firstKeystroke, keystroke, commit, edit };
}

// Budgets are generous for slow CI machines: about four times what a development machine takes.
// Before SE-PERF-1, 50,000 cells took about 320 ms per commit before calculation and 130 ms per
// keystroke, and about 920 ms for the whole edit, which includes recalculating every formula.
describe('SE-PERF-1: work per edit and per keystroke', () => {
  it('keeps a 10,000-cell workbook smooth', async () => {
    const time = await measure(1_000, 10);
    expect(time.keystroke).toBeLessThan(25);
    expect(time.firstKeystroke).toBeLessThan(50);
    expect(time.commit).toBeLessThan(100);
    expect(time.edit).toBeLessThan(400);
  }, 120_000);

  it('keeps a 50,000-cell workbook usable', async () => {
    const time = await measure(5_000, 10);
    expect(time.keystroke).toBeLessThan(40);
    expect(time.firstKeystroke).toBeLessThan(100);
    expect(time.commit).toBeLessThan(300);
    expect(time.edit).toBeLessThan(1_500);
  }, 120_000);
});
