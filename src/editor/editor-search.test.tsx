// @vitest-environment jsdom
import { act, renderHook } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { workbookCalculationKey } from '../calculation/publication.js';
import { createWorkbook } from '../model/workbook.js';
import { useSpreadsheetEditor } from './use-editor.js';

describe('SE-CHROME-1.1 and SE-CHROME-3.1: search, clipboard and selection statistics', () => {
  it('starts Previous at the last match and resets the active match for a new query', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['alpha', 'alpha', 'blue', 'alpha']]) }),
    );
    act(() => result.current.setQuery('alpha'));
    act(() => result.current.navigateMatch(-1));
    expect(result.current.selection.x1).toBe(3);
    expect(result.current.matchIndex).toBe(2);
    act(() => result.current.setQuery('blue'));
    expect(result.current.matchIndex).toBe(-1);
    act(() => result.current.navigateMatch(1));
    expect(result.current.selection.x1).toBe(2);
  });

  it('does not announce a match as active when committing the previous cell was rejected', () => {
    const onChange = vi.fn(() => {
      throw new Error('Rejected edit');
    });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['before', 'target']]), onChange }),
    );
    act(() => result.current.setQuery('target'));
    act(() => result.current.setDraftValue('pending'));
    act(() => expect(() => result.current.navigateMatch(1)).toThrow('Rejected edit'));
    expect(result.current.matchIndex).toBe(-1);
    expect(result.current.selection.x1).toBe(0);
  });

  it('copies authored inputs instead of blank or stale results while calculation is pending', async () => {
    const workbook = createWorkbook([['label', '=1+2']]);
    const writeText = vi.fn(async () => {});
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: workbook,
        calculation: {
          mode: 'external',
          snapshot: { inputKey: workbookCalculationKey(workbook), revision: 1, status: 'pending' },
        },
        clipboard: { readText: async () => '', writeText },
      }),
    );
    act(() => {
      result.current.select({ x1: 0, x2: 1, y1: 0, y2: 0 });
    });
    await act(() => result.current.copy());
    expect(writeText).toHaveBeenCalledExactlyOnceWith('label\t=1+2');
  });

  it('copies current calculated text when results are ready', async () => {
    const workbook = createWorkbook([['label', '=1+2']]);
    const writeText = vi.fn(async () => {});
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: workbook,
        calculation: {
          mode: 'external',
          snapshot: {
            inputKey: workbookCalculationKey(workbook),
            revision: 1,
            status: 'ready',
            result: calculateWorkbook(workbook),
          },
        },
        clipboard: { readText: async () => '', writeText },
      }),
    );
    act(() => {
      result.current.select({ x1: 0, x2: 1, y1: 0, y2: 0 });
    });
    await act(() => result.current.copy());
    expect(writeText).toHaveBeenCalledExactlyOnceWith('label\t3');
  });

  it('keeps quantity units in selection summaries and refuses incompatible aggregates', () => {
    const workbook = createWorkbook([['=1 m', '=2 m', '=3 s']]);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: workbook,
        calculation: {
          mode: 'external',
          snapshot: {
            inputKey: workbookCalculationKey(workbook),
            revision: 1,
            status: 'ready',
            result: calculateWorkbook(workbook),
          },
        },
      }),
    );
    act(() => {
      result.current.select({ x1: 0, x2: 1, y1: 0, y2: 0 });
    });
    expect(result.current.summary).toEqual({ count: 2, sum: '3 m', average: '1.5 m' });
    act(() => {
      result.current.select({ x1: 0, x2: 2, y1: 0, y2: 0 });
    });
    expect(result.current.summary).toEqual({ count: 3, sum: '—', average: '—' });
  });
});
