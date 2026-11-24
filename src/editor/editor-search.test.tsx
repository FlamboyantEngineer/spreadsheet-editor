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
    act(() => {
      result.current.setShowFind(true);
      result.current.setQuery('alpha');
    });
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
    act(() => {
      result.current.setShowFind(true);
      result.current.setQuery('target');
    });
    act(() => result.current.setDraftValue('pending'));
    act(() => expect(() => result.current.navigateMatch(1)).toThrow('Rejected edit'));
    expect(result.current.matchIndex).toBe(-1);
    expect(result.current.selection.x1).toBe(0);
  });

  it('retains a query and searches current worksheets when Find becomes visible again', () => {
    const { result, rerender } = renderHook(
      ({ value, search }) => useSpreadsheetEditor({ value, features: { search } }),
      { initialProps: { value: createWorkbook([['needle', '']]), search: true } },
    );
    act(() => {
      result.current.setShowFind(true);
      result.current.setQuery('needle');
    });
    expect(result.current.matches).toEqual([{ sheetId: 'sheet-1', x: 0, y: 0, authored: true }]);

    act(() => result.current.setShowFind(false));
    rerender({ value: createWorkbook([['', 'needle']]), search: true });
    expect(result.current.matches).toEqual([]);
    expect(result.current.query).toBe('needle');

    act(() => result.current.setShowFind(true));
    expect(result.current.matches).toEqual([{ sheetId: 'sheet-1', x: 1, y: 0, authored: true }]);
    rerender({ value: createWorkbook([['needle', '']]), search: false });
    expect(result.current.matches).toEqual([]);
    expect(result.current.showFind).toBe(false);
  });

  it('keeps a canonical input and creates no history for a replacement with identical text', () => {
    const workbook = createWorkbook([['1']]);
    const exact = { kind: 'number' as const, numerator: '1', denominator: '3', approximate: false };
    workbook.worksheets[0].inputValues = { A1: exact };
    const onChange = vi.fn();
    const { result } = renderHook(() => useSpreadsheetEditor({ defaultValue: workbook, onChange }));
    act(() => {
      result.current.setShowFind(true);
      result.current.setQuery('1');
    });
    act(() => expect(result.current.replace(true, '1')).toBe(0));
    expect(result.current.sheet.inputValues?.A1).toEqual(exact);
    expect(result.current.history.past).toHaveLength(0);
    expect(onChange).not.toHaveBeenCalled();

    act(() => expect(result.current.replace(false, '4')).toBe(1));
    expect(result.current.sheet.data[0][0]).toBe('4');
    expect(result.current.sheet.inputValues?.A1).toBeUndefined();
    expect(onChange).toHaveBeenCalledTimes(1);
  });

  it('keeps a protected draft when it makes the displayed replacement target stale', () => {
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['old value']]), onChange }),
    );
    act(() => {
      result.current.setShowFind(true);
      result.current.setQuery('old');
      result.current.setDraftValue('unsaved draft');
    });
    expect(result.current.matches).toHaveLength(1);
    act(() => expect(result.current.replace(false, 'new')).toBe(0));
    expect(result.current.draftValue).toBe('unsaved draft');
    expect(result.current.draftDirty).toBe(true);
    expect(result.current.sheet.data[0][0]).toBe('old value');
    expect(onChange).not.toHaveBeenCalled();
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
