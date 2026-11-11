// @vitest-environment jsdom
import { createRef } from 'react';
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { workbookCalculationKey } from '../calculation/publication.js';
import { createWorkbook, copyWorkbook } from '../model/workbook.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorProps } from './contracts.js';
import { formulaFunctionCatalog } from './public-host.js';
import { useSpreadsheetEditor } from './use-editor.js';

describe('SE-HOST-1/2: accepted transactions and session lifecycle', () => {
  it('rejects the host transaction before accepting data/history and retains its draft for save/flush', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>(),
      onSave = vi.fn();
    const onChange = vi.fn(() => {
      throw new Error('Stale row identity');
    });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['before']]),
        editorRef,
        onChange,
        onSave,
      }),
    );
    act(() => result.current.setDraftValue('after'));
    act(() => expect(() => editorRef.current!.commitDraft()).toThrow('Stale row identity'));
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('before');
    expect(result.current.history.past).toHaveLength(0);
    expect(editorRef.current!.getState().draftDirty).toBe(true);
    await act(async () => {
      await expect(editorRef.current!.save()).rejects.toThrow('Stale row identity');
    });
    await act(async () => {
      await expect(editorRef.current!.flush()).rejects.toThrow('Stale row identity');
    });
    expect(onSave).not.toHaveBeenCalled();
    expect(result.current.notice).toBe('Stale row identity');
  });
  it('commits a draft once before the durability barrier without a manual snapshot', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>(),
      order: string[] = [],
      onSave = vi.fn();
    const onChange = vi.fn((_next, context) => {
      order.push('transaction');
      expect(context.operations[0].kind).toBe('cells');
    });
    const onFlush = vi.fn(async (snapshot) => {
      order.push('stored');
      expect(snapshot.worksheets[0].data[0][0]).toBe('1E-25');
    });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['0']]),
        editorRef,
        onChange,
        onSave,
        onFlush,
      }),
    );
    const handle = editorRef.current!;
    act(() => result.current.setDraftValue('1E-25'));
    await act(() => handle.flush());
    expect(order).toEqual(['transaction', 'stored']);
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onSave).not.toHaveBeenCalled();
    expect(editorRef.current).toBe(handle);
    expect(handle.getState().dirty).toBe(false);
    const snapshot = handle.getSnapshot();
    snapshot.worksheets[0].data[0][0] = 'tampered';
    expect(handle.getSnapshot().worksheets[0].data[0][0]).toBe('1E-25');
  });
  it('keeps later local edits dirty while an earlier flush receipt is pending', async () => {
    let finish!: () => void;
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['a']]),
        editorRef,
        onFlush: () =>
          new Promise<void>((resolve) => {
            finish = resolve;
          }),
      }),
    );
    act(() => result.current.proposeCells([['b']], 0, 0));
    let flush!: Promise<void>;
    act(() => {
      flush = editorRef.current!.flush();
    });
    await waitFor(() => expect(finish).toBeTypeOf('function'));
    act(() => result.current.proposeCells([['c']], 0, 0));
    await act(async () => {
      finish();
      await flush;
    });
    expect(editorRef.current!.getState().dirty).toBe(true);
  });
  it('rejects a failed durability barrier and invalidates a disposed handle', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { result, unmount } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['a']]),
        editorRef,
        onFlush: async () => {
          throw new Error('Receipt refused');
        },
      }),
    );
    const handle = editorRef.current!;
    act(() => result.current.setDraftValue('b'));
    await act(async () => {
      await expect(handle.flush()).rejects.toThrow('Receipt refused');
    });
    expect(handle.getState()).toMatchObject({ dirty: true, error: 'Receipt refused' });
    unmount();
    expect(() => handle.getSnapshot()).toThrow('session has ended');
    await expect(handle.flush()).rejects.toThrow('session has ended');
  });
  it('does not dirty remote replacements and delegates undo across them', () => {
    const editorRef = createRef<SpreadsheetEditorHandle>(),
      undo = vi.fn(),
      onChange = vi.fn();
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          value: createWorkbook([['a']]),
          editorRef,
          undo: { canUndo: true, canRedo: false, undo, redo: vi.fn() },
          onChange,
        } as SpreadsheetEditorProps,
      },
    );
    const remote = copyWorkbook(result.current.workbook);
    remote.worksheets[0].data[0][0] = 'peer';
    rerender({ ...result.current.props, value: remote });
    act(() => result.current.undo());
    expect(undo).toHaveBeenCalledOnce();
    expect(onChange).not.toHaveBeenCalled();
    expect(editorRef.current!.getState().dirty).toBe(false);
  });
  it('records a blank insertion and its undo as semantic line commands using the same generated IDs', () => {
    const onChange = vi.fn(),
      ids = vi.fn((kind: string) => `${kind}-${ids.mock.calls.length}`);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook(), onChange, idFactory: ids }),
    );
    act(() => result.current.structure('row', 'insert'));
    expect(onChange.mock.calls[0][1].operations[0]).toMatchObject({
      kind: 'lines',
      axis: 'row',
      action: 'insert',
      count: 1,
    });
    const id = onChange.mock.calls[0][1].operations[0].ids[0];
    expect(result.current.sheet.rowIds![0]).toBe(id);
    act(() => result.current.undo());
    expect(onChange.mock.calls[1][1].source).toBe('undo');
    expect(onChange.mock.calls[1][1].operations).toContainEqual(
      expect.objectContaining({ kind: 'lines', action: 'delete', ids: [id] }),
    );
  });
});

describe('SE-HOST-3/4: authoritative values, drafts and result search', () => {
  it('accepts publications without an edit, rejects conflicting revisions and never calculates externally', () => {
    const workbook = createWorkbook([['=NOW()']]),
      onChange = vi.fn();
    const snapshot = {
      inputKey: workbookCalculationKey(workbook),
      revision: 3,
      status: 'ready' as const,
      result: calculateWorkbook(createWorkbook([['1']])),
    };
    snapshot.result.worksheets[0].id = workbook.worksheets[0].id;
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          value: workbook,
          calculation: { mode: 'external', snapshot },
          onChange,
        } as SpreadsheetEditorProps,
      },
    );
    expect(result.current.output![0][0]).toMatchObject({ numerator: '1' });
    const next = structuredClone(snapshot);
    next.revision = 4;
    next.result.worksheets[0].values[0][0] = 'fresh host text';
    rerender({ value: workbook, calculation: { mode: 'external', snapshot: next }, onChange });
    expect(result.current.output![0][0]).toBe('fresh host text');
    next.result.worksheets[0].values[0][0] = 'conflict';
    rerender({
      value: workbook,
      calculation: { mode: 'external', snapshot: structuredClone(next) },
      onChange,
    });
    expect(result.current.calculation.error).toContain('revision');
    expect(result.current.output).toBeUndefined();
    expect(onChange).not.toHaveBeenCalled();
  });
  it('paints a protected draft while only its actual dependents become pending', async () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['2', '=A1*2', '=7+1']]) }),
    );
    await waitFor(() => expect(result.current.output?.[0][2]).toMatchObject({ numerator: '8' }));
    act(() => result.current.setDraftValue('3'));
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('2');
    expect(result.current.draftCell).toMatchObject({ x: 0, y: 0, text: '3' });
    expect(result.current.output![0].slice(0, 2)).toEqual([null, null]);
    expect(result.current.output![0][2]).toMatchObject({ numerator: '8' });
    act(() => result.current.cancelFormula());
    expect(result.current.output![0][1]).toMatchObject({ numerator: '4' });
  });
  it('navigates result matches while replacing authored matches only', async () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['=2+2', '4']]) }),
    );
    await waitFor(() => expect(result.current.output?.[0][0]).toMatchObject({ numerator: '4' }));
    act(() => result.current.setQuery('4'));
    expect(result.current.matches).toHaveLength(2);
    let count = 0;
    act(() => {
      count = result.current.replace(true, '9') ?? 0;
    });
    expect(count).toBe(1);
    expect(result.current.sheet.data[0]).toEqual(['=2+2', '9']);
  });
  it('rejects duplicate host aliases and reserved built-ins', () => {
    expect(() =>
      formulaFunctionCatalog([{ name: 'sum', signature: 'SUM()', description: 'host' }]),
    ).toThrow('duplicate');
    expect(() =>
      formulaFunctionCatalog([
        { name: 'FOO', aliases: ['BAR'], signature: 'FOO()', description: 'host' },
        { name: 'bar', signature: 'BAR()', description: 'host' },
      ]),
    ).toThrow('duplicate');
    expect(
      formulaFunctionCatalog([{ name: 'foo', signature: 'FOO()', description: 'host' }]).at(-1)!
        .name,
    ).toBe('FOO');
  });
});
