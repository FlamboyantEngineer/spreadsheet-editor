// @vitest-environment jsdom
import { createRef } from 'react';
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { useSpreadsheetEditor } from './use-editor.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorProps } from './contracts.js';

const entry = { id: 'saved', timestamp: '2026-01-01T00:00:00Z', author: { id: 'a', name: 'Alex' } };

describe('SE-HOST-2: durability barriers cover each requested edit', () => {
  it('commits a newer draft and waits for a newer receipt when another flush is pending', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const receipts: (() => void)[] = [];
    const onFlush = vi.fn<NonNullable<SpreadsheetEditorProps['onFlush']>>(
      () => new Promise<void>((resolve) => receipts.push(resolve)),
    );
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['before']]), editorRef, onFlush }),
    );
    act(() => result.current.setDraftValue('first'));
    let first!: Promise<void>;
    act(() => {
      first = editorRef.current!.flush();
    });
    await waitFor(() => expect(onFlush).toHaveBeenCalledTimes(1));
    act(() => result.current.setDraftValue('second'));
    let second!: Promise<void>;
    let secondSettled = false;
    act(() => {
      second = editorRef.current!.flush().then(() => {
        secondSettled = true;
      });
    });
    expect(editorRef.current!.getSnapshot().worksheets[0].data[0][0]).toBe('second');
    await act(async () => {
      receipts[0]();
      await first;
    });
    await waitFor(() => expect(onFlush).toHaveBeenCalledTimes(2));
    expect(secondSettled).toBe(false);
    expect(onFlush.mock.calls[1][0].worksheets[0].data[0][0]).toBe('second');
    await act(async () => {
      receipts[1]();
      await second;
    });
    expect(editorRef.current!.getState().dirty).toBe(false);
  });

  it('aborts an active receipt when edit permission is removed', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    let finish!: () => void;
    const onFlush = vi.fn<NonNullable<SpreadsheetEditorProps['onFlush']>>(
      () =>
        new Promise<void>((resolve) => {
          finish = resolve;
        }),
    );
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          defaultValue: createWorkbook([['a']]),
          editorRef,
          onFlush,
        } as SpreadsheetEditorProps,
      },
    );
    act(() => result.current.proposeCells([['b']], 0, 0));
    let flushing!: Promise<void>;
    act(() => {
      flushing = editorRef.current!.flush();
      void flushing.catch(() => {});
    });
    await waitFor(() => expect(onFlush).toHaveBeenCalledTimes(1));
    rerender({ ...result.current.props, readOnly: true });
    expect(onFlush.mock.calls[0][1].signal.aborted).toBe(true);
    await act(async () => {
      finish();
      await expect(flushing).rejects.toThrow('cancelled');
    });
    expect(editorRef.current!.getState().dirty).toBe(true);
  });

  it('commits native input and awaits a flush-only host before history navigation', async () => {
    let finish!: () => void;
    const onFlush = vi.fn<NonNullable<SpreadsheetEditorProps['onFlush']>>(
      () =>
        new Promise<void>((resolve) => {
          finish = resolve;
        }),
    );
    const onOpen = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['a']]),
        onFlush,
        history: { entries: [entry], onOpen },
      }),
    );
    act(() => {
      result.current.setNativeDraft({ x: 0, y: 0, text: 'native' });
      result.current.grid.current = {
        select: vi.fn(),
        getAnchor: () => null,
        releaseKeyboard: vi.fn(),
        commitDraft: vi.fn().mockImplementationOnce(() => {
          result.current.proposeCells([['native']], 0, 0);
          result.current.setNativeDraft(null);
        }),
      };
    });
    let opening!: Promise<void>;
    act(() => {
      opening = result.current.openHistorical(entry);
    });
    await waitFor(() => expect(onFlush).toHaveBeenCalledTimes(1));
    expect(onFlush.mock.calls[0][0].worksheets[0].data[0][0]).toBe('native');
    expect(onOpen).not.toHaveBeenCalled();
    await act(async () => {
      finish();
      await opening;
    });
    expect(onOpen).toHaveBeenCalledExactlyOnceWith(entry);
  });

  it('keeps history closed after a refused receipt or an edit newer than its receipt', async () => {
    const onOpen = vi.fn();
    let finish!: () => void;
    const onFlush = vi
      .fn<NonNullable<SpreadsheetEditorProps['onFlush']>>()
      .mockRejectedValueOnce(new Error('Receipt refused'))
      .mockImplementationOnce(
        () =>
          new Promise<void>((resolve) => {
            finish = resolve;
          }),
      );
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['a']]),
        onFlush,
        history: { entries: [entry], onOpen },
      }),
    );
    act(() => result.current.proposeCells([['b']], 0, 0));
    await act(() => result.current.openHistorical(entry));
    expect(onOpen).not.toHaveBeenCalled();
    expect(result.current.notice).toBe('Receipt refused');
    let opening!: Promise<void>;
    act(() => {
      opening = result.current.openHistorical(entry);
    });
    await waitFor(() => expect(onFlush).toHaveBeenCalledTimes(2));
    act(() => result.current.proposeCells([['later']], 0, 0));
    await act(async () => {
      finish();
      await opening;
    });
    expect(onOpen).not.toHaveBeenCalled();
    expect(result.current.handle.getState().dirty).toBe(true);
  });

  it('does not manufacture a save or receipt while opening history from a pristine session', async () => {
    const onSave = vi.fn(),
      onFlush = vi.fn(),
      onOpen = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['a']]),
        onSave,
        onFlush,
        history: { entries: [entry], onOpen },
      }),
    );
    await act(() => result.current.openHistorical(entry));
    expect(onOpen).toHaveBeenCalledExactlyOnceWith(entry);
    expect(onSave).not.toHaveBeenCalled();
    expect(onFlush).not.toHaveBeenCalled();
  });
});
