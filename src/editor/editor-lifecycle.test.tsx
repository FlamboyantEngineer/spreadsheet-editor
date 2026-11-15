// @vitest-environment jsdom
import { createRef } from 'react';
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { useSpreadsheetEditor } from './use-editor.js';
import type { SpreadsheetEditorHandle } from './contracts.js';
import { createMemoryHost } from '../host/memory-host.js';
import { createWorkbookSession } from '../host/workbook-session.js';
import type { HostResult, SavedState, WorkbookStoreAdapter } from '../host/adapters.js';

const entry = { id: 'saved', timestamp: '2026-01-01T00:00:00Z', author: { id: 'a', name: 'Alex' } };

/** A loaded session whose saves wait until the test answers them through `answers`. */
async function heldSession(data: string[][]) {
  const inner = createMemoryHost({ workbooks: { book: createWorkbook(data) } }).workbooks;
  const answers: ((result?: HostResult<SavedState>) => void)[] = [];
  const save = vi.fn<WorkbookStoreAdapter['save']>(
    (id, request) =>
      new Promise((resolve) => {
        answers.push((result) =>
          result ? resolve(result) : void inner.save(id, request).then(resolve),
        );
      }),
  );
  const session = createWorkbookSession({
    workbooks: { load: inner.load, save },
    workbookId: 'book',
    autosaveDelay: Infinity,
  });
  await session.ready;
  return { session, save, answers };
}

describe('SE-HOST-2 and SE-SESSION-4: flushes cover each requested edit', () => {
  it('commits a newer draft and waits for a newer save when another flush is pending', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { session, save, answers } = await heldSession([['before']]);
    const { result } = renderHook(() => useSpreadsheetEditor({ session, editorRef }));
    act(() => result.current.setDraftValue('first'));
    let first!: Promise<void>;
    act(() => {
      first = editorRef.current!.flush();
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
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
      answers[0]();
      await first;
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(secondSettled).toBe(false);
    expect(save.mock.calls[1][1].content.worksheets[0].data[0][0]).toBe('second');
    await act(async () => {
      answers[1]();
      await second;
    });
    expect(editorRef.current!.getState().dirty).toBe(false);
  });

  it('commits native input and saves it before history navigation', async () => {
    const { session, save, answers } = await heldSession([['a']]);
    const onOpen = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, history: { entries: [entry], onOpen } }),
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
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
    expect(save.mock.calls[0][1].content.worksheets[0].data[0][0]).toBe('native');
    expect(onOpen).not.toHaveBeenCalled();
    await act(async () => {
      answers[0]();
      await opening;
    });
    expect(onOpen).toHaveBeenCalledExactlyOnceWith(entry);
  });

  it('keeps history closed after a refused save or an edit newer than its save', async () => {
    const onOpen = vi.fn();
    const { session, save, answers } = await heldSession([['a']]);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, history: { entries: [entry], onOpen } }),
    );
    act(() => result.current.proposeCells([['b']], 0, 0));
    let refused!: Promise<void>;
    act(() => {
      refused = result.current.openHistorical(entry);
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
    await act(async () => {
      answers[0]({ ok: false, error: { kind: 'unavailable', message: 'Save refused' } });
      await refused;
    });
    expect(onOpen).not.toHaveBeenCalled();
    expect(result.current.notice).toBe('Save refused');
    let opening!: Promise<void>;
    act(() => {
      opening = result.current.openHistorical(entry);
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    act(() => result.current.proposeCells([['later']], 0, 0));
    await act(async () => {
      answers[1]();
      await opening;
    });
    expect(onOpen).not.toHaveBeenCalled();
    expect(result.current.handle.getState().dirty).toBe(true);
  });

  it('does not manufacture a save while opening history from a pristine session', async () => {
    const onOpen = vi.fn();
    const { session, save } = await heldSession([['a']]);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, history: { entries: [entry], onOpen } }),
    );
    await act(() => result.current.openHistorical(entry));
    expect(onOpen).toHaveBeenCalledExactlyOnceWith(entry);
    expect(save).not.toHaveBeenCalled();
  });
});
