// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { createRef } from 'react';
import { describe, expect, it, vi } from 'vitest';
import type {
  HostResult,
  SavedState,
  VersionStoreAdapter,
  WorkbookStoreAdapter,
  WorkbookVersion,
} from '../host/adapters.js';
import { createMemoryHost } from '../host/memory-host.js';
import { createWorkbookSession } from '../host/workbook-session.js';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetEditorHandle } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';

const version: WorkbookVersion = {
  revision: 'r0',
  sequence: 0,
  savedAt: '2026-01-01T00:00:00Z',
  author: 'alex',
  label: 'Saved',
};
/** A version store holding `version`, whose content is a one-cell workbook. */
function versionStore() {
  return {
    list: vi.fn<VersionStoreAdapter['list']>(async () => ({ ok: true, value: [version] })),
    load: vi.fn<VersionStoreAdapter['load']>(async () => ({
      ok: true,
      value: createWorkbook([['old']]),
    })),
  };
}

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

describe('SE-HOST-2.1 and SE-SESSION-1.4: flushes cover each requested edit', () => {
  it('commits a newer draft and waits for a newer save when another flush is pending', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { session, save, answers } = await heldSession([['before']]);
    const { result } = renderHook(() => useSpreadsheetEditor({ session, editorRef }));
    act(() => {
      result.current.setDraftValue('first');
    });
    let first!: Promise<void>;
    act(() => {
      first = requireDefined(editorRef.current).flush();
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
    act(() => {
      result.current.setDraftValue('second');
    });
    let second!: Promise<void>;
    let secondSettled = false;
    act(() => {
      second = requireDefined(editorRef.current)
        .flush()
        .then(() => {
          secondSettled = true;
        });
    });
    expect(requireDefined(editorRef.current).getSnapshot().worksheets[0].data[0][0]).toBe('second');
    await act(async () => {
      answers[0]();
      await first;
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    expect(secondSettled).toBe(false);
    expect(requireDefined(editorRef.current).getState().flushing).toBe(true);
    expect(save.mock.calls[1][1].content.worksheets[0].data[0][0]).toBe('second');
    await act(async () => {
      answers[1]();
      await second;
    });
    expect(requireDefined(editorRef.current).getState().dirty).toBe(false);
    expect(requireDefined(editorRef.current).getState().flushing).toBe(false);
  });

  it('commits native input and saves it before history navigation', async () => {
    const { session, save, answers } = await heldSession([['a']]);
    const versions = versionStore();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, versions, workbookId: 'book' }),
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
    let opening!: Promise<unknown>;
    act(() => {
      opening = result.current.openVersion(version);
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
    expect(save.mock.calls[0][1].content.worksheets[0].data[0][0]).toBe('native');
    expect(result.current.viewedVersion).toBeNull();
    await act(async () => {
      answers[0]();
      await opening;
    });
    expect(result.current.viewedVersion).toEqual(version);
  });

  it('keeps history closed after a refused save or an edit newer than its save', async () => {
    const versions = versionStore();
    const { session, save, answers } = await heldSession([['a']]);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, versions, workbookId: 'book' }),
    );
    act(() => {
      result.current.proposeCells([['b']], 0, 0);
    });
    let refused!: Promise<unknown>;
    act(() => {
      refused = result.current.openVersion(version);
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(1));
    await act(async () => {
      answers[0]({ ok: false, error: { kind: 'unavailable', message: 'Save refused' } });
      await refused;
    });
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.notice).toBe('Save refused');
    let opening!: Promise<unknown>;
    act(() => {
      opening = result.current.openVersion(version);
    });
    await waitFor(() => expect(save).toHaveBeenCalledTimes(2));
    act(() => {
      result.current.proposeCells([['later']], 0, 0);
    });
    await act(async () => {
      answers[1]();
      await opening;
    });
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.handle.getState().dirty).toBe(true);
  });

  it('does not manufacture a save while opening history from a pristine session', async () => {
    const versions = versionStore();
    const { session, save } = await heldSession([['a']]);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, versions, workbookId: 'book' }),
    );
    await act(() => result.current.openVersion(version));
    expect(result.current.viewedVersion).toEqual(version);
    expect(save).not.toHaveBeenCalled();
  });

  it('keeps newer edits when a version finishes opening or restoring', async () => {
    const { session } = await heldSession([['current']]);
    let finish!: () => void;
    const pending = new Promise<void>((resolve) => {
      finish = resolve;
    });
    const versions: VersionStoreAdapter = {
      list: async () => ({ ok: true, value: [version] }),
      load: async () => {
        await pending;
        return { ok: true, value: createWorkbook([['old']]) };
      },
      restore: async () => {
        await pending;
        return { ok: true, value: { revision: 'r2', sequence: 2 } };
      },
    };
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, versions, workbookId: 'book' }),
    );
    let opening!: Promise<unknown>;
    let restoring!: Promise<unknown>;
    act(() => {
      opening = result.current.openVersion(version);
      restoring = result.current.restoreVersion(version);
    });
    await act(async () => {});
    await act(async () => {
      result.current.proposeCells([['later']], 0, 0);
    });
    await act(async () => {
      finish();
      await expect(opening).resolves.toBe(false);
      await expect(restoring).resolves.toBe(false);
    });
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.sheet.data[0][0]).toBe('later');
    expect(session.status().dirty).toBe(true);
    session.close();
  });
});
