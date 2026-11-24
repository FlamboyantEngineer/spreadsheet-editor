// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { createMemoryPresenceRoom } from '../host/memory-presence.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetClipboardData } from './clipboard-payload.js';
import { useSpreadsheetEditor } from './use-editor.js';

const at = (x: number) => ({ x1: x, x2: x, y1: 0, y2: 0 });

function richClipboard() {
  let data: SpreadsheetClipboardData = { text: '' };
  return {
    readText: async () => data.text,
    writeText: async (text: string) => {
      data = { text };
    },
    rich: {
      read: async () => data,
      write: async (next: SpreadsheetClipboardData) => {
        data = next;
      },
    },
  };
}

function heldRead() {
  let resolve = (_text: string) => {};
  const promise = new Promise<string>((finish) => {
    resolve = finish;
  });
  return { promise, resolve };
}

describe('SE-CHROME-3.1: native and host clipboard actions preserve ownership', () => {
  it('moves a verified cut exactly once and preserves a source when external text matches it', async () => {
    const clipboard = richClipboard();
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['=1+1']]), clipboard, onChange }),
    );
    await waitFor(() => expect(result.current.calculation.pending).toBe(false));
    await act(() => result.current.copy(true));
    expect(await clipboard.readText()).toBe('2');
    await clipboard.writeText('2');
    act(() => void result.current.select(at(1)));
    await act(() => result.current.paste());
    expect(result.current.sheet.data[0]).toEqual(['=1+1', '2']);

    act(() => void result.current.select(at(0)));
    await act(() => result.current.copy(true));
    act(() => void result.current.select(at(2)));
    await act(() => result.current.paste());
    expect(result.current.sheet.data[0]).toEqual([null, '2', '=1+1']);
    act(() => void result.current.select(at(3)));
    await act(() => result.current.paste());
    expect(result.current.sheet.data[0]).toEqual([null, '2', '=1+1', '2']);
    expect(onChange).toHaveBeenCalledTimes(3);
  });

  it('does not commit a delayed paste after the reader leaves Authoring', async () => {
    const held = heldRead();
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['old']]),
        clipboard: { readText: () => held.promise, writeText: async () => {} },
        onChange,
      }),
    );
    const paste = result.current.paste();
    act(() => result.current.setMode('read'));
    await act(async () => {
      held.resolve('new');
      await paste;
    });
    expect(result.current.sheet.data).toEqual([['old']]);
    expect(onChange).not.toHaveBeenCalled();
  });

  it('checks edit locks that arrive while the clipboard is being read', async () => {
    const held = heldRead();
    const onChange = vi.fn();
    const room = createMemoryPresenceRoom();
    const presence = room.join({ id: 'local', name: 'Local' });
    const peer = room.join({ id: 'remote', name: 'Remote' });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['old']]),
        workbookId: 'book',
        presence,
        clipboard: { readText: () => held.promise, writeText: async () => {} },
        onChange,
      }),
    );
    const paste = result.current.paste();
    const place = { sheetId: 'sheet-1', range: 'A1' };
    act(() => peer.publish('book', { at: place, editLock: place, status: 'active' }));
    await act(async () => {
      held.resolve('new');
      await paste;
    });
    expect(onChange).not.toHaveBeenCalled();
    expect(result.current.notice).toContain('Remote is editing a cell');
  });
});
