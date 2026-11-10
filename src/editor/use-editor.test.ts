// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { writeCells } from '../model/editing.js';
import type { SpreadsheetEditorProps, SpreadsheetSaveContext } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';

const actor = { id: 'alex', name: 'Alex' };
afterEach(() => vi.useRealTimers());
describe('Host SE-CR-2.12/3.4/3.6: capability and preview boundaries', () => {
  it('clamps reviewers, previews only their proposals, and blocks structural/save/decision commands', async () => {
    const onSave = vi.fn();
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          defaultValue: createWorkbook([['1E-25']]),
          currentUser: actor,
          onSave,
          permissions: { edit: false },
        } as SpreadsheetEditorProps,
      },
    );
    expect(result.current.mode).toBe('review');
    act(() => result.current.proposeCells([['2E-25']], 0, 0));
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('1E-25');
    expect(result.current.previewSheet.data[0][0]).toBe('2E-25');
    const proposal = result.current.workbook.suggestions![0];
    act(() => {
      result.current.structure('row', 'insert');
      result.current.sheetCommand('add');
      result.current.decide(proposal.id, 'accepted');
    });
    await act(() => result.current.save());
    expect(result.current.workbook.worksheets).toHaveLength(1);
    expect(result.current.workbook.worksheets[0].data).toEqual([['1E-25']]);
    expect(onSave).not.toHaveBeenCalled();
    const workbook = result.current.workbook;
    rerender({ value: workbook, currentUser: actor, onSave, permissions: { edit: true } });
    expect(result.current.previewSheet.data[0][0]).toBe('1E-25');
    act(() => result.current.decide(proposal.id, 'accepted'));
    expect(result.current.workbook.suggestions![0].status).toBe('pending'); // controlled host must publish the result
  });
  it('records provenance for authored changes and refuses writes to an advisory peer lock', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1', '2']]),
        currentUser: actor,
        collaborators: [
          {
            id: 'morgan',
            name: 'Morgan',
            sheetId: 'sheet-1',
            editLock: { x1: 1, x2: 1, y1: 0, y2: 0 },
          },
        ],
      }),
    );
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['3']]), 'Edit'),
    );
    expect(result.current.workbook.worksheets[0].cellProvenance!.A1.actor).toEqual(actor);
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 1, 0, [['4']]), 'Edit'),
    );
    expect(result.current.workbook.worksheets[0].data[0][1]).toBe('2');
    expect(result.current.notice).toContain('Morgan');
  });
});

describe('Host SE-CR-1.11/1.12 and SE-VH-1: save lifecycle', () => {
  it('coalesces saves, preserves changes made in flight, and aborts on permission loss', async () => {
    const requests: { context: SpreadsheetSaveContext; finish(): void }[] = [];
    const onSave = vi.fn(
      (_workbook, context: SpreadsheetSaveContext) =>
        new Promise<void>((resolve) => requests.push({ context, finish: resolve })),
    );
    const { result, rerender, unmount } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: { defaultValue: createWorkbook([['1']]), onSave } as SpreadsheetEditorProps,
      },
    );
    let first!: Promise<void>;
    act(() => {
      first = result.current.save()!;
    });
    await waitFor(() => expect(requests).toHaveLength(1));
    act(() => {
      void result.current.save();
    });
    expect(onSave).toHaveBeenCalledTimes(1);
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['2']]), 'Edit'),
    );
    await act(async () => {
      requests[0].finish();
      await first;
    });
    expect(result.current.savePhase).toBe('dirty');
    act(() => {
      void result.current.save();
    });
    await waitFor(() => expect(requests).toHaveLength(2));
    rerender({ onSave, readOnly: true });
    expect(requests[1].context.signal.aborted).toBe(true);
    unmount();
    requests[1].finish();
  });
  it('retries a transient autosave with the same request id and saves before opening history', async () => {
    vi.useFakeTimers();
    const ids: string[] = [];
    const onSave = vi.fn(async (_workbook, context: SpreadsheetSaveContext) => {
      ids.push(context.requestId);
      if (ids.length === 1) throw Object.assign(new Error('Busy'), { status: 503 });
    });
    const onOpen = vi.fn();
    const entry = { id: 'v1', timestamp: '2026-01-01T00:00:00Z', author: actor };
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1']]),
        onSave,
        autosaveMs: 100,
        history: { entries: [entry], onOpen },
      }),
    );
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['2']]), 'Edit'),
    );
    await act(async () => {
      await vi.advanceTimersByTimeAsync(100);
    });
    expect(result.current.savePhase).toBe('saving');
    expect(result.current.notice).toBeNull();
    await act(async () => {
      await vi.advanceTimersByTimeAsync(500);
    });
    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe(ids[1]);
    expect(result.current.savePhase).toBe('saved');
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['3']]), 'Edit'),
    );
    await act(async () => {
      await result.current.openHistorical(entry);
    });
    expect(onSave).toHaveBeenCalledTimes(3);
    expect(onOpen).toHaveBeenCalledWith(entry);
  });
  it('does not open a historical snapshot after a failed save or commit stale import results', async () => {
    const onOpen = vi.fn();
    let resolveImport!: (value: ReturnType<typeof createWorkbook>) => void;
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1']]),
        onSave: async () => {
          throw new Error('Save failed');
        },
        history: { entries: [], onOpen },
        onImportXlsx: () =>
          new Promise((resolve) => {
            resolveImport = resolve;
          }),
      }),
    );
    let importing!: Promise<void>;
    act(() => {
      importing = result.current.importExcel();
    });
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['2']]), 'Edit'),
    );
    await act(async () => {
      resolveImport(createWorkbook([['99']]));
      await expect(importing).rejects.toThrow('workbook changed');
    });
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('2');
    await act(async () => {
      await result.current.openHistorical({
        id: 'v1',
        timestamp: '2026-01-01T00:00:00Z',
        author: actor,
      });
    });
    expect(onOpen).not.toHaveBeenCalled();
    expect(result.current.notice).toBe('Save failed');
  });
});

it('Host SE-CR-2.5: preserves a protected draft across a controlled update until apply or cancel', () => {
  const first = createWorkbook([['1']]);
  const { result, rerender } = renderHook(({ value }) => useSpreadsheetEditor({ value }), {
    initialProps: { value: first },
  });
  act(() => result.current.setDraftValue('=1-1E-25'));
  rerender({ value: writeCells(first, 'sheet-1', 0, 0, [['9']]) });
  expect(result.current.draftValue).toBe('=1-1E-25');
  expect(result.current.draftDirty).toBe(true);
  act(() => result.current.cancelFormula());
  expect(result.current.draftValue).toBe('9');
});
