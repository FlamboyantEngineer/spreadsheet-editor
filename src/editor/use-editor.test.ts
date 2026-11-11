// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { writeCells } from '../model/editing.js';
import type { SpreadsheetEditorProps, SpreadsheetSaveContext } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';

const actor = { id: 'alex', name: 'Alex' };
afterEach(() => vi.useRealTimers());
describe('SE-CR-2.12/3.4/3.6: capability and preview boundaries', () => {
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

describe('SE-CR-1.11/1.12 and SE-VH-1: save lifecycle', () => {
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
    let cancelled!: Promise<void>;
    act(() => {
      cancelled = result.current.save();
      void cancelled.catch(() => {});
    });
    await waitFor(() => expect(requests).toHaveLength(2));
    rerender({ onSave, readOnly: true });
    expect(requests[1].context.signal.aborted).toBe(true);
    unmount();
    requests[1].finish();
    await expect(cancelled).rejects.toThrow('cancelled');
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
  it('saves a pending formula before opening history even when the workbook was clean', async () => {
    let finishSave!: () => void;
    const onSave = vi.fn<NonNullable<SpreadsheetEditorProps['onSave']>>(
      () =>
        new Promise<void>((resolve) => {
          finishSave = resolve;
        }),
    );
    const onOpen = vi.fn();
    const entry = { id: 'v1', timestamp: '2026-01-01T00:00:00Z', author: actor };
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1']]),
        onSave,
        history: { entries: [entry], onOpen },
      }),
    );
    act(() => result.current.setDraftValue('=1-1E-25'));
    expect(result.current.savePhase).toBe('idle');
    let opening!: Promise<void>;
    act(() => {
      opening = result.current.openHistorical(entry);
    });
    await waitFor(() => expect(onSave).toHaveBeenCalledTimes(1));
    expect(onSave.mock.calls[0][0].worksheets[0].data[0][0]).toBe('=1-1E-25');
    expect(onOpen).not.toHaveBeenCalled();
    await act(async () => {
      finishSave();
      await opening;
    });
    expect(onOpen).toHaveBeenCalledWith(entry);
    expect(result.current.draftDirty).toBe(false);
  });
  it('keeps history closed when a pending formula cannot commit or its save fails', async () => {
    const onOpen = vi.fn();
    const onSave = vi.fn(async () => {
      throw new Error('Save failed');
    });
    const entry = { id: 'v1', timestamp: '2026-01-01T00:00:00Z', author: actor };
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          defaultValue: createWorkbook([['1']]),
          onSave,
          history: { entries: [entry], onOpen },
          collaborators: [
            {
              id: 'peer',
              name: 'Morgan',
              sheetId: 'sheet-1',
              editLock: { x1: 0, x2: 0, y1: 0, y2: 0 },
            },
          ],
        } as SpreadsheetEditorProps,
      },
    );
    act(() => result.current.setDraftValue('1E-25'));
    await act(async () => {
      await result.current.openHistorical(entry);
    });
    expect(onOpen).not.toHaveBeenCalled();
    expect(onSave).not.toHaveBeenCalled();
    expect(result.current.draftValue).toBe('1E-25');
    rerender({ onSave, history: { entries: [entry], onOpen } });
    await act(async () => {
      await result.current.openHistorical(entry);
    });
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onOpen).not.toHaveBeenCalled();
    expect(result.current.sheet.data[0][0]).toBe('1E-25');
    expect(result.current.notice).toBe('Save failed');
  });
});

it('SE-CR-2.5: preserves a protected draft across a controlled update until apply or cancel', () => {
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

describe('SE-CHROME-2/4: worksheet targets and protected formula transactions', () => {
  it('discards pending drafts for toolbar undo/redo without adding them to workbook history', () => {
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1']]), onChange }),
    );
    act(() => result.current.proposeCells([['2']], 0, 0));
    act(() => result.current.setDraftValue('pending undo draft'));
    act(() => result.current.undo());
    expect(result.current.sheet.data[0][0]).toBe('1');
    expect(result.current.draftValue).toBe('1');
    expect(result.current.draftDirty).toBe(false);
    act(() => result.current.setDraftValue('pending redo draft'));
    act(() => result.current.redo());
    expect(result.current.sheet.data[0][0]).toBe('2');
    expect(result.current.draftValue).toBe('2');
    expect(result.current.draftDirty).toBe(false);
    expect(onChange).toHaveBeenCalledTimes(3);
    expect(result.current.history.past).toHaveLength(1);
  });
  it('combines a formula draft and formatting into one change and undo entry', () => {
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1']]), currentUser: actor, onChange }),
    );
    act(() => result.current.setDraftValue('=1-1E-25'));
    act(() => result.current.format({ 'font-weight': 'bold' }));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(result.current.sheet.data[0][0]).toBe('=1-1E-25');
    expect(result.current.sheet.style?.A1).toContain('font-weight: bold');
    expect(result.current.draftDirty).toBe(false);
    expect(result.current.history.past).toHaveLength(1);
    act(() => result.current.undo());
    expect(result.current.sheet.data[0][0]).toBe('1');
    expect(result.current.sheet.style?.A1).toBeUndefined();
  });
  it('attributes the draft before structural changes move its cell and formula references', () => {
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1'], ['=A1']]),
        currentUser: actor,
        onChange,
      }),
    );
    act(() => result.current.setDraftValue('1E-25'));
    act(() => result.current.structure('row', 'insert'));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(result.current.sheet.data[1][0]).toBe('1E-25');
    expect(result.current.sheet.data[2][0]).toBe('=A2');
    expect(result.current.sheet.cellProvenance?.A2.actor).toEqual(actor);
    expect(result.current.sheet.cellProvenance?.A1).toBeUndefined();
  });
  it('renames an inactive target and rewrites the draft without switching sheets', () => {
    const workbook = createWorkbook([['1']]);
    workbook.worksheets.push({ id: 'inputs', name: 'Inputs', data: [['2']] });
    const onChange = vi.fn();
    const { result } = renderHook(() => useSpreadsheetEditor({ defaultValue: workbook, onChange }));
    act(() => result.current.setDraftValue('=Inputs!A1+1E-25'));
    act(() => result.current.renameSheet('inputs', '  Input\n Data  '));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(result.current.sheet.id).toBe('sheet-1');
    expect(result.current.sheet.data[0][0]).toBe("='Input Data'!A1+1E-25");
    expect(result.current.workbook.worksheets[1].name).toBe('Input Data');
    expect(result.current.validateSheetName('inputs', ' Sheet1 ')).toBeTruthy();
    expect(result.current.validateSheetName('inputs', 'x'.repeat(257))).toBeTruthy();
    expect(result.current.validateSheetName('inputs', ' input  data ')).toBeNull();
    act(() => expect(result.current.renameSheet('missing', 'New')).toBe(false));
    expect(onChange).toHaveBeenCalledTimes(1);
  });
  it('moves an inactive target with wraparound and activates a duplicate only after success', () => {
    const workbook = createWorkbook([['1']]);
    workbook.worksheets.push(
      { id: 'inputs', name: 'Inputs', data: [['2']] },
      { id: 'other', name: 'Other', data: [['3']] },
    );
    const { result } = renderHook(() => useSpreadsheetEditor({ defaultValue: workbook }));
    act(() => result.current.sheetCommand('left', 'sheet-1'));
    expect(result.current.workbook.worksheets.map((sheet) => sheet.id)).toEqual([
      'other',
      'inputs',
      'sheet-1',
    ]);
    act(() => result.current.sheetCommand('right', 'inputs'));
    expect(result.current.sheet.id).toBe('sheet-1');
    act(() => result.current.sheetCommand('duplicate', 'inputs'));
    expect(result.current.sheet.data).toEqual([['2']]);
    expect(result.current.sheet.id).not.toBe('inputs');
  });
  it('refuses locked drafts without discarding them or moving selection or active worksheet', () => {
    const workbook = createWorkbook([['1', '2']]);
    workbook.worksheets.push({ id: 'inputs', name: 'Inputs', data: [['3']] });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: workbook,
        collaborators: [
          {
            id: 'peer',
            name: 'Morgan',
            sheetId: 'sheet-1',
            editLock: { x1: 0, x2: 0, y1: 0, y2: 0 },
          },
        ],
      }),
    );
    act(() => result.current.setDraftValue('1E-25'));
    act(() => {
      expect(result.current.applyFormula()).toBe(false);
      expect(result.current.select({ x1: 1, x2: 1, y1: 0, y2: 0 })).toBe(false);
      expect(result.current.activateSheet('inputs')).toBe(false);
      expect(result.current.sheetCommand('add')).toBe(false);
    });
    expect(result.current.workbook.worksheets).toHaveLength(2);
    expect(result.current.sheet.id).toBe('sheet-1');
    expect(result.current.selection.x1).toBe(0);
    expect(result.current.draftValue).toBe('1E-25');
  });
  it('flushes once before same-event navigation and does not reapply native committed drafts', () => {
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1', '2']]), onChange }),
    );
    act(() => result.current.setDraftValue('1E-25'));
    act(() => {
      expect(result.current.applyFormula()).toBe(true);
      expect(result.current.select({ x1: 1, x2: 1, y1: 0, y2: 0 })).toBe(true);
    });
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(result.current.selection.x1).toBe(1);
    act(() => result.current.setDraftValue('native draft'));
    act(() => {
      result.current.updateSheet(
        writeCells(result.current.workbook, 'sheet-1', 1, 0, [['native']]).worksheets[0],
      );
      result.current.onGridSelection({ x1: 0, x2: 0, y1: 0, y2: 0 });
    });
    expect(onChange).toHaveBeenCalledTimes(2);
    expect(result.current.sheet.data[0]).toEqual(['1E-25', 'native']);
    expect(result.current.selection.x1).toBe(0);
    expect(result.current.notice).toBeNull();
  });
  it('includes pending draft in a size dialog transaction and manual save snapshot', async () => {
    const onChange = vi.fn(),
      onSave = vi.fn<NonNullable<SpreadsheetEditorProps['onSave']>>(async () => {});
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1']]), onChange, onSave }),
    );
    act(() => result.current.setDraftValue('1E-25'));
    act(() => result.current.openDialog('column-width'));
    act(() => result.current.setDialogValue('140'));
    act(() => result.current.applyDialog());
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(result.current.sheet.data[0][0]).toBe('1E-25');
    expect(result.current.sheet.columns?.[0].width).toBe(140);
    act(() => result.current.setDraftValue('=1-1E-25'));
    await act(async () => {
      await result.current.save();
    });
    expect(onSave.mock.calls[0][0].worksheets[0].data[0][0]).toBe('=1-1E-25');
    expect(onChange).toHaveBeenCalledTimes(2);
  });
  it('records draft plus formatting as a review proposal while blocking worksheet changes', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1']]), defaultMode: 'review' }),
    );
    act(() => result.current.setDraftValue('1E-25'));
    act(() => result.current.format({ 'font-weight': 'bold' }));
    expect(result.current.sheet.data[0][0]).toBe('1');
    expect(result.current.previewSheet.data[0][0]).toBe('1E-25');
    expect(result.current.previewSheet.style?.A1).toContain('font-weight: bold');
    expect(result.current.workbook.suggestions).toHaveLength(1);
    act(() => {
      expect(result.current.renameSheet('sheet-1', 'Renamed')).toBe(false);
      expect(result.current.sheetCommand('add')).toBe(false);
    });
    expect(result.current.sheet.name).toBe('Sheet1');
  });
});

describe('SE-CHROME-3: asynchronous clipboard and protected drafts', () => {
  it('keeps a later cell draft when an earlier paste resolves after selection changes', async () => {
    let resolveRead!: (text: string) => void;
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1', '2']]),
        onChange,
        clipboard: {
          readText: () =>
            new Promise<string>((resolve) => {
              resolveRead = resolve;
            }),
          writeText: async () => {},
        },
      }),
    );
    let paste!: Promise<void>;
    act(() => {
      paste = result.current.paste();
    });
    act(() => result.current.select({ x1: 1, x2: 1, y1: 0, y2: 0 }));
    act(() => result.current.setDraftValue('=1-1E-25'));
    await act(async () => {
      resolveRead('7');
      await paste;
    });
    expect(result.current.sheet.data[0]).toEqual(['7', '2']);
    expect(result.current.selection.x1).toBe(1);
    expect(result.current.draftValue).toBe('=1-1E-25');
    expect(result.current.draftDirty).toBe(true);
    expect(onChange).toHaveBeenCalledTimes(1);
    act(() => {
      expect(result.current.applyFormula()).toBe(true);
    });
    expect(result.current.sheet.data[0]).toEqual(['7', '=1-1E-25']);
    expect(onChange).toHaveBeenCalledTimes(2);
    expect(result.current.notice).toBeNull();
  });
  it('refuses a delayed cut when the same cell has a newer protected draft', async () => {
    let resolveWrite!: () => void;
    const onChange = vi.fn(),
      writeText = vi.fn(
        () =>
          new Promise<void>((resolve) => {
            resolveWrite = resolve;
          }),
      );
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1']]),
        onChange,
        clipboard: { readText: async () => '2E-25', writeText },
      }),
    );
    act(() => result.current.setDraftValue('2E-25'));
    let cut!: Promise<void>;
    act(() => {
      cut = result.current.copy(true);
    });
    act(() => result.current.setDraftValue('3E-25'));
    await act(async () => {
      resolveWrite();
      await cut;
    });
    expect(writeText).toHaveBeenCalledExactlyOnceWith('2E-25');
    expect(result.current.sheet.data[0][0]).toBe('1');
    expect(result.current.draftValue).toBe('3E-25');
    expect(result.current.draftDirty).toBe(true);
    expect(onChange).not.toHaveBeenCalled();
    expect(result.current.notice).toContain('Try Cut again');
    // The completed clipboard write remains a copy; it cannot later delete the newer source.
    act(() => result.current.select({ x1: 1, x2: 1, y1: 0, y2: 0 }));
    await act(async () => {
      await result.current.paste();
    });
    expect(result.current.sheet.data[0]).toEqual(['3E-25', '2E-25']);
  });
});

describe('SE-CHROME-5: workbook replacement', () => {
  it('clears an old formula draft when a replacement reuses the active sheet id', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1']]) }),
    );
    act(() => result.current.setDraftValue('=1-1E-25'));
    act(() => result.current.replaceWorkbook(createWorkbook([['Replacement']]), 'New workbook'));
    expect(result.current.draftValue).toBe('Replacement');
    expect(result.current.draftDirty).toBe(false);
    act(() => result.current.applyFormula());
    expect(result.current.sheet.data[0][0]).toBe('Replacement');
  });
});
