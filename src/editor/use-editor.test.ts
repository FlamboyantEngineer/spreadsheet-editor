// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import type { PresenceAdapter } from '../host/presence-adapters.js';
import { calculateWorkbook, type CalculationResult } from '../calculation/evaluator.js';
import { calculationCellKey } from '../calculation/publication.js';
import { writeCells } from '../model/editing.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { createWorkbookSession } from '../host/workbook-session.js';
import type {
  HostResult,
  SavedState,
  VersionStoreAdapter,
  WorkbookStoreAdapter,
  WorkbookVersion,
} from '../host/adapters.js';
import { useSpreadsheetEditor } from './use-editor.js';
import { createMemoryHost } from '../host/memory-host.js';
import { FULL_CAPABILITIES } from '../host/capabilities.js';

const actor = { id: 'alex', name: 'Alex' };
const identity = createMemoryHost({ reader: actor }).identity;
const reviewer = { ...FULL_CAPABILITIES, edit: false };
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
afterEach(() => vi.useRealTimers());

const peerLocks = new Map<string, PresenceAdapter>();
/**
 * A presence adapter with one peer, Morgan, writing in a cell of the first worksheet. It is the
 * same adapter for the same arguments, as a host's adapter stays the same between renders.
 */
function peerLock(id: string, range: string): PresenceAdapter {
  const key = `${id} ${range}`;
  const place = { sheetId: 'sheet-1', range };
  const known = peerLocks.get(key);
  if (known) return known;
  const adapter: PresenceAdapter = {
    publish: vi.fn(),
    subscribe(_workbookId, onPeers) {
      onPeers([{ id, name: 'Morgan', at: place, editLock: place, status: 'active' }]);
      return () => {};
    },
  };
  peerLocks.set(key, adapter);
  return adapter;
}

/**
 * A loaded session over a memory store whose `save` is a spy; with `hold`, each save waits until
 * the test calls its entry in `held`.
 */
async function openSession(
  data: string[][],
  options: { hold?: boolean; autosaveDelay?: number } = {},
) {
  const inner = createMemoryHost({ workbooks: { book: createWorkbook(data) } }).workbooks;
  const held: (() => void)[] = [];
  const store = {
    load: inner.load,
    save: vi.fn<WorkbookStoreAdapter['save']>((id, save) =>
      options.hold
        ? new Promise<HostResult<SavedState>>((resolve) => {
            held.push(() => resolve(inner.save(id, save)));
          })
        : inner.save(id, save),
    ),
  };
  const session = createWorkbookSession({
    workbooks: store,
    workbookId: 'book',
    autosaveDelay: options.autosaveDelay ?? Infinity,
  });
  await session.ready;
  return { session, store, held };
}
describe('SE-CR-2.12/3.4/3.6: capability and preview boundaries', () => {
  it('clamps reviewers, previews only their proposals, and blocks structural/save/decision commands', async () => {
    const { session, store } = await openSession([['1E-25']]);
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          session,
          identity,
          capabilities: reviewer,
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
    expect(store.save).not.toHaveBeenCalled();
    const workbook = result.current.workbook;
    rerender({ value: workbook, identity, session, capabilities: FULL_CAPABILITIES });
    expect(result.current.previewSheet.data[0][0]).toBe('1E-25');
    act(() => result.current.decide(proposal.id, 'accepted'));
    expect(result.current.workbook.suggestions![0].status).toBe('pending'); // controlled host must publish the result
  });
  it('records provenance for authored changes and refuses writes to an advisory peer lock', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1', '2']]),
        identity,
        workbookId: 'book',
        presence: peerLock('morgan', 'B1'),
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

describe('SE-SESSION-4 and SE-VH-1: saving through the session', () => {
  it('joins a save in flight and keeps changes made during it dirty', async () => {
    const { session, store, held } = await openSession([['1']], { hold: true });
    const { result } = renderHook(() => useSpreadsheetEditor({ session }));
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['2']]), 'Edit'),
    );
    let first!: Promise<void>;
    act(() => {
      first = result.current.save();
    });
    await waitFor(() => expect(held).toHaveLength(1));
    act(() => {
      void result.current.save();
    });
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(result.current.savePhase).toBe('saving');
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['3']]), 'Edit'),
    );
    await act(async () => {
      held[0]();
      await first;
    });
    expect(store.save.mock.calls[0][1].content.worksheets[0].data[0][0]).toBe('2');
    expect(result.current.savePhase).toBe('dirty');
    expect(result.current.lifecycleState.dirty).toBe(true);
  });
  it('autosaves through the session and saves before opening history', async () => {
    vi.useFakeTimers();
    const versions = versionStore();
    const { session, store } = await openSession([['1']], { autosaveDelay: 100 });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, versions, workbookId: 'book' }),
    );
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['2']]), 'Edit'),
    );
    await act(async () => {
      await vi.advanceTimersByTimeAsync(100);
    });
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(result.current.savePhase).toBe('saved');
    expect(result.current.saveAutomatic).toBe(true);
    act(() =>
      result.current.mutate(writeCells(result.current.workbook, 'sheet-1', 0, 0, [['3']]), 'Edit'),
    );
    await act(async () => {
      await result.current.openVersion(version);
    });
    expect(store.save).toHaveBeenCalledTimes(2);
    expect(result.current.viewedVersion).toEqual(version);
  });
  it('does not open a historical snapshot after a failed save or commit stale import results', async () => {
    const versions = versionStore();
    let resolveImport!: (value: ReturnType<typeof createWorkbook>) => void;
    const { session, store } = await openSession([['1']]);
    store.save.mockResolvedValue({
      ok: false,
      error: { kind: 'unavailable', message: 'Save failed' },
    });
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        session,
        versions,
        workbookId: 'book',
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
      await result.current.openVersion(version);
    });
    expect(result.current.viewedVersion).toBeNull();
    expect(result.current.notice).toBe('Save failed');
  });
  it('saves a pending formula before opening history even when the workbook was clean', async () => {
    const { session, store, held } = await openSession([['1']], { hold: true });
    const versions = versionStore();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ session, versions, workbookId: 'book' }),
    );
    act(() => result.current.setDraftValue('=1-1E-25'));
    expect(result.current.savePhase).toBe('saved');
    let opening!: Promise<unknown>;
    act(() => {
      opening = result.current.openVersion(version);
    });
    await waitFor(() => expect(store.save).toHaveBeenCalledTimes(1));
    expect(store.save.mock.calls[0][1].content.worksheets[0].data[0][0]).toBe('=1-1E-25');
    expect(result.current.viewedVersion).toBeNull();
    await act(async () => {
      held[0]();
      await opening;
    });
    expect(result.current.viewedVersion).toEqual(version);
    expect(result.current.draftDirty).toBe(false);
  });
  it('keeps history closed when a pending formula cannot commit or its save fails', async () => {
    const versions = versionStore();
    const { session, store } = await openSession([['1']]);
    store.save.mockResolvedValue({
      ok: false,
      error: { kind: 'unavailable', message: 'Save failed' },
    });
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          session,
          versions,
          workbookId: 'book',
          presence: peerLock('peer', 'A1'),
        } as SpreadsheetEditorProps,
      },
    );
    act(() => result.current.setDraftValue('1E-25'));
    await act(async () => {
      await result.current.openVersion(version);
    });
    expect(result.current.viewedVersion).toBeNull();
    expect(store.save).not.toHaveBeenCalled();
    expect(result.current.draftValue).toBe('1E-25');
    rerender({ session, versions, workbookId: 'book' });
    await act(async () => {
      await result.current.openVersion(version);
    });
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(result.current.viewedVersion).toBeNull();
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
      useSpreadsheetEditor({ defaultValue: createWorkbook([['1']]), identity, onChange }),
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
        identity,
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
        workbookId: 'book',
        presence: peerLock('peer', 'A1'),
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
    const onChange = vi.fn();
    const { session, store } = await openSession([['1']]);
    const { result } = renderHook(() => useSpreadsheetEditor({ session, onChange }));
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
    expect(store.save.mock.calls[0][1].content.worksheets[0].data[0][0]).toBe('=1-1E-25');
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

describe('SE-HOST-MODEL-2 and SE-PERF-1: change context', () => {
  it('gives the host a detached, frozen baseline and patches for the edited cell only', () => {
    const onChange = vi.fn<NonNullable<SpreadsheetEditorProps['onChange']>>();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([
          ['1', '2'],
          ['3', '4'],
        ]),
        identity,
        onChange,
      }),
    );
    const before = result.current.workbook;
    act(() => result.current.setDraftValue('5'));
    act(() => void result.current.applyFormula());
    const [next, context] = onChange.mock.calls[0];
    expect(next.worksheets[0].data[0][0]).toBe('5');
    expect(context!.baseline).toEqual(before);
    expect(context!.baseline).not.toBe(before);
    expect(context!.baseline).toBe(context!.baseline);
    expect(Object.isFrozen(context!.baseline.worksheets[0].data[0])).toBe(true);
    expect(context!.operations).toEqual([
      {
        kind: 'cells',
        sheetId: 'sheet-1',
        changes: [
          expect.objectContaining({
            address: 'A1',
            before: { value: '1' },
            after: expect.objectContaining({ value: '5', provenance: expect.anything() }),
          }),
        ],
      },
    ]);
  });
});

describe('SE-PERF-1: recalculating after an edit', () => {
  it('keeps cells that do not read the edited cell ready while the result is pending', async () => {
    const replies: { workbook: Workbook; finish(result: CalculationResult): void }[] = [];
    const provider = {
      calculate: (workbook: Workbook) =>
        new Promise<CalculationResult>((finish) => replies.push({ workbook, finish })),
    };
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1', '=A1*2', '5', '=C1+1']]),
        calculation: { mode: 'server', provider },
      }),
    );
    const key = (x: number) => calculationCellKey('sheet-1', x, 0);
    await waitFor(() => expect(replies).toHaveLength(1));
    await act(async () => replies[0].finish(calculateWorkbook(replies[0].workbook)));
    expect(result.current.readyCells.has(key(3))).toBe(true);
    act(() => result.current.setDraftValue('2'));
    act(() => void result.current.applyFormula());
    await waitFor(() => expect(replies).toHaveLength(2));
    expect(result.current.calculation.status).toBe('pending');
    expect(result.current.calculation.result).toBeNull();
    expect([0, 1, 2, 3].map((x) => result.current.readyCells.has(key(x)))).toEqual([
      false,
      false,
      true,
      true,
    ]);
    expect(result.current.output?.[0][3]).toMatchObject({ numerator: '6' });
    await act(async () => replies[1].finish(calculateWorkbook(replies[1].workbook)));
    expect(result.current.output?.[0][1]).toMatchObject({ numerator: '4' });
  });
});
