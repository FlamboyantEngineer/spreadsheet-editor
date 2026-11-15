// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { createRef, useState } from 'react';
import { describe, expect, it, vi } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { workbookCalculationKey } from '../calculation/publication.js';
import type { WorkbookStoreAdapter } from '../host/adapters.js';
import { FULL_CAPABILITIES, READ_ONLY_CAPABILITIES } from '../host/capabilities.js';
import { createMemoryHost } from '../host/memory-host.js';
import { createMemoryPresenceRoom } from '../host/memory-presence.js';
import type { SpreadsheetPeer } from '../host/presence-adapters.js';
import { createWorkbookSession } from '../host/workbook-session.js';
import { requireDefined } from '../model/require-defined.js';
import { previewSuggestions } from '../model/review.js';
import { copyWorkbook, createWorkbook } from '../model/workbook.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorProps } from './contracts.js';
import { editorMenus } from './editor-menus.js';
import { formulaFunctionCatalog } from './public-host.js';
import { useSpreadsheetEditor } from './use-editor.js';

const reviewerIdentity = createMemoryHost({
  reader: { id: 'reviewer', name: 'Reviewer' },
}).identity;

/** A loaded session over a memory store whose `save` is a spy. */
async function openSession(data: string[][]) {
  const inner = createMemoryHost({ workbooks: { book: createWorkbook(data) } }).workbooks;
  const store = {
    inner,
    load: inner.load,
    save: vi.fn<WorkbookStoreAdapter['save']>((id, save) => inner.save(id, save)),
  };
  const session = createWorkbookSession({
    workbooks: store,
    workbookId: 'book',
    autosaveDelay: Infinity,
  });
  await session.ready;
  return { session, store };
}

describe('SE-HOST-1.1 and SE-HOST-2.1: accepted transactions and session lifecycle', () => {
  it('restores an uncommitted host draft and controls host panels through the stable handle', () => {
    const editorRef = createRef<SpreadsheetEditorHandle>(),
      onChange = vi.fn(),
      onPanelChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['accepted']]),
        editorRef,
        onChange,
        onPanelChange,
      }),
    );
    act(() => {
      requireDefined(editorRef.current).restoreDraft({
        sheetId: 'sheet-1',
        x: 0,
        y: 0,
        text: 'retained draft',
      });
    });
    expect(result.current.draftValue).toBe('retained draft');
    expect(requireDefined(editorRef.current).getSnapshot().worksheets[0].data[0][0]).toBe(
      'accepted',
    );
    expect(onChange).not.toHaveBeenCalled();
    act(() => {
      requireDefined(editorRef.current).setPanel('workflow');
    });
    expect(onPanelChange).toHaveBeenLastCalledWith({ id: 'workflow', expanded: true });
    act(() => {
      requireDefined(editorRef.current).setPanel(null);
    });
    expect(onPanelChange).toHaveBeenLastCalledWith({ id: 'workflow', expanded: false });
    act(() => {
      requireDefined(editorRef.current).cancelDraft();
    });
    expect(requireDefined(editorRef.current).getState().dirty).toBe(false);
  });

  it('notifies an inline panel callback only when the panel changes', () => {
    const editorRef = createRef<SpreadsheetEditorHandle>(),
      states: { id: string | null; expanded: boolean }[] = [];
    let renders = 0;
    const { result } = renderHook(() => {
      renders += 1;
      const [, setPanelState] = useState<{ id: string | null; expanded: boolean } | null>(null);
      return useSpreadsheetEditor({
        defaultValue: createWorkbook([['accepted']]),
        editorRef,
        onPanelChange: (state) => {
          states.push(state);
          // Stop storing after 50 renders so a regression fails instead of hanging.
          if (renders < 50) setPanelState(state);
        },
      });
    });
    expect(renders).toBeLessThan(10);
    expect(states).toEqual([{ id: null, expanded: false }]);
    act(() => {
      requireDefined(editorRef.current).setPanel('workflow');
    });
    expect(states).toEqual([
      { id: null, expanded: false },
      { id: 'workflow', expanded: true },
    ]);
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('accepted');
  });

  it('lets the host own creation and imports without exposing replacement actions', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['accepted']]),
        features: { newWorkbook: false, import: false },
      }),
    );
    const fileIds = editorMenus(result.current).file.map((item) => item.id);
    expect(fileIds).not.toContain('new');
    expect(fileIds).not.toContain('import');
    expect(fileIds).toContain('export');
    act(() => {
      expect(result.current.replaceWorkbook(createWorkbook(), 'New', 'new-workbook')).toBe(false);
    });
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('accepted');
  });
  it('rejects unsupported host formats before publishing a partial change', () => {
    const onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1']]),
        onChange,
        formatting: { kinds: ['general', 'number'], decimalPlaces: false },
      }),
    );
    act(() => {
      expect(() => result.current.format({}, { kind: 'fraction' })).toThrow(
        'The host does not support this number format.',
      );
    });
    act(() => {
      expect(() => result.current.format({}, { kind: 'number', decimals: 4 })).toThrow(
        'The host does not support decimal-place overrides.',
      );
    });
    expect(onChange).not.toHaveBeenCalled();
    act(() => {
      result.current.format({}, { kind: 'number' });
    });
    expect(onChange).toHaveBeenCalledTimes(1);
  });
  it('reports formula and native drafts without publishing them as accepted data', () => {
    const onDraftChange = vi.fn(),
      onChange = vi.fn();
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['accepted']]),
        onDraftChange,
        onChange,
      }),
    );
    act(() => {
      result.current.setDraftValue('formula draft');
    });
    expect(onDraftChange).toHaveBeenLastCalledWith({
      sheetId: 'sheet-1',
      x: 0,
      y: 0,
      text: 'formula draft',
    });
    act(() => {
      result.current.cancelFormula();
    });
    expect(onDraftChange).toHaveBeenLastCalledWith(null);
    act(() => {
      result.current.setNativeDraft({ x: 1, y: 2, text: 'native draft' });
    });
    expect(onDraftChange).toHaveBeenLastCalledWith({
      sheetId: 'sheet-1',
      x: 1,
      y: 2,
      text: 'native draft',
    });
    expect(onChange).not.toHaveBeenCalled();
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('accepted');
  });

  it('publishes separately keyed authoritative review results without calculating locally', () => {
    const workbook = createWorkbook([['2', '=A1*3']]);
    const accepted = {
      inputKey: workbookCalculationKey(workbook),
      revision: 1,
      status: 'ready' as const,
      result: calculateWorkbook(workbook),
    };
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          defaultValue: workbook,
          defaultMode: 'review',
          identity: reviewerIdentity,
          calculation: { mode: 'external', snapshot: accepted },
        } as SpreadsheetEditorProps,
      },
    );
    act(() => {
      result.current.proposeCells([['7']], 0, 0);
    });
    expect(result.current.visibleCalculation.result).toBeNull();
    const projected = previewSuggestions(result.current.workbook, 'reviewer');
    const publication = {
      inputKey: workbookCalculationKey(projected),
      revision: 2,
      status: 'ready' as const,
      result: {
        schemaVersion: 1 as const,
        worksheets: [
          {
            id: 'sheet-1',
            values: [
              [
                { kind: 'number' as const, numerator: '7', denominator: '1', approximate: false },
                { kind: 'number' as const, numerator: '97', denominator: '1', approximate: false },
              ],
            ],
          },
        ],
      },
    };
    rerender({
      defaultValue: workbook,
      defaultMode: 'review',
      identity: reviewerIdentity,
      calculation: { mode: 'external', snapshot: accepted },
      reviewCalculation: publication,
    });
    expect(result.current.visibleCalculation.result?.worksheets[0].values[0][1]).toMatchObject({
      numerator: '97',
    });
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('2');
  });
  it('lets a commenter post without authoring content or suggesting edits', () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['accepted']]),
        capabilities: { ...READ_ONLY_CAPABILITIES, comment: true },
      }),
    );
    expect(result.current.mode).toBe('review');
    expect(result.current.canEdit).toBe(false);
    expect(result.current.canSuggest).toBe(false);
    act(() => {
      result.current.postComment('Comment only');
    });
    expect(result.current.workbook.comments?.[0].entries[0].text).toBe('Comment only');
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('accepted');
  });

  it('uses only the authoritative host catalog when builtins are disabled', () => {
    expect(
      formulaFunctionCatalog(
        [{ name: 'SUM', signature: 'SUM(…)', description: 'Host sum' }],
        false,
      ),
    ).toEqual([{ name: 'SUM', signature: 'SUM(…)', description: 'Host sum', aliases: [] }]);
    expect(formulaFunctionCatalog([], false)).toEqual([]);
    expect(formulaFunctionCatalog().some((entry) => entry.name === 'SUM')).toBe(true);
  });
  it('uses the host factory for worksheet identities and publishes that same identity', () => {
    const onChange = vi.fn();
    let sequence = 0;
    const idFactory = (kind: string) => `host-${kind}-${++sequence}`;
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook(), idFactory, onChange }),
    );
    act(() => {
      result.current.sheetCommand('add');
    });
    const sheet = requireDefined(result.current.workbook.worksheets.at(-1));
    expect(sheet.id).toBe('host-sheet-1');
    expect(onChange.mock.calls[0][1].operations).toContainEqual(
      expect.objectContaining({
        kind: 'sheet-add',
        sheet: expect.objectContaining({ id: sheet.id }),
      }),
    );
  });
  it('rejects the host transaction before accepting data/history and retains its draft for save/flush', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { session, store } = await openSession([['before']]);
    const onChange = vi.fn(() => {
      throw new Error('Stale row identity');
    });
    const { result } = renderHook(() => useSpreadsheetEditor({ session, editorRef, onChange }));
    act(() => {
      result.current.setDraftValue('after');
    });
    act(() => {
      expect(() => requireDefined(editorRef.current).commitDraft()).toThrow('Stale row identity');
    });
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('before');
    expect(result.current.history.past).toHaveLength(0);
    expect(requireDefined(editorRef.current).getState().draftDirty).toBe(true);
    await act(async () => {
      await expect(requireDefined(editorRef.current).save()).rejects.toThrow('Stale row identity');
    });
    await act(async () => {
      await expect(requireDefined(editorRef.current).flush()).rejects.toThrow('Stale row identity');
    });
    expect(store.save).not.toHaveBeenCalled();
    expect(session.status().dirty).toBe(false);
    expect(result.current.notice).toBe('Stale row identity');
  });
  it('SE-SESSION-1.4: commits a draft once, then saves it through the session', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>(),
      order: string[] = [];
    const { session, store } = await openSession([['0']]);
    store.save.mockImplementation(async (id, save) => {
      order.push('stored');
      expect(save.content.worksheets[0].data[0][0]).toBe('1E-25');
      return store.inner.save(id, save);
    });
    const onChange = vi.fn((_next, context) => {
      order.push('transaction');
      expect(context.operations[0].kind).toBe('cells');
    });
    const { result } = renderHook(() => useSpreadsheetEditor({ session, editorRef, onChange }));
    const handle = requireDefined(editorRef.current);
    act(() => {
      result.current.setDraftValue('1E-25');
    });
    await act(() => handle.flush());
    expect(order).toEqual(['transaction', 'stored']);
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(editorRef.current).toBe(handle);
    expect(handle.getState().dirty).toBe(false);
    expect(session.status().revision).toBe('r2');
    const snapshot = handle.getSnapshot();
    snapshot.worksheets[0].data[0][0] = 'tampered';
    expect(handle.getSnapshot().worksheets[0].data[0][0]).toBe('1E-25');
  });
  it('SE-SESSION-1.2: keeps later local edits dirty while an earlier save is pending', async () => {
    let finish!: () => void;
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { session, store } = await openSession([['a']]);
    store.save.mockImplementation(
      (id, save) =>
        new Promise((resolve) => {
          finish = () => resolve(store.inner.save(id, save));
        }),
    );
    const { result } = renderHook(() => useSpreadsheetEditor({ session, editorRef }));
    act(() => {
      result.current.proposeCells([['b']], 0, 0);
    });
    let flush!: Promise<void>;
    act(() => {
      flush = session.save();
    });
    await waitFor(() => expect(finish).toBeTypeOf('function'));
    act(() => {
      result.current.proposeCells([['c']], 0, 0);
    });
    await act(async () => {
      finish();
      await flush;
    });
    expect(requireDefined(editorRef.current).getState().dirty).toBe(true);
    expect(store.save.mock.calls[0][1].content.worksheets[0].data[0][0]).toBe('b');
  });
  it('rejects a failed save and invalidates a disposed handle', async () => {
    const editorRef = createRef<SpreadsheetEditorHandle>();
    const { session, store } = await openSession([['a']]);
    store.save.mockResolvedValue({
      ok: false,
      error: { kind: 'unavailable', message: 'Receipt refused' },
    });
    const { result, unmount } = renderHook(() => useSpreadsheetEditor({ session, editorRef }));
    const handle = requireDefined(editorRef.current);
    act(() => {
      result.current.setDraftValue('b');
    });
    await act(async () => {
      await expect(handle.flush()).rejects.toThrow('Receipt refused');
    });
    expect(handle.getState()).toMatchObject({ dirty: true, error: 'Receipt refused' });
    expect(result.current.savePhase).toBe('error');
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
    act(() => {
      result.current.undo();
    });
    expect(undo).toHaveBeenCalledOnce();
    expect(onChange).not.toHaveBeenCalled();
    expect(requireDefined(editorRef.current).getState().dirty).toBe(false);
  });
  it('records a blank insertion and its undo as semantic line commands using the same generated IDs', () => {
    const onChange = vi.fn(),
      ids = vi.fn((kind: string) => `${kind}-${ids.mock.calls.length}`);
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook(), onChange, idFactory: ids }),
    );
    act(() => {
      result.current.structure('row', 'insert');
    });
    expect(onChange.mock.calls[0][1].operations[0]).toMatchObject({
      kind: 'lines',
      axis: 'row',
      action: 'insert',
      count: 1,
    });
    const id = onChange.mock.calls[0][1].operations[0].ids[0];
    expect(requireDefined(result.current.sheet.rowIds)[0]).toBe(id);
    act(() => {
      result.current.undo();
    });
    expect(onChange.mock.calls[1][1].source).toBe('undo');
    expect(onChange.mock.calls[1][1].operations).toContainEqual(
      expect.objectContaining({ kind: 'lines', action: 'delete', ids: [id] }),
    );
  });
});

describe('SE-HOST-3.1 and SE-HOST-4.1: authoritative values, drafts and result search', () => {
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
    expect(requireDefined(result.current.output)[0][0]).toMatchObject({ numerator: '1' });
    const next = structuredClone(snapshot);
    next.revision = 4;
    next.result.worksheets[0].values[0][0] = 'fresh host text';
    rerender({ value: workbook, calculation: { mode: 'external', snapshot: next }, onChange });
    expect(requireDefined(result.current.output)[0][0]).toBe('fresh host text');
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
    act(() => {
      result.current.setDraftValue('3');
    });
    expect(result.current.workbook.worksheets[0].data[0][0]).toBe('2');
    expect(result.current.draftCell).toMatchObject({ x: 0, y: 0, text: '3' });
    expect(requireDefined(result.current.output)[0].slice(0, 2)).toEqual([null, null]);
    expect(requireDefined(result.current.output)[0][2]).toMatchObject({ numerator: '8' });
    act(() => {
      result.current.cancelFormula();
    });
    expect(requireDefined(result.current.output)[0][1]).toMatchObject({ numerator: '4' });
  });
  it('navigates result matches while replacing authored matches only', async () => {
    const { result } = renderHook(() =>
      useSpreadsheetEditor({ defaultValue: createWorkbook([['=2+2', '4']]) }),
    );
    await waitFor(() => expect(result.current.output?.[0][0]).toMatchObject({ numerator: '4' }));
    act(() => {
      result.current.setQuery('4');
    });
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
      requireDefined(
        formulaFunctionCatalog([{ name: 'foo', signature: 'FOO()', description: 'host' }]).at(-1),
      ).name,
    ).toBe('FOO');
  });
});

describe('SE-REVIEW-1.1: comment authors are named through identity', () => {
  it('resolves stored person ids and shows an unknown id as itself', async () => {
    const workbook = createWorkbook([['1']]);
    workbook.comments = [
      {
        id: 'thread',
        sheetId: 'sheet-1',
        range: 'A1',
        state: 'open',
        entries: [
          { id: 'one', author: 'morgan', text: 'First', createdAt: '2026-11-10T12:00:00Z' },
          {
            id: 'two',
            author: 'gone',
            text: 'Second',
            createdAt: '2026-11-10T12:01:00Z',
            reactions: { '+1': ['reviewer'] },
          },
        ],
      },
    ];
    const identity = createMemoryHost({
      reader: { id: 'reviewer', name: 'Reviewer' },
      people: [{ id: 'morgan', name: 'Morgan' }],
    }).identity;
    const resolvePeople = vi.spyOn(identity, 'resolvePeople');
    const { result } = renderHook(() => useSpreadsheetEditor({ defaultValue: workbook, identity }));
    expect(result.current.personName('morgan')).toBe('morgan');
    await waitFor(() => expect(result.current.personName('morgan')).toBe('Morgan'));
    expect(result.current.personName('gone')).toBe('gone');
    expect(result.current.personName('reviewer')).toBe('Reviewer');
    expect(resolvePeople).toHaveBeenCalledWith(['gone', 'morgan']);
  });
});

describe('SE-ADAPTERS-4.3: the editor shares presence through the adapter', () => {
  it('publishes the place and lock, shows the others and withdraws on unmount', () => {
    const room = createMemoryPresenceRoom();
    const reader = createMemoryHost({
      reader: { id: 'reviewer', name: 'Reviewer' },
      presence: room,
    });
    const other = room.join({ id: 'morgan', name: 'Morgan' });
    const seen: (readonly SpreadsheetPeer[])[] = [];
    const stop = other.subscribe('book', (peers) => seen.push(peers));
    other.publish('book', {
      at: { sheetId: 'sheet-1', range: 'B1' },
      editLock: { sheetId: 'sheet-1', range: 'B1' },
      status: 'active',
    });
    const { result, unmount } = renderHook(() =>
      useSpreadsheetEditor({
        defaultValue: createWorkbook([['1', '2']]),
        identity: reader.identity,
        presence: reader.presence,
        workbookId: 'book',
      }),
    );
    expect(result.current.peers).toMatchObject([
      { id: 'morgan', sheetId: 'sheet-1', editLock: { x1: 1, x2: 1, y1: 0, y2: 0 } },
    ]);
    expect(seen.at(-1)).toMatchObject([
      { id: 'reviewer', at: { sheetId: 'sheet-1', range: 'A1' }, editLock: null },
    ]);
    act(() => {
      result.current.publishPresence(true);
    });
    expect(seen.at(-1)?.[0].editLock).toEqual({ sheetId: 'sheet-1', range: 'A1' });
    act(() => {
      result.current.select({ x1: 1, x2: 1, y1: 0, y2: 0 });
    });
    expect(result.current.lockedSelection?.name).toBe('Morgan');
    expect(seen.at(-1)?.[0].at).toEqual({ sheetId: 'sheet-1', range: 'B1' });
    unmount();
    expect(seen.at(-1)).toEqual([]);
    stop();
  });
});

describe('SE-ADAPTERS-2.3: the editor takes identity and capabilities', () => {
  it('attributes the reader from identity and offers at most what the capabilities allow', () => {
    const { result, rerender } = renderHook(
      (props: SpreadsheetEditorProps) => useSpreadsheetEditor(props),
      {
        initialProps: {
          defaultValue: createWorkbook([['1']]),
          identity: reviewerIdentity,
          capabilities: { ...FULL_CAPABILITIES, decide: false },
        } as SpreadsheetEditorProps,
      },
    );
    expect(result.current.actor).toEqual({ id: 'reviewer', name: 'Reviewer' });
    expect(result.current.mode).toBe('author');
    expect(result.current.canEdit).toBe(true);
    expect(result.current.canDecide).toBe(false);
    act(() => {
      result.current.postComment('Mine');
    });
    expect(result.current.workbook.comments?.[0].entries[0].author).toBe('reviewer');
    rerender({
      ...result.current.props,
      capabilities: { ...READ_ONLY_CAPABILITIES, suggest: true },
    });
    expect(result.current.mode).toBe('review');
    expect(result.current.canSuggest).toBe(true);
    expect(result.current.canComment).toBe(false);
    act(() => {
      result.current.setMode('author');
    });
    expect(result.current.mode).toBe('review');
    rerender({ ...result.current.props, capabilities: READ_ONLY_CAPABILITIES });
    expect(result.current.mode).toBe('read');
    expect(result.current.locked).toBe(true);
  });
});
