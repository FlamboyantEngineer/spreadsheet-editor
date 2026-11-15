import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookChangeContext } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import type { HostResult, SavedState, WorkbookSave, WorkbookStoreAdapter } from './adapters.js';
import { createMemoryHost } from './memory-host.js';
import { createWorkbookSession, type WorkbookSession } from './workbook-session.js';

let counter = 0;
function change(session: WorkbookSession, value: string, transactionId = `t${++counter}`) {
  const next = createWorkbook([[value]]);
  const context: WorkbookChangeContext = {
    transactionId,
    source: 'user',
    baseline: requireDefined(session.workbook()),
    operations: [],
  };
  session.change(next, context);
  return transactionId;
}

/** A store whose saves wait until the test answers them. */
function heldStore(inner: WorkbookStoreAdapter) {
  const held: { save: WorkbookSave; answer: () => Promise<void> }[] = [];
  const store: WorkbookStoreAdapter = {
    load: (id, signal) => inner.load(id, signal),
    subscribe: inner.subscribe?.bind(inner),
    save: (id, save) =>
      new Promise<HostResult<SavedState>>((resolve) => {
        held.push({ save, answer: async () => resolve(await inner.save(id, save)) });
      }),
  };
  return { store, held };
}

function host() {
  return createMemoryHost({ workbooks: { book: createWorkbook([['0']]) } });
}

afterEach(() => vi.useRealTimers());

describe('SE-SESSION-1: opening a stored workbook', () => {
  it('loads the workbook at its revision and reports a missing one', async () => {
    const memory = host();
    const session = createWorkbookSession({ workbooks: memory.workbooks, workbookId: 'book' });
    expect(session.status().phase).toBe('loading');
    expect(await session.ready).toBe(true);
    expect(session.workbook()?.worksheets[0].data).toEqual([['0']]);
    expect(session.status()).toMatchObject({ phase: 'saved', revision: 'r1', dirty: false });
    const missing = createWorkbookSession({ workbooks: memory.workbooks, workbookId: 'gone' });
    expect(await missing.ready).toBe(false);
    expect(missing.status()).toMatchObject({ phase: 'unavailable', error: { kind: 'not-found' } });
  });
});

describe('SE-SESSION-2: causal save coverage', () => {
  it('keeps an edit made while a save is in flight dirty after that save answers', async () => {
    const memory = host();
    const { store, held } = heldStore(memory.workbooks);
    const session = createWorkbookSession({
      workbooks: store,
      workbookId: 'book',
      autosaveDelay: Infinity,
    });
    await session.ready;
    change(session, '1');
    const first = session.save();
    expect(session.status().phase).toBe('saving');
    change(session, '2');
    await held[0].answer();
    await first;
    expect(held[0].save.content.worksheets[0].data).toEqual([['1']]);
    expect(session.status()).toMatchObject({ phase: 'dirty', dirty: true, revision: 'r2' });
    const second = session.save();
    await held[1].answer();
    await second;
    expect(held[1].save).toMatchObject({ baseRevision: 'r2' });
    expect(session.status()).toMatchObject({ phase: 'saved', dirty: false, revision: 'r3' });
  });

  it('covers comment-only changes like any other transaction', async () => {
    const memory = host();
    const session = createWorkbookSession({
      workbooks: memory.workbooks,
      workbookId: 'book',
      autosaveDelay: Infinity,
    });
    await session.ready;
    const next: Workbook = { ...requireDefined(session.workbook()), comments: [] };
    next.comments = [
      {
        id: 'c',
        sheetId: 'sheet-1',
        range: 'A1',
        state: 'open',
        entries: [{ id: 'c', author: 'reader', text: 'Note', createdAt: '2026-11-15T12:00:00Z' }],
      },
    ];
    session.change(next, {
      transactionId: 'comment',
      source: 'user',
      baseline: requireDefined(session.workbook()),
      operations: [],
    });
    expect(session.status().dirty).toBe(true);
    await session.save();
    expect(session.status().dirty).toBe(false);
    const loaded = await memory.workbooks.load('book');
    expect(loaded.ok && loaded.value.content.comments?.[0].id).toBe('c');
  });

  it('clears only through an acknowledged transaction; unknown and repeated receipts clear nothing', async () => {
    const memory = host();
    const session = createWorkbookSession({
      workbooks: memory.workbooks,
      workbookId: 'book',
      autosaveDelay: Infinity,
    });
    await session.ready;
    const first = change(session, '1');
    change(session, '2');
    session.acknowledge('unknown');
    expect(session.status().dirty).toBe(true);
    session.acknowledge(first);
    expect(session.status().dirty).toBe(true);
    session.acknowledge(first);
    expect(session.status().dirty).toBe(true);
    const third = change(session, '3');
    session.acknowledge(third);
    expect(session.status()).toMatchObject({ phase: 'saved', dirty: false });
  });

  it('retries an unavailable store while autosaving, keeping newer edits dirty', async () => {
    vi.useFakeTimers();
    const memory = host();
    const session = createWorkbookSession({
      workbooks: memory.workbooks,
      workbookId: 'book',
      autosaveDelay: 10,
    });
    await session.ready;
    memory.memory.failNext('save', { kind: 'unavailable' });
    change(session, '1');
    await vi.advanceTimersByTimeAsync(10);
    expect(session.status().phase).toBe('saving');
    change(session, '2');
    await vi.advanceTimersByTimeAsync(510);
    // The retried save covered only '1'; the edit after it took a save of its own.
    expect(session.status()).toMatchObject({ phase: 'saved', dirty: false, revision: 'r3' });
    const loaded = await memory.workbooks.load('book');
    expect(loaded.ok && loaded.value.content.worksheets[0].data).toEqual([['2']]);
  });

  it('stops on a manual save failure and on a conflict, then keeps mine or takes the stored one', async () => {
    const memory = host();
    const session = createWorkbookSession({
      workbooks: memory.workbooks,
      workbookId: 'book',
      autosaveDelay: Infinity,
    });
    await session.ready;
    change(session, '1');
    memory.memory.failNext('save', { kind: 'unavailable', message: 'Offline' });
    await expect(session.save()).rejects.toThrow('Offline');
    expect(session.status()).toMatchObject({ phase: 'error', dirty: true });
    await memory.workbooks.save('book', {
      content: createWorkbook([['other']]),
      baseRevision: 'r1',
    });
    await expect(session.save()).rejects.toThrow('changed');
    expect(session.status()).toMatchObject({ phase: 'conflict', dirty: true });
    await expect(session.save()).rejects.toThrow();
    await session.keepMine();
    expect(session.status()).toMatchObject({ phase: 'saved', revision: 'r3' });
    await memory.workbooks.save('book', { content: createWorkbook([['x']]), baseRevision: 'r3' });
    change(session, '2');
    await expect(session.save()).rejects.toThrow();
    await session.takeStored();
    expect(session.workbook()?.worksheets[0].data).toEqual([['x']]);
    expect(session.status()).toMatchObject({ phase: 'saved', dirty: false, revision: 'r4' });
  });

  it('SE-SESSION-2: retains a failed-save state through edits until an explicit retry', async () => {
    vi.useFakeTimers();
    const memory = host();
    const save = vi.spyOn(memory.workbooks, 'save');
    const session = createWorkbookSession({
      workbooks: memory.workbooks,
      workbookId: 'book',
      autosaveDelay: 10,
      lifecycle: memory.lifecycle,
    });
    await session.ready;
    change(session, 'first');
    memory.memory.failNext('save', { kind: 'unavailable', message: 'Offline' });
    await expect(session.save()).rejects.toThrow('Offline');
    change(session, 'later');
    expect(session.status()).toMatchObject({
      phase: 'error',
      dirty: true,
      error: { message: 'Offline' },
    });
    memory.memory.emitLifecycle('hidden');
    await vi.advanceTimersByTimeAsync(1000);
    expect(save).toHaveBeenCalledTimes(1);
    expect(session.status().phase).toBe('error');
    await session.save();
    expect(save).toHaveBeenCalledTimes(2);
    expect(session.status()).toMatchObject({ phase: 'saved', dirty: false, error: null });
    session.close();
  });

  it('flushes pending edits and shows other saves only while clean', async () => {
    const memory = host();
    const session = createWorkbookSession({
      workbooks: memory.workbooks,
      workbookId: 'book',
      autosaveDelay: Infinity,
    });
    await session.ready;
    change(session, '1');
    await memory.workbooks.save('book', {
      content: createWorkbook([['other']]),
      baseRevision: 'r1',
    });
    await Promise.resolve();
    expect(session.workbook()?.worksheets[0].data).toEqual([['1']]);
    await session.takeStored();
    change(session, '2');
    await session.flush();
    expect(session.status()).toMatchObject({ phase: 'saved', dirty: false, revision: 'r3' });
    await memory.workbooks.save('book', {
      content: createWorkbook([['later']]),
      baseRevision: 'r3',
    });
    await vi.waitFor(() => expect(session.workbook()?.worksheets[0].data).toEqual([['later']]));
    expect(session.status().revision).toBe('r4');
  });
});

describe('SE-SESSION-3: drafts across a reload', () => {
  it('keeps unsaved content, offers it to the next session and clears it once saved', async () => {
    vi.useFakeTimers();
    const memory = host();
    const options = {
      workbooks: memory.workbooks,
      drafts: memory.drafts,
      workbookId: 'book',
      autosaveDelay: Infinity,
      draftDelay: 5,
      now: () => '2026-11-15T12:00:00Z',
    };
    const first = createWorkbookSession(options);
    await first.ready;
    change(first, 'unsaved');
    await vi.advanceTimersByTimeAsync(5);
    first.close();
    const second = createWorkbookSession(options);
    await second.ready;
    expect(second.draft()).toEqual({
      savedAt: '2026-11-15T12:00:00Z',
      baseRevision: 'r1',
      stale: false,
    });
    expect(second.workbook()?.worksheets[0].data).toEqual([['0']]);
    expect(second.restoreDraft()).toBe(true);
    expect(second.workbook()?.worksheets[0].data).toEqual([['unsaved']]);
    expect(second.status().dirty).toBe(true);
    await second.save();
    expect((await memory.drafts.get('book')).ok).toBe(true);
    expect(await memory.drafts.get('book')).toEqual({ ok: true, value: null });
  });
});
