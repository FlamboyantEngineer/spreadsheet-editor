import { describe, expect, it } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetHost } from './adapters.js';
import { READ_ONLY_CAPABILITIES } from './capabilities.js';
import { createMemoryHost } from './memory-host.js';
import { createMemoryPresenceRoom } from './memory-presence.js';
import type { SpreadsheetPeer } from './presence-adapters.js';

const at = '2026-11-15T12:00:00.000Z';

describe('SE-ADAPTERS-1: host contract', () => {
  it('SE-ADAPTERS-1.1: a host needs only identity and permissions', () => {
    const memory = createMemoryHost();
    const host: SpreadsheetHost = { identity: memory.identity, permissions: memory.permissions };
    expect(host.workbooks).toBeUndefined();
  });

  it('SE-ADAPTERS-1.2: expected failures resolve to results', async () => {
    const host = createMemoryHost();
    expect(await host.workbooks.load('missing')).toEqual({
      ok: false,
      error: { kind: 'not-found', message: 'No workbook missing.' },
    });
    host.memory.failNext('load', { kind: 'unavailable' });
    host.memory.failNext('load', { kind: 'denied' });
    expect(await host.workbooks.load('a')).toEqual({ ok: false, error: { kind: 'unavailable' } });
    expect(await host.workbooks.load('a')).toEqual({ ok: false, error: { kind: 'denied' } });
  });

  it('SE-ADAPTERS-1.3: a save names its base revision and a stale one is a conflict with the stored workbook', async () => {
    const host = createMemoryHost({ workbooks: { a: createWorkbook([['1']]) }, now: () => at });
    const loaded = await host.workbooks.load('a');
    if (!loaded.ok) throw new Error('not loaded');
    expect(loaded.value).toMatchObject({ id: 'a', revision: 'r1', sequence: 1 });
    const saved: unknown[] = [];
    requireDefined(host.workbooks.subscribe)('a', (state) => saved.push(state));
    const first = await host.workbooks.save('a', {
      content: createWorkbook([['2']]),
      baseRevision: loaded.value.revision,
    });
    expect(first).toEqual({ ok: true, value: { revision: 'r2', sequence: 2 } });
    expect(saved).toEqual([{ revision: 'r2', sequence: 2 }]);
    const stale = await host.workbooks.save('a', {
      content: createWorkbook([['3']]),
      baseRevision: 'r1',
    });
    expect(stale.ok).toBe(false);
    if (stale.ok) return;
    expect(stale.error.kind).toBe('conflict');
    expect(stale.error.current?.revision).toBe('r2');
    expect(requireDefined(stale.error.current?.content.worksheets[0]).data).toEqual([['2']]);
  });
});

describe('SE-ADAPTERS-2: capabilities', () => {
  it('SE-ADAPTERS-2.2: the host follows changes to what the reader may do', async () => {
    const host = createMemoryHost();
    expect((await host.permissions.capabilities('a')).edit).toBe(true);
    const seen: boolean[] = [];
    const stop = requireDefined(host.permissions.subscribe)('a', (next) => seen.push(next.edit));
    host.memory.setCapabilities(READ_ONLY_CAPABILITIES);
    stop();
    host.memory.setCapabilities({ ...READ_ONLY_CAPABILITIES, edit: true });
    expect(seen).toEqual([false]);
    expect((await host.permissions.capabilities('a')).edit).toBe(true);
  });
});

describe('SE-ADAPTERS-3: in-memory host', () => {
  it('SE-ADAPTERS-3.1: defaults to a reader named Reader with every capability', async () => {
    const host = createMemoryHost();
    expect(host.identity.currentPerson()).toEqual({ id: 'reader', name: 'Reader' });
    expect(await host.permissions.capabilities('any')).toEqual({
      edit: true,
      suggest: true,
      comment: true,
      decide: true,
      saveVersion: true,
      history: true,
    });
  });

  it('SE-ADAPTERS-3.2: returns copies and shares nothing between hosts', async () => {
    const workbook = createWorkbook([['1']]);
    const host = createMemoryHost({ workbooks: { a: workbook } });
    requireDefined(requireDefined(workbook.worksheets[0]).data[0])[0] = 'changed';
    const loaded = await host.workbooks.load('a');
    if (!loaded.ok) throw new Error('not loaded');
    requireDefined(requireDefined(loaded.value.content.worksheets[0]).data[0])[0] = 'mutated';
    const again = await host.workbooks.load('a');
    expect(again.ok && requireDefined(again.value.content.worksheets[0]).data).toEqual([['1']]);
    expect((await createMemoryHost().workbooks.load('a')).ok).toBe(false);
  });

  it('SE-ADAPTERS-3.3: finds people by name and resolves known ids', async () => {
    const host = createMemoryHost({
      reader: { id: 'r', name: 'Robin' },
      people: [
        { id: 'm', name: 'Morgan' },
        { id: 'a', name: 'Alex' },
      ],
    });
    expect((await host.identity.findPeople('  MOR ')).map((person) => person.id)).toEqual(['m']);
    expect((await host.identity.findPeople('')).map((person) => person.id)).toEqual([
      'r',
      'm',
      'a',
    ]);
    expect(await host.identity.resolvePeople(['a', 'gone', 'r'])).toEqual([
      { id: 'a', name: 'Alex' },
      { id: 'r', name: 'Robin' },
    ]);
  });

  it('SE-ADAPTERS-3.4: keeps every saved state as a version, newest first, and restores one', async () => {
    const host = createMemoryHost({ workbooks: { a: createWorkbook([['1']]) }, now: () => at });
    await host.workbooks.save('a', { content: createWorkbook([['2']]), baseRevision: 'r1' });
    const changes: string[] = [];
    requireDefined(host.versions.subscribe)('a', () => changes.push('changed'));
    const named = await requireDefined(host.versions.create)('a', {
      content: createWorkbook([['3']]),
      label: 'Named',
    });
    expect(named).toEqual({
      ok: true,
      value: { revision: 'r3', sequence: 3, savedAt: at, author: 'reader', label: 'Named' },
    });
    await requireDefined(host.versions.rename)('a', 'r1', 'First');
    const list = await host.versions.list('a');
    expect(list.ok && list.value.map((version) => [version.revision, version.label])).toEqual([
      ['r3', 'Named'],
      ['r2', null],
      ['r1', 'First'],
    ]);
    const restored = await requireDefined(host.versions.restore)('a', 'r1');
    expect(restored).toEqual({ ok: true, value: { revision: 'r4', sequence: 4 } });
    const current = await host.workbooks.load('a');
    expect(current.ok && requireDefined(current.value.content.worksheets[0]).data).toEqual([['1']]);
    expect(await host.versions.load('a', 'r9')).toMatchObject({ ok: false });
    expect(changes).toEqual(['changed', 'changed', 'changed']);
  });

  it('SE-ADAPTERS-3.5: keeps one draft per workbook', async () => {
    const host = createMemoryHost();
    const draft = { content: createWorkbook([['x']]), baseRevision: 'r1', savedAt: at };
    expect(await host.drafts.get('a')).toEqual({ ok: true, value: null });
    await host.drafts.put('a', draft);
    requireDefined(requireDefined(draft.content.worksheets[0]).data[0])[0] = 'changed';
    const kept = await host.drafts.get('a');
    expect(kept.ok && requireDefined(kept.value?.content.worksheets[0]).data).toEqual([['x']]);
    await host.drafts.clear('a');
    expect(await host.drafts.get('a')).toEqual({ ok: true, value: null });
  });

  it('SE-ADAPTERS-3.6: tells lifecycle subscribers of a change', () => {
    const host = createMemoryHost();
    const events: string[] = [];
    const stop = host.lifecycle.subscribe((event) => events.push(event));
    host.memory.emitLifecycle('hidden');
    stop();
    host.memory.emitLifecycle('visible');
    expect(events).toEqual(['hidden']);
  });
});

describe('SE-ADAPTERS-4: presence', () => {
  it('SE-ADAPTERS-4.2: people in one room see each other, not themselves, by name', () => {
    const room = createMemoryPresenceRoom();
    const robin = createMemoryHost({ reader: { id: 'r', name: 'Robin' }, presence: room });
    const alex = createMemoryHost({
      reader: { id: 'a', name: 'Alex', avatarUrl: 'https://example.com/a.png' },
      presence: room,
    });
    const seen: (readonly SpreadsheetPeer[])[] = [];
    const stop = requireDefined(robin.presence).subscribe('book', (peers) => seen.push(peers));
    requireDefined(robin.presence).publish('book', {
      at: { sheetId: 's', range: 'A1' },
      status: 'active',
    });
    requireDefined(alex.presence).publish('book', {
      at: { sheetId: 's', range: 'B2:C3' },
      editLock: { sheetId: 's', range: 'B2' },
      status: 'idle',
    });
    requireDefined(alex.presence).publish('other', { at: null, status: 'active' });
    requireDefined(alex.presence).publish('book', null);
    stop();
    expect(seen).toEqual([
      [],
      [],
      [
        {
          id: 'a',
          name: 'Alex',
          avatar: 'https://example.com/a.png',
          at: { sheetId: 's', range: 'B2:C3' },
          editLock: { sheetId: 's', range: 'B2' },
          status: 'idle',
        },
      ],
      [],
    ]);
  });
});
