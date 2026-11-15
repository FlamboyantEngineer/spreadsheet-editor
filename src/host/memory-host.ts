import type { Workbook } from '../model/workbook.js';
import type {
  DraftsAdapter,
  EditorPerson,
  HostError,
  HostResult,
  IdentityAdapter,
  LifecycleAdapter,
  LifecycleEvent,
  PermissionsAdapter,
  SavedState,
  SpreadsheetHost,
  VersionStoreAdapter,
  WorkbookDraft,
  WorkbookStoreAdapter,
} from './adapters.js';
import { type EditorCapabilities, FULL_CAPABILITIES } from './capabilities.js';
import { createMemoryPresenceRoom, type MemoryPresenceRoom } from './memory-presence.js';
import {
  copy,
  createMemoryStore,
  type MemoryRequest,
  type MemoryStore,
  notFound,
  stateOf,
  versionOf,
} from './memory-store.js';

export type { MemoryRequest } from './memory-store.js';

export interface MemoryHostOptions {
  /** The reader; Reader, with id `reader`, by default. */
  reader?: EditorPerson | undefined;
  /** Other people the identity adapter knows. */
  people?: readonly EditorPerson[] | undefined;
  /** What the reader may do with every workbook; everything by default. */
  capabilities?: EditorCapabilities | undefined;
  /** The stored workbooks, by id, each at its first saved state. */
  workbooks?: Readonly<Record<string, Workbook>> | undefined;
  /** The time saves and versions are recorded at. */
  now?: (() => string) | undefined;
  /** A presence room shared with other memory hosts; a room of its own by default. */
  presence?: MemoryPresenceRoom | undefined;
}

/** Controls for labs and tests. */
export interface MemoryHostControls {
  /** Fails the next request of a kind with this error. */
  failNext(request: MemoryRequest, error: HostError): void;
  /** Gives the reader other capabilities and tells the subscribers. */
  setCapabilities(capabilities: EditorCapabilities): void;
  /** Tells the lifecycle subscribers of a lifecycle change. */
  emitLifecycle(event: LifecycleEvent): void;
}

export interface MemoryHost extends SpreadsheetHost {
  identity: IdentityAdapter;
  permissions: PermissionsAdapter;
  workbooks: WorkbookStoreAdapter;
  versions: VersionStoreAdapter;
  drafts: DraftsAdapter;
  lifecycle: LifecycleAdapter;
  memory: MemoryHostControls;
}

function identityAdapter(reader: EditorPerson, people: readonly EditorPerson[]): IdentityAdapter {
  return {
    currentPerson: () => ({ ...reader }),
    async findPeople(query) {
      const needle = query.trim().toLowerCase();
      return people
        .filter((person) => person.name.toLowerCase().includes(needle))
        .map((person) => ({ ...person }));
    },
    async resolvePeople(ids) {
      return ids.flatMap((id) => {
        const person = people.find((candidate) => candidate.id === id);
        return person ? [{ ...person }] : [];
      });
    },
  };
}

function workbookAdapter(store: MemoryStore): WorkbookStoreAdapter {
  return {
    async load(id) {
      const failure = store.failed('load');
      if (failure) return failure;
      const entry = store.existing(id);
      return entry ? { ok: true, value: store.stored(id, entry) } : notFound(`workbook ${id}`);
    },
    async save(id, save): Promise<HostResult<SavedState>> {
      const failure = store.failed('save');
      if (failure) return failure;
      const entry = store.existing(id);
      if (!entry) return notFound(`workbook ${id}`);
      if (save.baseRevision !== store.latest(entry).revision)
        return {
          ok: false,
          error: {
            kind: 'conflict',
            message: 'The workbook changed.',
            current: store.stored(id, entry),
          },
        };
      return { ok: true, value: stateOf(store.saveNew(entry, save.content, null)) };
    },
    subscribe(id, onSaved) {
      const entry = store.entryOf(id);
      entry.saved.add(onSaved);
      return () => entry.saved.delete(onSaved);
    },
  };
}

function versionAdapter(store: MemoryStore): VersionStoreAdapter {
  const versionIn = (id: string, revision: string) => {
    const entry = store.existing(id);
    const version = entry?.versions.find((item) => item.revision === revision);
    return entry && version ? { entry, version } : undefined;
  };
  return {
    async list(id) {
      const failure = store.failed('versions');
      if (failure) return failure;
      const entry = store.existing(id);
      return entry
        ? { ok: true, value: entry.versions.map(versionOf) }
        : notFound(`workbook ${id}`);
    },
    async load(id, revision) {
      const failure = store.failed('versions');
      if (failure) return failure;
      const version = store.existing(id)?.versions.find((item) => item.revision === revision);
      return version ? { ok: true, value: copy(version.content) } : notFound(`version ${revision}`);
    },
    async create(id, { content, label }) {
      const entry = store.existing(id);
      if (!entry) return notFound(`workbook ${id}`);
      return { ok: true, value: versionOf(store.saveNew(entry, content, label)) };
    },
    async rename(id, revision, label) {
      const found = versionIn(id, revision);
      if (!found) return notFound(`version ${revision}`);
      found.version.label = label;
      found.entry.changed.forEach((onChange) => onChange());
      return { ok: true, value: versionOf(found.version) };
    },
    async restore(id, revision) {
      const found = versionIn(id, revision);
      if (!found) return notFound(`version ${revision}`);
      return { ok: true, value: stateOf(store.saveNew(found.entry, found.version.content, null)) };
    },
    subscribe(id, onChange) {
      const entry = store.entryOf(id);
      entry.changed.add(onChange);
      return () => entry.changed.delete(onChange);
    },
  };
}

function draftsAdapter(store: MemoryStore): DraftsAdapter {
  const drafts = new Map<string, WorkbookDraft>();
  return {
    async get(id) {
      const failure = store.failed('drafts');
      if (failure) return failure;
      const draft = drafts.get(id);
      return { ok: true, value: draft ? structuredClone(draft) : null };
    },
    async put(id, draft) {
      const failure = store.failed('drafts');
      if (failure) return failure;
      drafts.set(id, structuredClone(draft));
      return { ok: true, value: null };
    },
    async clear(id) {
      drafts.delete(id);
      return { ok: true, value: null };
    },
  };
}

/** A complete host kept in memory, for labs, tests and standalone pages (SE-ADAPTERS-3). */
export function createMemoryHost(options: MemoryHostOptions = {}): MemoryHost {
  const reader = { ...(options.reader ?? { id: 'reader', name: 'Reader' }) };
  const people = [reader, ...(options.people ?? []).filter((person) => person.id !== reader.id)];
  const now = options.now ?? (() => new Date().toISOString());
  let capabilities = { ...(options.capabilities ?? FULL_CAPABILITIES) };
  const capabilityListeners = new Set<(capabilities: EditorCapabilities) => void>();
  const lifecycleListeners = new Set<(event: LifecycleEvent) => void>();
  const store = createMemoryStore({ workbooks: options.workbooks ?? {}, now }, reader.id);
  const presence = (options.presence ?? createMemoryPresenceRoom()).join(reader);
  return {
    identity: identityAdapter(reader, people),
    permissions: {
      capabilities: async () => ({ ...capabilities }),
      subscribe(_workbookId, onCapabilities) {
        capabilityListeners.add(onCapabilities);
        return () => capabilityListeners.delete(onCapabilities);
      },
    },
    workbooks: workbookAdapter(store),
    versions: versionAdapter(store),
    presence,
    drafts: draftsAdapter(store),
    lifecycle: {
      subscribe(onEvent) {
        lifecycleListeners.add(onEvent);
        return () => lifecycleListeners.delete(onEvent);
      },
    },
    memory: {
      failNext(request, error) {
        store.failures.set(request, [...(store.failures.get(request) ?? []), { ...error }]);
      },
      setCapabilities(next) {
        capabilities = { ...next };
        capabilityListeners.forEach((onCapabilities) => onCapabilities({ ...capabilities }));
      },
      emitLifecycle(event) {
        lifecycleListeners.forEach((onEvent) => onEvent(event));
      },
    },
  };
}
