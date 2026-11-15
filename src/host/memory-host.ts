import { assertWorkbook, type Workbook } from '../model/workbook.js';
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
  StoredWorkbook,
  VersionStoreAdapter,
  WorkbookDraft,
  WorkbookStoreAdapter,
  WorkbookVersion,
} from './adapters.js';
import { FULL_CAPABILITIES, type EditorCapabilities } from './capabilities.js';
import { createMemoryPresenceRoom, type MemoryPresenceRoom } from './memory-presence.js';

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

/** Requests the memory controls can make fail. */
export type MemoryRequest = 'load' | 'save' | 'versions' | 'drafts';

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

interface Entry {
  versions: (WorkbookVersion & { content: Workbook })[];
  saved: Set<(state: SavedState) => void>;
  changed: Set<() => void>;
}

const notFound = (what: string): { ok: false; error: HostError } => ({
  ok: false,
  error: { kind: 'not-found', message: `No ${what}.` },
});

function copy(workbook: Workbook): Workbook {
  const next = structuredClone(workbook);
  assertWorkbook(next);
  return next;
}

/** A complete host kept in memory, for labs, tests and standalone pages (SE-ADAPTERS-3). */
export function createMemoryHost(options: MemoryHostOptions = {}): MemoryHost {
  const reader = { ...(options.reader ?? { id: 'reader', name: 'Reader' }) };
  const people = [reader, ...(options.people ?? []).filter((person) => person.id !== reader.id)];
  const now = options.now ?? (() => new Date().toISOString());
  let capabilities = { ...(options.capabilities ?? FULL_CAPABILITIES) };
  const capabilityListeners = new Set<(capabilities: EditorCapabilities) => void>();
  const lifecycleListeners = new Set<(event: LifecycleEvent) => void>();
  const failures = new Map<MemoryRequest, HostError[]>();
  const drafts = new Map<string, WorkbookDraft>();
  const entries = new Map<string, Entry>();

  const record = (entry: Entry, content: Workbook, label: string | null) => {
    const sequence = (entry.versions[0]?.sequence ?? 0) + 1;
    const version = {
      revision: `r${sequence}`,
      sequence,
      savedAt: now(),
      author: reader.id,
      label,
      content: copy(content),
    };
    entry.versions.unshift(version);
    entry.changed.forEach((onChange) => onChange());
    return version;
  };
  const latest = (entry: Entry) => entry.versions[0]!;
  const stateOf = ({ revision, sequence }: SavedState): SavedState => ({ revision, sequence });
  const versionOf = (version: WorkbookVersion): WorkbookVersion => ({
    revision: version.revision,
    sequence: version.sequence,
    savedAt: version.savedAt,
    author: version.author,
    label: version.label,
  });
  const stored = (id: string, entry: Entry): StoredWorkbook => ({
    id,
    ...stateOf(latest(entry)),
    content: copy(latest(entry).content),
  });
  const saveNew = (entry: Entry, content: Workbook, label: string | null) => {
    const version = record(entry, content, label);
    entry.saved.forEach((onSaved) => onSaved(stateOf(version)));
    return version;
  };
  const entryOf = (id: string) => {
    let entry = entries.get(id);
    if (!entry) {
      entry = { versions: [], saved: new Set(), changed: new Set() };
      entries.set(id, entry);
    }
    return entry;
  };
  for (const [id, workbook] of Object.entries(options.workbooks ?? {}))
    record(entryOf(id), workbook, null);
  const failed = (request: MemoryRequest): { ok: false; error: HostError } | null => {
    const error = failures.get(request)?.shift();
    return error ? { ok: false, error: { ...error } } : null;
  };
  const existing = (id: string) => {
    const entry = entries.get(id);
    return entry?.versions.length ? entry : undefined;
  };
  const presence = (options.presence ?? createMemoryPresenceRoom()).join(reader);

  return {
    identity: {
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
    },
    permissions: {
      capabilities: async () => ({ ...capabilities }),
      subscribe(_workbookId, onCapabilities) {
        capabilityListeners.add(onCapabilities);
        return () => capabilityListeners.delete(onCapabilities);
      },
    },
    workbooks: {
      async load(id) {
        const failure = failed('load');
        if (failure) return failure;
        const entry = existing(id);
        return entry ? { ok: true, value: stored(id, entry) } : notFound(`workbook ${id}`);
      },
      async save(id, save): Promise<HostResult<SavedState>> {
        const failure = failed('save');
        if (failure) return failure;
        const entry = existing(id);
        if (!entry) return notFound(`workbook ${id}`);
        if (save.baseRevision !== latest(entry).revision)
          return {
            ok: false,
            error: {
              kind: 'conflict',
              message: 'The workbook changed.',
              current: stored(id, entry),
            },
          };
        return { ok: true, value: stateOf(saveNew(entry, save.content, null)) };
      },
      subscribe(id, onSaved) {
        const entry = entryOf(id);
        entry.saved.add(onSaved);
        return () => entry.saved.delete(onSaved);
      },
    },
    versions: {
      async list(id) {
        const failure = failed('versions');
        if (failure) return failure;
        const entry = existing(id);
        return entry
          ? { ok: true, value: entry.versions.map(versionOf) }
          : notFound(`workbook ${id}`);
      },
      async load(id, revision) {
        const failure = failed('versions');
        if (failure) return failure;
        const version = existing(id)?.versions.find((item) => item.revision === revision);
        return version
          ? { ok: true, value: copy(version.content) }
          : notFound(`version ${revision}`);
      },
      async create(id, { content, label }) {
        const entry = existing(id);
        if (!entry) return notFound(`workbook ${id}`);
        return { ok: true, value: versionOf(saveNew(entry, content, label)) };
      },
      async rename(id, revision, label) {
        const entry = existing(id);
        const version = entry?.versions.find((item) => item.revision === revision);
        if (!entry || !version) return notFound(`version ${revision}`);
        version.label = label;
        entry.changed.forEach((onChange) => onChange());
        return { ok: true, value: versionOf(version) };
      },
      async restore(id, revision) {
        const entry = existing(id);
        const version = entry?.versions.find((item) => item.revision === revision);
        if (!entry || !version) return notFound(`version ${revision}`);
        return { ok: true, value: stateOf(saveNew(entry, version.content, null)) };
      },
      subscribe(id, onChange) {
        const entry = entryOf(id);
        entry.changed.add(onChange);
        return () => entry.changed.delete(onChange);
      },
    },
    presence,
    drafts: {
      async get(id) {
        const failure = failed('drafts');
        if (failure) return failure;
        const draft = drafts.get(id);
        return { ok: true, value: draft ? structuredClone(draft) : null };
      },
      async put(id, draft) {
        const failure = failed('drafts');
        if (failure) return failure;
        drafts.set(id, structuredClone(draft));
        return { ok: true, value: null };
      },
      async clear(id) {
        drafts.delete(id);
        return { ok: true, value: null };
      },
    },
    lifecycle: {
      subscribe(onEvent) {
        lifecycleListeners.add(onEvent);
        return () => lifecycleListeners.delete(onEvent);
      },
    },
    memory: {
      failNext(request, error) {
        failures.set(request, [...(failures.get(request) ?? []), { ...error }]);
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
