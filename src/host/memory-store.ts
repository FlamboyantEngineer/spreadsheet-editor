import { requireDefined } from '../model/require-defined.js';
import { assertWorkbook, type Workbook } from '../model/workbook.js';
import type { HostError, SavedState, StoredWorkbook, WorkbookVersion } from './adapters.js';

/** Requests the memory controls can make fail. */
export type MemoryRequest = 'load' | 'save' | 'versions' | 'drafts';

interface Entry {
  versions: (WorkbookVersion & { content: Workbook })[];
  saved: Set<(state: SavedState) => void>;
  changed: Set<() => void>;
}

export const notFound = (what: string): { ok: false; error: HostError } => ({
  ok: false,
  error: { kind: 'not-found', message: `No ${what}.` },
});

export function copy(workbook: Workbook): Workbook {
  const next = structuredClone(workbook);
  assertWorkbook(next);
  return next;
}

export const stateOf = ({ revision, sequence }: SavedState): SavedState => ({ revision, sequence });
export const versionOf = (version: WorkbookVersion): WorkbookVersion => ({
  revision: version.revision,
  sequence: version.sequence,
  savedAt: version.savedAt,
  author: version.author,
  label: version.label,
});

/** The stored workbooks, their versions and listeners, and the failures queued for requests. */
export function createMemoryStore(
  { workbooks, now }: { workbooks: Readonly<Record<string, Workbook>>; now: () => string },
  author: string,
) {
  const failures = new Map<MemoryRequest, HostError[]>();
  const entries = new Map<string, Entry>();
  const latest = (entry: Entry) => requireDefined(entry.versions[0]);
  const record = (entry: Entry, content: Workbook, label: string | null) => {
    const sequence = (entry.versions[0]?.sequence ?? 0) + 1;
    const version = {
      revision: `r${sequence}`,
      sequence,
      savedAt: now(),
      author,
      label,
      content: copy(content),
    };
    entry.versions.unshift(version);
    entry.changed.forEach((onChange) => onChange());
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
  for (const [id, workbook] of Object.entries(workbooks)) record(entryOf(id), workbook, null);
  return {
    failures,
    latest,
    entryOf,
    stored: (id: string, entry: Entry): StoredWorkbook => ({
      id,
      ...stateOf(latest(entry)),
      content: copy(latest(entry).content),
    }),
    saveNew(entry: Entry, content: Workbook, label: string | null) {
      const version = record(entry, content, label);
      entry.saved.forEach((onSaved) => onSaved(stateOf(version)));
      return version;
    },
    failed(request: MemoryRequest): { ok: false; error: HostError } | null {
      const error = failures.get(request)?.shift();
      return error ? { ok: false, error: { ...error } } : null;
    },
    existing(id: string) {
      const entry = entries.get(id);
      return entry?.versions.length ? entry : undefined;
    },
  };
}
export type MemoryStore = ReturnType<typeof createMemoryStore>;
