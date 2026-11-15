// The host contract: everything the editor reaches outside itself, as adapters the host passes in.
// The names and shapes follow the document editor's host adapters, so a host that embeds both
// editors can pass one identity, one permissions adapter and one versions store to each.

import type { Workbook } from '../model/workbook.js';
import type { EditorCapabilities } from './capabilities.js';
import type { PresenceAdapter } from './presence-adapters.js';

export type * from './presence-adapters.js';

/** A failure the host reports for a request it could not complete. */
export interface HostError {
  kind: 'not-found' | 'denied' | 'conflict' | 'unavailable';
  /** Text the editor may show the reader. */
  message?: string | undefined;
  /**
   * For a workbook save refused as `conflict`, the workbook as it is stored now. Without it, the
   * editor loads the workbook again.
   */
  current?: StoredWorkbook | undefined;
}

/** The outcome of a host request: a value, or the reason there is none. */
export type HostResult<T> = { ok: true; value: T } | { ok: false; error: HostError };

/** A person as the editor shows them: in comments, presence, provenance and history. */
export interface EditorPerson {
  id: string;
  name: string;
  /** Image URL for the person's avatar. */
  avatarUrl?: string | undefined;
}

export interface IdentityAdapter {
  /** The reader using this editor instance. */
  currentPerson(): EditorPerson;
  /** People whose names match `query`, for mentions and pickers. */
  findPeople(query: string, signal?: AbortSignal): Promise<readonly EditorPerson[]>;
  /** People by id, for names in history, comments and provenance. Unknown ids are left out. */
  resolvePeople(ids: readonly string[]): Promise<readonly EditorPerson[]>;
}

export interface PermissionsAdapter {
  /** What the reader may do with the workbook. */
  capabilities(workbookId: string): Promise<EditorCapabilities>;
  /** Calls `onCapabilities` whenever what the reader may do changes, until unsubscribed. */
  subscribe?(
    workbookId: string,
    onCapabilities: (capabilities: EditorCapabilities) => void,
  ): () => void;
}

/**
 * A saved state of a stored workbook. The revision is the host's own identifier, which the editor
 * only compares; the sequence orders saved states, and the editor sorts by it alone.
 */
export interface SavedState {
  revision: string;
  sequence: number;
}

/** A stored workbook and the saved state it was read at. */
export interface StoredWorkbook extends SavedState {
  id: string;
  content: Workbook;
}

export interface WorkbookSave {
  content: Workbook;
  /** The revision this content was edited from. Any other stored revision fails as `conflict`. */
  baseRevision: string;
}

export interface WorkbookStoreAdapter {
  load(workbookId: string, signal?: AbortSignal): Promise<HostResult<StoredWorkbook>>;
  /** Saves new content and returns the new saved state. */
  save(workbookId: string, save: WorkbookSave): Promise<HostResult<SavedState>>;
  /**
   * Calls `onSaved` with the workbook's saved state whenever anyone saves it, until the returned
   * function is called. The editor loads the content it lacks.
   */
  subscribe?(workbookId: string, onSaved: (state: SavedState) => void): () => void;
}

/** A saved version of a workbook. */
export interface WorkbookVersion extends SavedState {
  /** When it was saved: an ISO 8601 date and time. */
  savedAt: string;
  /** The id of the person who saved it, when the host knows. */
  author: string | null;
  /** The version's name; null for a version saved automatically. */
  label: string | null;
}

/** A workbook's saved versions, for its history. */
export interface VersionStoreAdapter {
  /** The workbook's versions, newest first. */
  list(workbookId: string, signal?: AbortSignal): Promise<HostResult<WorkbookVersion[]>>;
  /** A version's content, by its revision. */
  load(workbookId: string, revision: string, signal?: AbortSignal): Promise<HostResult<Workbook>>;
  /** Saves `content` as a new version, named by `label` or null, and returns it. */
  create?(
    workbookId: string,
    version: { content: Workbook; label: string | null },
  ): Promise<HostResult<WorkbookVersion>>;
  /** Names a version, or takes its name away with null, and returns it. */
  rename?(
    workbookId: string,
    revision: string,
    label: string | null,
  ): Promise<HostResult<WorkbookVersion>>;
  /** Makes a version's content the workbook's current content, as a new saved state. */
  restore?(workbookId: string, revision: string): Promise<HostResult<SavedState>>;
  /** Calls `onChange` whenever the workbook's versions change, until unsubscribed. */
  subscribe?(workbookId: string, onChange: () => void): () => void;
}

/** Unsaved content kept so a reload or a crash does not lose it, and the revision it was edited from. */
export interface WorkbookDraft {
  content: Workbook;
  baseRevision: string;
  /** When it was kept: an ISO 8601 date and time. */
  savedAt: string;
}

/** Where the editor keeps a workbook's draft, such as the browser's indexed storage. */
export interface DraftsAdapter {
  get(workbookId: string): Promise<HostResult<WorkbookDraft | null>>;
  /** Keeps a draft in place of the workbook's earlier one. */
  put(workbookId: string, draft: WorkbookDraft): Promise<HostResult<null>>;
  clear(workbookId: string): Promise<HostResult<null>>;
}

/** A change in the page's lifecycle that the editor saves or pauses on. */
export type LifecycleEvent = 'hidden' | 'visible' | 'leaving';

/** The page's lifecycle, such as the document's visibility and its `pagehide` event. */
export interface LifecycleAdapter {
  subscribe(onEvent: (event: LifecycleEvent) => void): () => void;
}

/**
 * Everything the editor needs from its host. Identity and permissions are required; the rest are
 * optional and leave their features off when absent.
 */
export interface SpreadsheetHost {
  identity: IdentityAdapter;
  permissions: PermissionsAdapter;
  workbooks?: WorkbookStoreAdapter | undefined;
  versions?: VersionStoreAdapter | undefined;
  presence?: PresenceAdapter | undefined;
  drafts?: DraftsAdapter | undefined;
  lifecycle?: LifecycleAdapter | undefined;
}
