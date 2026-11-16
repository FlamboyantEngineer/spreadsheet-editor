import type { WorkbookChangeContext } from '../model/operations.js';
// A workbook session: one stored workbook, loaded through the host's workbook store, edited, and
// saved against the revision it was edited from. It needs no framework or DOM; the editor takes it
// as its `session` prop. Coverage is causal: a save covers the local transactions accepted before
// it started, and only its answer clears them (SE-SESSION-1.2).
import type { Workbook } from '../model/workbook.js';
import type {
  DraftsAdapter,
  HostError,
  LifecycleAdapter,
  WorkbookStoreAdapter,
} from './adapters.js';
import { StoredWorkbookSession } from './stored-workbook-session.js';

/** Where a session is: loading, failed to load, or open with or without unsaved changes. */
export type WorkbookSessionPhase =
  'loading' | 'unavailable' | 'saved' | 'dirty' | 'saving' | 'error' | 'conflict';

export interface WorkbookSessionStatus {
  readonly phase: WorkbookSessionPhase;
  /** The stored revision the workbook was last loaded or saved at; null until loaded. */
  readonly revision: string | null;
  /** Whether the session holds changes no answered save covers. */
  readonly dirty: boolean;
  /** Whether the last save was started by autosave rather than by the reader. */
  readonly automatic: boolean;
  /** Why the last load or save failed, while it is the latest outcome. */
  readonly error: HostError | null;
}

export interface WorkbookSessionOptions {
  workbooks: WorkbookStoreAdapter;
  workbookId: string;
  /** Keeps unsaved content across a reload, when given. */
  drafts?: DraftsAdapter | undefined;
  /** Saves when the page is hidden or left, when given. */
  lifecycle?: LifecycleAdapter | undefined;
  /** How long after the last change the session saves, in milliseconds. Defaults to 1000. */
  autosaveDelay?: number | undefined;
  /** How long after the last change the draft is kept, in milliseconds. Defaults to 300. */
  draftDelay?: number | undefined;
  /** The time drafts are recorded at. */
  now?: (() => string) | undefined;
}

/** An earlier visit's unsaved content, offered once the workbook loads. */
export interface WorkbookDraftOffer {
  readonly savedAt: string;
  readonly baseRevision: string;
  /** Whether the stored workbook moved on since the draft was edited from it. */
  readonly stale: boolean;
}

export interface WorkbookSession {
  readonly workbookId: string;
  /** Resolves true once loaded, false when loading failed or the session closed first. */
  readonly ready: Promise<boolean>;
  /** The workbook as the reader has it; null until loaded. */
  workbook(): Workbook | null;
  status(): WorkbookSessionStatus;
  draft(): WorkbookDraftOffer | null;
  /** Calls `listener` whenever the workbook, the status or the draft offer changes. */
  subscribe(listener: () => void): () => void;
  /** Records an accepted local change: the editor's `onChange`. */
  change(next: Workbook, context: WorkbookChangeContext): void;
  /**
   * Marks a local transaction as kept by the host by other means, such as an operation feed: it
   * and every earlier one count as saved. An unknown or already covered id changes nothing.
   */
  acknowledge(transactionId: string): void;
  /** Saves now, or joins the save in flight; resolves once saved or rejects with the failure. */
  save(): Promise<void>;
  /**
   * Saves what is pending and waits until every change accepted before the call is saved, or
   * rejects; a change made meanwhile stays unsaved.
   */
  flush(): Promise<void>;
  /** After a `conflict`, saves the reader's workbook over the stored one. */
  keepMine(): Promise<void>;
  /** After a `conflict`, or to drop unsaved changes, takes the stored workbook. */
  takeStored(): Promise<void>;
  /** Shows the offered draft as unsaved changes to the stored workbook. */
  restoreDraft(): boolean;
  /** Removes the offered draft. */
  discardDraft(): void;
  /** Closes the session: cancels timers and listeners, and keeps no further drafts. */
  close(): void;
}

/** Opens a stored workbook (SE-SESSION-1.1). */
export function createWorkbookSession(options: WorkbookSessionOptions): WorkbookSession {
  return new StoredWorkbookSession(options);
}
