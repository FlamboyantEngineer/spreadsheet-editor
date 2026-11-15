import type { WorkbookChangeContext } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
// A workbook session: one stored workbook, loaded through the host's workbook store, edited, and
// saved against the revision it was edited from. It needs no framework or DOM; the editor takes it
// as its `session` prop. Coverage is causal: a save covers the local transactions accepted before
// it started, and only its answer clears them (SE-SESSION-1.2).
import { copyWorkbook, type Workbook } from '../model/workbook.js';
import type {
  DraftsAdapter,
  HostError,
  LifecycleAdapter,
  SavedState,
  StoredWorkbook,
  WorkbookDraft,
  WorkbookStoreAdapter,
} from './adapters.js';
import { createTransactionReceiptTracker } from './transaction-receipts.js';

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

const DEFAULT_AUTOSAVE_DELAY = 1000;
const DEFAULT_DRAFT_DELAY = 300;
/** Autosaves retry an unavailable store, with backoff, for this long. */
const RETRY_WINDOW = 30_000;

function failure(error: unknown): HostError {
  return { kind: 'unavailable', message: error instanceof Error ? error.message : undefined };
}

class HostRequestError extends Error {
  constructor(readonly error: HostError) {
    super(error.message ?? `The workbook could not be saved (${error.kind}).`);
  }
}

/** Opens a stored workbook (SE-SESSION-1.1). */
export function createWorkbookSession(options: WorkbookSessionOptions): WorkbookSession {
  const { workbooks, workbookId, drafts } = options;
  const autosaveDelay = options.autosaveDelay ?? DEFAULT_AUTOSAVE_DELAY;
  const draftDelay = options.draftDelay ?? DEFAULT_DRAFT_DELAY;
  const now = options.now ?? (() => new Date().toISOString());
  const listeners = new Set<() => void>();
  const loading = new AbortController();
  const tracker = createTransactionReceiptTracker();

  let workbook: Workbook | null = null;
  let stored: SavedState | null = null;
  let status: WorkbookSessionStatus = {
    phase: 'loading',
    revision: null,
    dirty: false,
    automatic: false,
    error: null,
  };
  let offer: (WorkbookDraftOffer & { content: Workbook }) | null = null;
  let closed = false;
  /** The local revision of the latest accepted change, and the latest one a save answer covers. */
  let localRevision = 0;
  let savedThrough = 0;
  let inFlight: Promise<void> | null = null;
  let autosaveTimer: ReturnType<typeof setTimeout> | null = null;
  let draftTimer: ReturnType<typeof setTimeout> | null = null;
  let stopStore: (() => void) | undefined;
  let stopLifecycle: (() => void) | undefined;

  function emit() {
    for (const listener of [...listeners]) listener();
  }
  function setStatus(next: Partial<WorkbookSessionStatus>) {
    const dirty = localRevision > savedThrough;
    status = { ...status, dirty, ...next };
    emit();
  }
  function idlePhase(): WorkbookSessionPhase {
    return localRevision > savedThrough ? 'dirty' : 'saved';
  }
  function clearTimer(timer: ReturnType<typeof setTimeout> | null) {
    if (timer) clearTimeout(timer);
  }

  function scheduleAutosave() {
    clearTimer(autosaveTimer);
    autosaveTimer = null;
    if (
      closed ||
      status.phase === 'conflict' ||
      status.phase === 'error' ||
      !Number.isFinite(autosaveDelay)
    )
      return;
    autosaveTimer = setTimeout(() => {
      autosaveTimer = null;
      void run(true).catch(() => {});
    }, autosaveDelay);
  }
  function scheduleDraft() {
    if (!drafts || closed) return;
    clearTimer(draftTimer);
    draftTimer = setTimeout(() => {
      draftTimer = null;
      if (closed || !workbook || !stored || localRevision <= savedThrough) return;
      const draft: WorkbookDraft = {
        content: copyWorkbook(workbook),
        baseRevision: stored.revision,
        savedAt: now(),
      };
      void drafts.put(workbookId, draft).catch(() => {});
    }, draftDelay);
  }
  function clearDraft() {
    clearTimer(draftTimer);
    draftTimer = null;
    if (drafts) void drafts.clear(workbookId).catch(() => {});
  }

  async function saveOnce(automatic: boolean, baseRevision: string): Promise<void> {
    const covered = localRevision,
      content = copyWorkbook(requireDefined(workbook));
    const started = Date.now();
    let retry = 500;
    setStatus({ phase: 'saving', automatic, error: null });
    while (true) {
      let result;
      try {
        result = await workbooks.save(workbookId, { content, baseRevision });
      } catch (error) {
        result = { ok: false as const, error: failure(error) };
      }
      if (closed) throw new Error('The workbook session has closed.');
      if (result.ok) {
        stored = result.value;
        savedThrough = Math.max(savedThrough, covered);
        tracker.forgetThrough(covered);
        if (localRevision === savedThrough) clearDraft();
        setStatus({ phase: idlePhase(), revision: stored.revision, error: null });
        if (localRevision > savedThrough) scheduleAutosave();
        return;
      }
      const { error } = result;
      const retrying =
        automatic && error.kind === 'unavailable' && Date.now() - started < RETRY_WINDOW;
      if (!retrying) {
        if (error.kind === 'conflict') {
          if (error.current) stored = error.current;
          setStatus({ phase: 'conflict', error });
        } else setStatus({ phase: 'error', error });
        throw new HostRequestError(error);
      }
      await new Promise((resolve) => setTimeout(resolve, retry));
      retry = Math.min(4000, retry * 2);
      if (closed) throw new Error('The workbook session has closed.');
    }
  }

  /** Saves until an answer covers every change accepted before it was asked for. */
  function run(automatic: boolean, baseRevision?: string): Promise<void> {
    if (closed) return Promise.reject(new Error('The workbook session has closed.'));
    if (!workbook || !stored) return Promise.reject(new Error('The workbook is not loaded.'));
    if (inFlight) {
      const requested = localRevision;
      return inFlight.then(() => (savedThrough < requested ? run(automatic) : undefined));
    }
    if (
      (status.phase === 'conflict' && baseRevision === undefined) ||
      (automatic && status.phase === 'error')
    )
      return Promise.reject(new HostRequestError(status.error ?? { kind: 'conflict' }));
    if (localRevision <= savedThrough) return Promise.resolve();
    clearTimer(autosaveTimer);
    autosaveTimer = null;
    const operation = saveOnce(automatic, baseRevision ?? stored.revision).finally(() => {
      inFlight = null;
    });
    inFlight = operation;
    return operation;
  }

  async function load(): Promise<boolean> {
    let result;
    try {
      result = await workbooks.load(workbookId, loading.signal);
    } catch (error) {
      result = { ok: false as const, error: failure(error) };
    }
    if (closed) return false;
    if (!result.ok) {
      setStatus({ phase: 'unavailable', error: result.error });
      return false;
    }
    adopt(result.value);
    if (drafts) {
      const kept = await drafts.get(workbookId).catch(() => null);
      if (!closed && kept?.ok && kept.value) {
        offer = {
          content: kept.value.content,
          savedAt: kept.value.savedAt,
          baseRevision: kept.value.baseRevision,
          stale: kept.value.baseRevision !== result.value.revision,
        };
        emit();
      }
    }
    if (!closed && workbooks.subscribe)
      stopStore = workbooks.subscribe(workbookId, (state) => void reloadAfter(state));
    if (!closed && options.lifecycle)
      stopLifecycle = options.lifecycle.subscribe((event) => {
        if (event !== 'visible' && localRevision > savedThrough) void run(true).catch(() => {});
      });
    return !closed;
  }
  function adopt(value: StoredWorkbook) {
    workbook = value.content;
    stored = { revision: value.revision, sequence: value.sequence };
    setStatus({ phase: idlePhase(), revision: value.revision, error: null });
  }
  /** Another save landed: a session without unsaved changes shows it. */
  async function reloadAfter(state: SavedState) {
    if (
      closed ||
      !stored ||
      state.sequence <= stored.sequence ||
      state.revision === stored.revision
    )
      return;
    if (localRevision > savedThrough || inFlight) return;
    const result = await workbooks.load(workbookId).catch(() => null);
    if (closed || !result?.ok || localRevision > savedThrough || inFlight) return;
    if (result.value.sequence > (stored?.sequence ?? -1)) adopt(result.value);
  }

  const ready = load();

  return {
    workbookId,
    ready,
    workbook: () => workbook,
    status: () => status,
    draft: () =>
      offer
        ? { savedAt: offer.savedAt, baseRevision: offer.baseRevision, stale: offer.stale }
        : null,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    change(next, context) {
      if (closed) throw new Error('The workbook session has closed.');
      if (!stored) throw new Error('The workbook is not loaded.');
      localRevision++;
      tracker.record(context.transactionId, localRevision);
      workbook = next;
      setStatus(
        status.phase === 'saving' || status.phase === 'conflict' || status.phase === 'error'
          ? {}
          : { phase: 'dirty' },
      );
      scheduleDraft();
      scheduleAutosave();
    },
    acknowledge(transactionId) {
      const revision = tracker.acknowledge(transactionId);
      if (revision === undefined || revision <= savedThrough) return;
      savedThrough = revision;
      if (localRevision === savedThrough) clearDraft();
      setStatus(status.phase === 'saving' ? {} : { phase: idlePhase() });
    },
    save: () => run(false),
    async flush() {
      clearTimer(draftTimer);
      const requested = localRevision;
      while (savedThrough < requested) await run(false);
      if (inFlight) await inFlight;
    },
    async keepMine() {
      if (status.phase !== 'conflict' || !stored) return;
      await run(false, stored.revision);
    },
    async takeStored() {
      if (closed) return;
      const result = await workbooks.load(workbookId);
      if (closed) return;
      if (!result.ok) {
        setStatus({ phase: 'error', error: result.error });
        throw new HostRequestError(result.error);
      }
      savedThrough = localRevision;
      clearTimer(autosaveTimer);
      autosaveTimer = null;
      clearDraft();
      adopt(result.value);
    },
    restoreDraft() {
      if (!offer || closed || !stored) return false;
      workbook = copyWorkbook(offer.content);
      offer = null;
      localRevision++;
      tracker.record(`draft:${localRevision}`, localRevision);
      setStatus({ phase: 'dirty' });
      scheduleAutosave();
      return true;
    },
    discardDraft() {
      if (!offer) return;
      offer = null;
      clearDraft();
      emit();
    },
    close() {
      if (closed) return;
      closed = true;
      loading.abort();
      clearTimer(autosaveTimer);
      clearTimer(draftTimer);
      stopStore?.();
      stopLifecycle?.();
      listeners.clear();
    },
  };
}
