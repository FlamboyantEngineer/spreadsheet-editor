import type { WorkbookChangeContext } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
import { copyWorkbook, type Workbook } from '../model/workbook.js';
import type { HostError, SavedState, StoredWorkbook, WorkbookDraft } from './adapters.js';
import {
  answer,
  closedError,
  DEFAULT_AUTOSAVE_DELAY,
  DEFAULT_DRAFT_DELAY,
  HostRequestError,
  RETRY_WINDOW,
  type Timer,
  wait,
} from './session-requests.js';
import { createTransactionReceiptTracker } from './transaction-receipts.js';
import type {
  WorkbookDraftOffer,
  WorkbookSession,
  WorkbookSessionOptions,
  WorkbookSessionPhase,
  WorkbookSessionStatus,
} from './workbook-session.js';

/**
 * The session behind `createWorkbookSession`. Its members are bound, so the editor can pass them
 * on as callbacks.
 */
export class StoredWorkbookSession implements WorkbookSession {
  readonly workbookId: string;
  readonly ready: Promise<boolean>;
  private readonly autosaveDelay: number;
  private readonly draftDelay: number;
  private readonly now: () => string;
  private readonly listeners = new Set<() => void>();
  private readonly loading = new AbortController();
  private readonly tracker = createTransactionReceiptTracker();
  private current: Workbook | null = null;
  private stored: SavedState | null = null;
  private state: WorkbookSessionStatus = {
    phase: 'loading',
    revision: null,
    dirty: false,
    automatic: false,
    error: null,
  };
  private offer: (WorkbookDraftOffer & { content: Workbook }) | null = null;
  private closed = false;
  /** The local revision of the latest accepted change, and the latest one a save answer covers. */
  private localRevision = 0;
  private savedThrough = 0;
  private inFlight: Promise<void> | null = null;
  private autosaveTimer: Timer | null = null;
  private draftTimer: Timer | null = null;
  private stopStore: (() => void) | undefined;
  private stopLifecycle: (() => void) | undefined;

  constructor(private readonly options: WorkbookSessionOptions) {
    this.workbookId = options.workbookId;
    this.autosaveDelay = options.autosaveDelay ?? DEFAULT_AUTOSAVE_DELAY;
    this.draftDelay = options.draftDelay ?? DEFAULT_DRAFT_DELAY;
    this.now = options.now ?? (() => new Date().toISOString());
    this.ready = this.load();
  }

  readonly workbook = () => this.current;
  readonly status = () => this.state;
  readonly draft = (): WorkbookDraftOffer | null =>
    this.offer
      ? {
          savedAt: this.offer.savedAt,
          baseRevision: this.offer.baseRevision,
          stale: this.offer.stale,
        }
      : null;
  readonly subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  readonly change = (next: Workbook, context: WorkbookChangeContext) => {
    if (this.closed) throw closedError();
    if (!this.stored) throw new Error('The workbook is not loaded.');
    this.localRevision++;
    this.tracker.record(context.transactionId, this.localRevision);
    this.current = next;
    const { phase } = this.state;
    this.setStatus(
      phase === 'saving' || phase === 'conflict' || phase === 'error' ? {} : { phase: 'dirty' },
    );
    this.scheduleDraft();
    this.scheduleAutosave();
  };

  readonly acknowledge = (transactionId: string) => {
    const revision = this.tracker.acknowledge(transactionId);
    if (revision === undefined || revision <= this.savedThrough) return;
    this.savedThrough = revision;
    if (this.localRevision === this.savedThrough) this.clearDraft();
    this.setStatus(this.state.phase === 'saving' ? {} : { phase: this.idlePhase() });
  };

  readonly save = () => this.run(false);

  readonly flush = async () => {
    this.clearTimer(this.draftTimer);
    const requested = this.localRevision;
    while (this.savedThrough < requested) await this.run(false);
    if (this.inFlight) await this.inFlight;
  };

  readonly keepMine = async () => {
    if (this.state.phase !== 'conflict' || !this.stored) return;
    await this.run(false, this.stored.revision);
  };

  readonly takeStored = async () => {
    if (this.closed) return;
    const result = await this.options.workbooks.load(this.workbookId);
    if (this.closed) return;
    if (!result.ok) {
      this.setStatus({ phase: 'error', error: result.error });
      throw new HostRequestError(result.error);
    }
    this.savedThrough = this.localRevision;
    this.stopAutosave();
    this.clearDraft();
    this.adopt(result.value);
  };

  readonly restoreDraft = () => {
    if (!this.offer || this.closed || !this.stored) return false;
    this.current = copyWorkbook(this.offer.content);
    this.offer = null;
    this.localRevision++;
    this.tracker.record(`draft:${this.localRevision}`, this.localRevision);
    this.setStatus({ phase: 'dirty' });
    this.scheduleAutosave();
    return true;
  };

  readonly discardDraft = () => {
    if (!this.offer) return;
    this.offer = null;
    this.clearDraft();
    this.emit();
  };

  readonly close = () => {
    if (this.closed) return;
    this.closed = true;
    this.loading.abort();
    this.clearTimer(this.autosaveTimer);
    this.clearTimer(this.draftTimer);
    this.stopStore?.();
    this.stopLifecycle?.();
    this.listeners.clear();
  };

  private get unsaved() {
    return this.localRevision > this.savedThrough;
  }
  private emit() {
    for (const listener of [...this.listeners]) listener();
  }
  private setStatus(next: Partial<WorkbookSessionStatus>) {
    this.state = { ...this.state, dirty: this.unsaved, ...next };
    this.emit();
  }
  private idlePhase(): WorkbookSessionPhase {
    return this.unsaved ? 'dirty' : 'saved';
  }
  private clearTimer(timer: Timer | null) {
    if (timer) clearTimeout(timer);
  }
  private stopAutosave() {
    this.clearTimer(this.autosaveTimer);
    this.autosaveTimer = null;
  }

  private scheduleAutosave() {
    this.stopAutosave();
    const { phase } = this.state;
    if (this.closed || phase === 'conflict' || phase === 'error') return;
    if (!Number.isFinite(this.autosaveDelay)) return;
    this.autosaveTimer = setTimeout(() => {
      this.autosaveTimer = null;
      void this.run(true).catch(() => {});
    }, this.autosaveDelay);
  }
  private scheduleDraft() {
    const { drafts } = this.options;
    if (!drafts || this.closed) return;
    this.clearTimer(this.draftTimer);
    this.draftTimer = setTimeout(() => {
      this.draftTimer = null;
      if (this.closed || !this.current || !this.stored || !this.unsaved) return;
      const draft: WorkbookDraft = {
        content: copyWorkbook(this.current),
        baseRevision: this.stored.revision,
        savedAt: this.now(),
      };
      void drafts.put(this.workbookId, draft).catch(() => {});
    }, this.draftDelay);
  }
  private clearDraft() {
    this.clearTimer(this.draftTimer);
    this.draftTimer = null;
    if (this.options.drafts) void this.options.drafts.clear(this.workbookId).catch(() => {});
  }

  private saved(value: SavedState, covered: number) {
    this.stored = value;
    this.savedThrough = Math.max(this.savedThrough, covered);
    this.tracker.forgetThrough(covered);
    if (this.localRevision === this.savedThrough) this.clearDraft();
    this.setStatus({ phase: this.idlePhase(), revision: value.revision, error: null });
    if (this.unsaved) this.scheduleAutosave();
  }
  private refused(error: HostError): never {
    if (error.kind === 'conflict') {
      if (error.current) this.stored = error.current;
      this.setStatus({ phase: 'conflict', error });
    } else this.setStatus({ phase: 'error', error });
    throw new HostRequestError(error);
  }

  /** One save, retried with backoff while an autosave finds the store unavailable. */
  private async saveOnce(automatic: boolean, baseRevision: string): Promise<void> {
    const covered = this.localRevision,
      content = copyWorkbook(requireDefined(this.current));
    const started = Date.now();
    let retry = 500;
    this.setStatus({ phase: 'saving', automatic, error: null });
    while (true) {
      const result = await answer(() =>
        this.options.workbooks.save(this.workbookId, { content, baseRevision }),
      );
      if (this.closed) throw closedError();
      if (result.ok) return this.saved(result.value, covered);
      const { error } = result;
      if (!automatic || error.kind !== 'unavailable' || Date.now() - started >= RETRY_WINDOW)
        this.refused(error);
      await wait(retry);
      retry = Math.min(4000, retry * 2);
      if (this.closed) throw closedError();
    }
  }

  /** Saves until an answer covers every change accepted before it was asked for. */
  private run(automatic: boolean, baseRevision?: string): Promise<void> {
    if (this.closed) return Promise.reject(closedError());
    if (!this.current || !this.stored)
      return Promise.reject(new Error('The workbook is not loaded.'));
    if (this.inFlight) {
      const requested = this.localRevision;
      return this.inFlight.then(() =>
        this.savedThrough < requested ? this.run(automatic) : undefined,
      );
    }
    const { phase } = this.state;
    if ((phase === 'conflict' && baseRevision === undefined) || (automatic && phase === 'error'))
      return Promise.reject(new HostRequestError(this.state.error ?? { kind: 'conflict' }));
    if (!this.unsaved) return Promise.resolve();
    this.stopAutosave();
    const operation = this.saveOnce(automatic, baseRevision ?? this.stored.revision).finally(() => {
      this.inFlight = null;
    });
    this.inFlight = operation;
    return operation;
  }

  private async offerDraft(revision: string) {
    const { drafts } = this.options;
    if (!drafts) return;
    const kept = await drafts.get(this.workbookId).catch(() => null);
    if (this.closed || !kept?.ok || !kept.value) return;
    const { content, savedAt, baseRevision } = kept.value;
    this.offer = { content, savedAt, baseRevision, stale: baseRevision !== revision };
    this.emit();
  }
  private listen() {
    const { workbooks, lifecycle } = this.options;
    if (!this.closed && workbooks.subscribe)
      this.stopStore = workbooks.subscribe(
        this.workbookId,
        (state) => void this.reloadAfter(state),
      );
    if (!this.closed && lifecycle)
      this.stopLifecycle = lifecycle.subscribe((event) => {
        if (event !== 'visible' && this.unsaved) void this.run(true).catch(() => {});
      });
  }
  private async load(): Promise<boolean> {
    const result = await answer(() =>
      this.options.workbooks.load(this.workbookId, this.loading.signal),
    );
    if (this.closed) return false;
    if (!result.ok) {
      this.setStatus({ phase: 'unavailable', error: result.error });
      return false;
    }
    this.adopt(result.value);
    await this.offerDraft(result.value.revision);
    this.listen();
    return !this.closed;
  }
  private adopt(value: StoredWorkbook) {
    this.current = value.content;
    this.stored = { revision: value.revision, sequence: value.sequence };
    this.setStatus({ phase: this.idlePhase(), revision: value.revision, error: null });
  }
  /** Another save landed: a session without unsaved changes shows it. */
  private async reloadAfter(state: SavedState) {
    const { stored } = this;
    if (this.closed || !stored || state.sequence <= stored.sequence) return;
    if (state.revision === stored.revision || this.unsaved || this.inFlight) return;
    const result = await this.options.workbooks.load(this.workbookId).catch(() => null);
    if (this.closed || !result?.ok || this.unsaved || this.inFlight) return;
    if (result.value.sequence > (this.stored?.sequence ?? -1)) this.adopt(result.value);
  }
}
