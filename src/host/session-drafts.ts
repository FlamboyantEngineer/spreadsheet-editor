import { copyWorkbook } from '../model/workbook.js';
import type { WorkbookDraft } from './adapters.js';
import { DEFAULT_DRAFT_DELAY, type Timer } from './session-requests.js';
import type { WorkbookSessionOptions } from './workbook-session.js';

/** Order writes so a late put cannot outlive a newer clear or put. */
export function createSessionDrafts({
  drafts: adapter,
  workbookId,
  draftDelay = DEFAULT_DRAFT_DELAY,
  now = () => new Date().toISOString(),
}: WorkbookSessionOptions) {
  let writes = Promise.resolve();
  let timer: Timer | null = null;
  function enqueue(draft: WorkbookDraft | null) {
    if (!adapter) return;
    writes = writes
      .then(async () => {
        if (draft) await adapter.put(workbookId, draft);
        else await adapter.clear(workbookId);
      })
      .catch(() => {});
  }
  function stop() {
    if (timer) clearTimeout(timer);
    timer = null;
  }
  return {
    stop,
    schedule(read: () => Omit<WorkbookDraft, 'savedAt'> | null) {
      if (!adapter) return;
      stop();
      timer = setTimeout(() => {
        timer = null;
        const draft = read();
        if (draft) enqueue({ ...draft, content: copyWorkbook(draft.content), savedAt: now() });
      }, draftDelay);
    },
    clear() {
      stop();
      enqueue(null);
    },
    async read(revision: string) {
      const answer = await adapter?.get(workbookId).catch(() => null);
      if (!answer?.ok || !answer.value) return null;
      const { content, baseRevision, savedAt } = answer.value;
      return { content, details: { baseRevision, savedAt, stale: baseRevision !== revision } };
    },
  };
}
