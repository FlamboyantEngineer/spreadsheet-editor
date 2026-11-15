import { useSyncExternalStore } from 'react';
import type {
  WorkbookDraftOffer,
  WorkbookSession,
  WorkbookSessionStatus,
} from '../host/workbook-session.js';
import type { Workbook } from '../model/workbook.js';

const NO_SUBSCRIPTION = () => () => {};
const NOTHING = () => null;

/** The session's workbook, status and draft offer, or nulls without a session. */
export function useSessionState(session: WorkbookSession | undefined): {
  workbook: Workbook | null;
  status: WorkbookSessionStatus | null;
  draft: WorkbookDraftOffer | null;
} {
  const subscribe = session ? session.subscribe : NO_SUBSCRIPTION;
  const workbook = useSyncExternalStore(subscribe, session ? session.workbook : NOTHING);
  const status = useSyncExternalStore(subscribe, session ? session.status : NOTHING);
  const draftKey = useSyncExternalStore(subscribe, () => {
    const draft = session?.draft();
    return draft ? `${draft.baseRevision}\n${draft.savedAt}\n${draft.stale}` : '';
  });
  const draft = draftKey ? (session?.draft() ?? null) : null;
  return { workbook, status, draft };
}

/** The save status the toolbar shows for a session's phase. */
export function sessionSavePhase(
  status: WorkbookSessionStatus | null,
): 'idle' | 'dirty' | 'saving' | 'saved' | 'error' {
  if (!status) return 'idle';
  switch (status.phase) {
    case 'loading':
    case 'unavailable':
      return 'idle';
    case 'conflict':
      return 'error';
    default:
      return status.phase;
  }
}

/** The notice for a failed save or flush. */
export function saveFailure(error: unknown, session: WorkbookSession): string {
  if (session.status().phase === 'conflict')
    return 'Someone else saved this workbook after you opened it.';
  return error instanceof Error ? error.message : 'Could not save the workbook.';
}
