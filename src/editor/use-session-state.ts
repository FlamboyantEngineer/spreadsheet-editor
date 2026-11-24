'use client';

import { useSyncExternalStore } from 'react';
import type {
  WorkbookSession,
  WorkbookSessionSnapshot,
  WorkbookSessionStatus,
} from '../host/workbook-session.js';

const NO_SUBSCRIPTION = () => () => {};
const empty = { workbook: null, status: null, draft: null, localRevision: 0 } as const;
const emptySnapshot = () => empty;

/** The session's workbook, status and draft offer, or nulls without a session. */
export function useSessionState(session: WorkbookSession | undefined) {
  return useSyncExternalStore<WorkbookSessionSnapshot | typeof empty>(
    session?.subscribe ?? NO_SUBSCRIPTION,
    session?.getSnapshot ?? emptySnapshot,
    emptySnapshot,
  );
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

/** The current session failure, including autosaves started outside an editor action. */
export function sessionErrorMessage(status: WorkbookSessionStatus | null): string | null {
  if (status?.phase === 'conflict') return 'Someone else saved this workbook after you opened it.';
  if (status?.phase === 'error') return status.error?.message ?? 'Could not save the workbook.';
  return null;
}

/** The notice for a failed save or flush. */
export function saveFailure(error: unknown, session: WorkbookSession): string {
  if (session.status().phase === 'conflict')
    return 'Someone else saved this workbook after you opened it.';
  return error instanceof Error ? error.message : 'Could not save the workbook.';
}
