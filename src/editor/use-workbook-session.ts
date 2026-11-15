'use client';

import { useEffect, useState } from 'react';
import {
  createWorkbookSession,
  type WorkbookSession,
  type WorkbookSessionOptions,
} from '../host/workbook-session.js';

/**
 * Opens a workbook session while the component is mounted, and a new one when the store, the
 * workbook id or the drafts or lifecycle adapters change; the previous one closes, keeping its
 * draft. It is null until the first effect runs, which the editor's `session` prop shows as
 * opening (SE-SESSION-1.4).
 */
export function useWorkbookSession(options: WorkbookSessionOptions): WorkbookSession | null {
  const { workbooks, workbookId, drafts, lifecycle, autosaveDelay, draftDelay, now } = options;
  const [session, setSession] = useState<WorkbookSession | null>(null);
  useEffect(() => {
    const opened = createWorkbookSession({
      workbooks,
      workbookId,
      drafts,
      lifecycle,
      autosaveDelay,
      draftDelay,
      now,
    });
    // eslint-disable-next-line react-hooks/set-state-in-effect -- the session is an external resource opened by this effect
    setSession(opened);
    return () => {
      opened.close();
      setSession((current) => (current === opened ? null : current));
    };
    // `now` is read once per session, like the delays.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- a new clock does not reopen the session
  }, [workbooks, workbookId, drafts, lifecycle, autosaveDelay, draftDelay]);
  return session;
}
