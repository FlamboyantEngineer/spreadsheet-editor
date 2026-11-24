'use client';

import { useEffect, useMemo, useState } from 'react';
import {
  createWorkbookSession,
  type WorkbookSession,
  type WorkbookSessionOptions,
} from '../host/workbook-session.js';

/**
 * Opens a workbook session while the component is mounted, and a new one when the store, the
 * workbook id or the drafts or lifecycle adapters change. The previous session closes. The hook
 * returns null until the effect opens a session for the current options, which the editor shows as
 * opening (SE-SESSION-1.4).
 */
export function useWorkbookSession(options: WorkbookSessionOptions): WorkbookSession | null {
  const { workbooks, workbookId, drafts, lifecycle, autosaveDelay, draftDelay, now } = options;
  const configuration = useMemo(
    () => ({
      workbooks,
      workbookId,
      drafts,
      lifecycle,
      autosaveDelay,
      draftDelay,
    }),
    [workbooks, workbookId, drafts, lifecycle, autosaveDelay, draftDelay],
  );
  const [opened, setOpened] = useState<{
    configuration: typeof configuration;
    session: WorkbookSession;
  } | null>(null);
  useEffect(() => {
    const session = createWorkbookSession({ ...configuration, now });
    // eslint-disable-next-line react-hooks/set-state-in-effect -- the session is an external resource opened by this effect
    setOpened({ configuration, session });
    return () => session.close();
    // `now` is read once per session.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- a new clock does not reopen the session
  }, [configuration]);
  return opened?.configuration === configuration ? opened.session : null;
}
