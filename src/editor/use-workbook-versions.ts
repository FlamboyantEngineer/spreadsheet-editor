'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import type { VersionStoreAdapter, WorkbookVersion } from '../host/adapters.js';

export interface WorkbookVersionList {
  /** The workbook's versions, newest first. Empty until the first list answers. */
  versions: readonly WorkbookVersion[];
  loading: boolean;
  error: string | null;
  refresh(): void;
}

const NO_VERSIONS: readonly WorkbookVersion[] = [];

/**
 * Lists a workbook's saved versions through the host's version store while `active`, such as
 * while the history pane is open, and lists them again whenever the store reports a change
 * (SE-ADAPTERS-5.1). Only the latest list request is shown.
 */
export function useWorkbookVersions(
  store: VersionStoreAdapter | undefined,
  workbookId: string | undefined,
  active: boolean,
): WorkbookVersionList {
  const [versions, setVersions] = useState(NO_VERSIONS);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const request = useRef<AbortController | null>(null);

  const refresh = useCallback(() => {
    if (!store || workbookId === undefined) return;
    request.current?.abort();
    const controller = new AbortController();
    request.current = controller;
    setLoading(true);
    void store
      .list(workbookId, controller.signal)
      .catch((reason: unknown) => ({
        ok: false as const,
        error: { message: reason instanceof Error ? reason.message : 'Could not list versions.' },
      }))
      .then((result) => {
        if (controller.signal.aborted) return;
        setLoading(false);
        if (result.ok) {
          setVersions(result.value);
          setError(null);
        } else setError(result.error.message ?? 'Could not list versions.');
      });
  }, [store, workbookId]);

  useEffect(() => {
    if (!active || !store || workbookId === undefined) return;
    // eslint-disable-next-line react-hooks/set-state-in-effect -- the list loads from the host's store while the pane is open
    refresh();
    const stop = store.subscribe?.(workbookId, refresh);
    return () => {
      stop?.();
      request.current?.abort();
      request.current = null;
      setLoading(false);
    };
  }, [active, store, workbookId, refresh]);

  return { versions, loading, error, refresh };
}
