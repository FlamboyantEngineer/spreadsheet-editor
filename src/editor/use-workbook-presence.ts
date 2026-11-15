import { useEffect, useRef, useState } from 'react';
import type {
  LocalPresence,
  PresenceAdapter,
  PresenceStatus,
  SpreadsheetPeer,
} from '../host/presence-adapters.js';

const NO_PEERS: readonly SpreadsheetPeer[] = [];

function pageStatus(): PresenceStatus {
  return typeof document !== 'undefined' && document.visibilityState === 'hidden'
    ? 'away'
    : 'active';
}

/**
 * Subscribes to the other people in a workbook and shares the reader's place with them through
 * the host's presence adapter (SE-ADAPTERS-4.3). The reader is `away` while the page is hidden.
 * Without an adapter or a workbook id there are no peers and nothing is shared.
 */
export function useWorkbookPresence(
  presence: PresenceAdapter | undefined,
  workbookId: string | undefined,
  local: Omit<LocalPresence, 'status'>,
): readonly SpreadsheetPeer[] {
  const [peers, setPeers] = useState(NO_PEERS);
  const [status, setStatus] = useState(pageStatus);
  const latest = useRef({ ...local, status });
  latest.current = { ...local, status };

  useEffect(() => {
    if (!presence || workbookId === undefined) return;
    const stop = presence.subscribe(workbookId, setPeers);
    return () => {
      stop();
      presence.publish(workbookId, null);
      setPeers(NO_PEERS);
    };
  }, [presence, workbookId]);

  useEffect(() => {
    if (typeof document === 'undefined') return;
    const onVisibility = () => setStatus(pageStatus());
    document.addEventListener('visibilitychange', onVisibility);
    return () => document.removeEventListener('visibilitychange', onVisibility);
  }, []);

  const at = local.at && `${local.at.sheetId}!${local.at.range}`;
  const lock = local.editLock && `${local.editLock.sheetId}!${local.editLock.range}`;
  useEffect(() => {
    if (presence && workbookId !== undefined) presence.publish(workbookId, latest.current);
  }, [presence, workbookId, at, lock, status]);

  return peers;
}
