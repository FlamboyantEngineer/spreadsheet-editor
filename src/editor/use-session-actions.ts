'use client';

import { type RefObject, useState } from 'react';
import type { Workbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { saveFailure } from './use-session-state.js';

/** Saving, flushing and conflict choices through the host's workbook session (SE-SESSION-1.4). */
export function useSessionActions({
  commitDraft,
  live,
  alive,
  setNotice,
  warn,
}: {
  commitDraft: () => Workbook;
  live: RefObject<{ canEdit: boolean; props: SpreadsheetEditorProps }>;
  alive: RefObject<boolean>;
  setNotice: (message: string | null) => void;
  warn: (message: string) => void;
}) {
  const [flushes, setFlushes] = useState(0);
  /** Saves through the workbook session now (SE-SESSION-1.4). */
  async function save() {
    if (!alive.current) throw new Error('The editor session has ended.');
    commitDraft();
    const current = live.current.props.session;
    if (!current) return;
    if (!live.current.canEdit) {
      warn('Saving is available only in Authoring mode.');
      return;
    }
    try {
      await current.save();
      if (alive.current) setNotice(null);
    } catch (error) {
      if (alive.current) setNotice(saveFailure(error, current));
      throw error;
    }
  }
  /** Commits the draft and waits until the session holds no unsaved change (SE-SESSION-1.4). */
  async function flush(): Promise<void> {
    if (!alive.current) throw new Error('The editor session has ended.');
    commitDraft();
    const current = live.current.props.session;
    if (!current) return;
    setFlushes((count) => count + 1);
    try {
      await current.flush();
      if (alive.current) setNotice(null);
    } catch (error) {
      if (alive.current) setNotice(saveFailure(error, current));
      throw error;
    } finally {
      if (alive.current) setFlushes((count) => count - 1);
    }
  }
  async function resolveConflict(choice: 'mine' | 'stored') {
    const current = live.current.props.session;
    if (!current) return;
    try {
      await (choice === 'mine' ? current.keepMine() : current.takeStored());
      if (alive.current) setNotice(null);
    } catch (error) {
      if (alive.current) setNotice(saveFailure(error, current));
    }
  }
  return { flushing: flushes > 0, save, flush, resolveConflict };
}
