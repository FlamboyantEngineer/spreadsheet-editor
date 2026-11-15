'use client';

import { type RefObject, useEffect, useEffectEvent, useState } from 'react';
import { type CellSelection, selectionAddress } from '../model/editing.js';
import {
  addComment,
  addWorkbookComment,
  decideSuggestion,
  findOpenComment,
  type ReviewStamp,
} from '../model/review.js';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import type { Workbook, Worksheet } from '../model/workbook.js';
import type { GridController } from './contracts.js';

interface CommentNavigation {
  workbook: Workbook;
  initialCommentId: string | undefined;
  grid: RefObject<GridController | null>;
  select: (range: CellSelection, targetSheetId?: string) => boolean;
}

/** The open thread and the composer: opening a thread selects its range (SE-REVIEW-1.1). */
function useOpenComment({ workbook, initialCommentId, grid, select }: CommentNavigation) {
  const [commentComposer, setCommentComposer] = useState(false);
  const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
  function openComment(id: string) {
    const thread = workbook.comments?.find((thread) => thread.id === id);
    if (!thread) return;
    if (thread.scope === 'workbook' || thread.state === 'resolved' || thread.orphaned) {
      setActiveCommentId(thread.id);
      return;
    }
    const range = parseSpreadsheetRange(thread.range ?? '');
    if (range && thread.sheetId)
      select(
        { x1: range.start.x, x2: range.end.x, y1: range.start.y, y2: range.end.y },
        thread.sheetId,
      );
    setCommentComposer(false);
    setActiveCommentId(id);
    grid.current?.releaseKeyboard();
  }
  const openInitialComment = useEffectEvent((id: string) => openComment(id));
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- a host deep link opens its thread when it is given
    if (initialCommentId) openInitialComment(initialCommentId);
  }, [initialCommentId]);
  const closeVanishedComment = useEffectEvent((current: Workbook) => {
    const thread = current.comments?.find((thread) => thread.id === activeCommentId);
    if (activeCommentId !== 'draft' && (!thread || thread.state === 'resolved' || thread.orphaned))
      setActiveCommentId(null);
  });
  // A thread that leaves the workbook closes, and stays closed if an undo brings it back.
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- closes once per workbook change
    closeVanishedComment(workbook);
  }, [workbook]);
  return {
    commentComposer,
    setCommentComposer,
    activeCommentId,
    setActiveCommentId,
    openComment,
  };
}

interface CommentInputs extends CommentNavigation {
  sheet: Worksheet;
  selection: CellSelection;
  canComment: boolean;
  canDecide: boolean;
  openPanel: (id: string) => void;
  commit: (next: Workbook, label?: string) => void;
  stamp: () => ReviewStamp;
}

/** The workbook with a new range or workbook comment. */
function commentedWorkbook(
  workbook: Workbook,
  {
    scope,
    sheetId,
    selection,
  }: { scope: 'range' | 'workbook'; sheetId: string; selection: CellSelection },
  { text, meta }: { text: string; meta: ReviewStamp },
) {
  return scope === 'workbook'
    ? addWorkbookComment(workbook, text, meta)
    : addComment(workbook, { sheetId, range: selectionAddress(selection), text }, meta);
}

/** Comment threads and suggestion decisions: the composer, the open thread and reply drafts. */
export function useComments({
  workbook,
  sheet,
  selection,
  initialCommentId,
  canComment,
  canDecide,
  grid,
  select,
  openPanel,
  commit,
  stamp,
}: CommentInputs) {
  const { commentComposer, setCommentComposer, activeCommentId, setActiveCommentId, openComment } =
    useOpenComment({ workbook, initialCommentId, grid, select });
  const [commentScope, setCommentScope] = useState<'range' | 'workbook'>('range');
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({});
  function beginComment() {
    if (!canComment) return;
    const existing = findOpenComment(workbook, sheet.id, selectionAddress(selection));
    if (existing) {
      openComment(existing.id);
      return;
    }
    setCommentScope('range');
    setCommentComposer(true);
    setActiveCommentId('draft');
    grid.current?.releaseKeyboard();
  }
  function beginWorkbookComment() {
    if (!canComment) return;
    openPanel('comments');
    setCommentScope('workbook');
    setActiveCommentId(null);
    setCommentComposer(true);
  }
  function postComment(text: string) {
    if (!canComment) return;
    const meta = stamp();
    commit(
      commentedWorkbook(
        workbook,
        { scope: commentScope, sheetId: sheet.id, selection },
        { text, meta },
      ),
      'Add comment',
    );
    setCommentComposer(false);
    if (commentScope === 'range') setActiveCommentId(meta.id);
  }
  function openSuggestion(id: string) {
    const item = workbook.suggestions?.find((suggestion) => suggestion.id === id);
    if (item?.commentThreadId) openComment(item.commentThreadId);
  }
  function decide(id: string, action: 'accepted' | 'rejected') {
    if (canDecide)
      commit(
        decideSuggestion(workbook, { id, decision: action }, stamp()),
        `${action === 'accepted' ? 'Accept' : 'Reject'} suggestion`,
      );
  }
  return {
    commentComposer,
    setCommentComposer,
    commentScope,
    activeCommentId,
    setActiveCommentId,
    replyDrafts,
    setReplyDrafts,
    openComment,
    beginComment,
    beginWorkbookComment,
    postComment,
    openSuggestion,
    decide,
  };
}
