import { useEffect, useState, type RefObject } from 'react';
import { selectionAddress, type CellSelection } from '../model/editing.js';
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
}: {
  workbook: Workbook;
  sheet: Worksheet;
  selection: CellSelection;
  initialCommentId: string | undefined;
  canComment: boolean;
  canDecide: boolean;
  grid: RefObject<GridController | null>;
  select: (range: CellSelection, targetSheetId?: string) => boolean;
  openPanel: (id: string) => void;
  commit: (next: Workbook, label?: string) => void;
  stamp: () => ReviewStamp;
}) {
  const [commentComposer, setCommentComposer] = useState(false);
  const [commentScope, setCommentScope] = useState<'range' | 'workbook'>('range');
  const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({});
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
      commentScope === 'workbook'
        ? addWorkbookComment(workbook, text, meta)
        : addComment(workbook, sheet.id, selectionAddress(selection), text, meta),
      'Add comment',
    );
    setCommentComposer(false);
    if (commentScope === 'range') setActiveCommentId(meta.id);
  }
  useEffect(() => {
    if (initialCommentId) openComment(initialCommentId);
  }, [initialCommentId]);
  function openSuggestion(id: string) {
    const item = workbook.suggestions?.find((suggestion) => suggestion.id === id);
    if (item?.commentThreadId) openComment(item.commentThreadId);
  }
  useEffect(() => {
    const thread = workbook.comments?.find((thread) => thread.id === activeCommentId);
    if (activeCommentId !== 'draft' && (!thread || thread.state === 'resolved' || thread.orphaned))
      setActiveCommentId(null);
  }, [workbook]);
  function decide(id: string, action: 'accepted' | 'rejected') {
    if (canDecide)
      commit(
        decideSuggestion(workbook, id, action, stamp()),
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
