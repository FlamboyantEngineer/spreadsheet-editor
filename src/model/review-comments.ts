import { extendWorksheetIdentities, type WorkbookIdFactory } from './extent.js';
import { requireDefined } from './require-defined.js';
import type { ReviewStamp, SpreadsheetComment } from './review.js';
import { parseSpreadsheetRange } from './spreadsheet-addressing.js';
import { copyWorkbook, type Workbook } from './workbook.js';

/** A new comment thread: the range it anchors to and its opening text. */
export interface NewComment {
  sheetId: string;
  range: string;
  text: string;
  idFactory?: WorkbookIdFactory | undefined;
}
/** One comment of a thread, by the thread's id and the comment's. */
export interface CommentEntryRef {
  threadId: string;
  postId: string;
}

export function addComment(
  workbook: Workbook,
  { sheetId, range, text, idFactory }: NewComment,
  stamp: ReviewStamp,
): Workbook {
  if (!workbook.worksheets.some((sheet) => sheet.id === sheetId) || !parseSpreadsheetRange(range))
    throw new RangeError('Invalid comment anchor.');
  if (findOpenComment(workbook, sheetId, range))
    throw new Error('An open comment already covers this selection.');
  const next = copyWorkbook(workbook);
  if (!text.trim()) throw new TypeError('A comment needs text.');
  next.comments = [
    ...(next.comments ?? []),
    {
      id: stamp.id,
      sheetId,
      range: requireDefined(parseSpreadsheetRange(range)).address,
      state: 'open',
      entries: [{ id: stamp.id, author: stamp.actor.id, text: text.trim(), createdAt: stamp.time }],
    },
  ];
  extendWorksheetIdentities(
    requireDefined(next.worksheets.find((sheet) => sheet.id === sheetId)),
    idFactory,
    next,
  );
  return next;
}

export function updateComment(
  workbook: Workbook,
  id: string,
  update: (thread: SpreadsheetComment) => void,
): Workbook {
  const next = copyWorkbook(workbook),
    thread = next.comments?.find((thread) => thread.id === id);
  if (!thread) throw new RangeError('Unknown comment.');
  update(thread);
  return next;
}

/** Resolve or reopen a thread, recording who resolved it and when (SE-REVIEW-1.1). */
export function setCommentState(
  workbook: Workbook,
  { id, state }: { id: string; state: SpreadsheetComment['state'] },
  stamp: Pick<ReviewStamp, 'time' | 'actor'>,
): Workbook {
  return updateComment(workbook, id, (thread) => {
    thread.state = state;
    if (state === 'resolved') {
      thread.resolvedBy = stamp.actor.id;
      thread.resolvedAt = stamp.time;
    } else {
      delete thread.resolvedBy;
      delete thread.resolvedAt;
    }
  });
}

export function replyToComment(
  workbook: Workbook,
  { id, text }: { id: string; text: string },
  stamp: ReviewStamp,
): Workbook {
  if (!text.trim()) throw new TypeError('A reply needs text.');
  return updateComment(workbook, id, (thread) => {
    if (thread.state === 'resolved') throw new Error('Reopen the thread before replying.');
    thread.entries.push({
      id: stamp.id,
      author: stamp.actor.id,
      text: text.trim(),
      createdAt: stamp.time,
    });
  });
}

export function editCommentEntry(
  workbook: Workbook,
  { threadId, postId }: CommentEntryRef,
  { text, actorId, time }: { text: string; actorId: string; time: string },
): Workbook {
  if (!text.trim()) throw new TypeError('A comment needs text.');
  return updateComment(workbook, threadId, (thread) => {
    const post = thread.entries.find((post) => post.id === postId);
    if (!post || post.author !== actorId) throw new Error('Only the author can edit this comment.');
    post.text = text.trim();
    post.editedAt = time;
  });
}

export function deleteCommentEntry(
  workbook: Workbook,
  { threadId, postId }: CommentEntryRef,
  actorId: string,
): Workbook {
  const next = updateComment(workbook, threadId, (thread) => {
    const post = thread.entries.find((post) => post.id === postId);
    if (!post || post.author !== actorId)
      throw new Error('Only the author can delete this comment.');
    thread.entries = thread.entries.filter((post) => post.id !== postId);
  });
  next.comments = next.comments?.filter((thread) => thread.entries.length);
  return next;
}

export function toggleCommentReaction(
  workbook: Workbook,
  { threadId, postId }: CommentEntryRef,
  { emoji, actorId }: { emoji: string; actorId: string },
): Workbook {
  return updateComment(workbook, threadId, (thread) => {
    const post = thread.entries.find((post) => post.id === postId);
    if (!post) throw new RangeError('Unknown comment entry.');
    const users = post.reactions?.[emoji] ?? [];
    post.reactions = {
      ...post.reactions,
      [emoji]: users.includes(actorId)
        ? users.filter((user) => user !== actorId)
        : [...users, actorId],
    };
  });
}

export function findOpenComment(
  workbook: Workbook,
  sheetId: string,
  range: string,
): SpreadsheetComment | undefined {
  const target = parseSpreadsheetRange(range);
  if (!target) return;
  return workbook.comments?.find((thread) => {
    if (
      thread.state === 'resolved' ||
      thread.orphaned ||
      thread.suggestion ||
      thread.sheetId !== sheetId
    )
      return false;
    const anchor = parseSpreadsheetRange(thread.range ?? '');
    return (
      !!anchor &&
      anchor.start.x <= target.end.x &&
      anchor.end.x >= target.start.x &&
      anchor.start.y <= target.end.y &&
      anchor.end.y >= target.start.y
    );
  });
}

export function addWorkbookComment(workbook: Workbook, text: string, stamp: ReviewStamp): Workbook {
  if (!text.trim()) throw new TypeError('A comment needs text.');
  const next = copyWorkbook(workbook);
  (next.comments ??= []).push({
    id: stamp.id,
    scope: 'workbook',
    state: 'open',
    entries: [{ id: stamp.id, author: stamp.actor.id, text: text.trim(), createdAt: stamp.time }],
  });
  return next;
}
