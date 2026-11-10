import { describe, expect, it } from 'vitest';
import { createWorkbook, copyWorkbook } from './workbook.js';
import { writeCells, formatCells } from './editing.js';
import {
  addComment,
  replyToComment,
  updateComment,
  editCommentPost,
  deleteCommentPost,
  toggleCommentReaction,
  proposeWorkbookChange,
  decideSuggestion,
  type ReviewStamp,
} from './review.js';
import { changeWorkbookStructure } from './structure.js';
import { calculateWorkbook } from '../calculation/evaluator.js';

const alex = { id: 'alex', name: 'Alex' },
  morgan = { id: 'morgan', name: 'Morgan' };
const stamp = (id: string, actor = alex): ReviewStamp => ({
  id,
  actor,
  time: '2026-11-10T12:00:00.000Z',
});

describe('SE-REVIEW-1: persisted threaded comments', () => {
  it('round-trips anchors, authors, replies, resolution and reactions without changing cells', () => {
    const source = createWorkbook([['1E-25', '=1-A1']]);
    let next = addComment(source, 'sheet-1', 'B1', 'Check the precision', stamp('thread'));
    next = replyToComment(next, 'thread', 'Confirmed', stamp('reply', morgan));
    next = toggleCommentReaction(next, 'thread', 'reply', '👍', alex.id);
    next = updateComment(next, 'thread', (thread) => {
      thread.resolved = true;
    });
    const stored = copyWorkbook(JSON.parse(JSON.stringify(next)));
    expect(stored.comments?.[0]).toMatchObject({
      range: 'B1',
      resolved: true,
      posts: [{ author: alex }, { author: morgan, reactions: { '👍': ['alex'] } }],
    });
    expect(stored.worksheets).toEqual(source.worksheets);
    expect(source.comments).toBeUndefined();
    expect(() => replyToComment(stored, 'thread', 'New reply', stamp('later'))).toThrow('Reopen');
    next = updateComment(stored, 'thread', (thread) => {
      thread.resolved = false;
    });
    expect(
      replyToComment(next, 'thread', 'Reopened', stamp('later')).comments?.[0].posts,
    ).toHaveLength(3);
  });
  it('restricts edit/delete to the post author and preserves immutable source snapshots', () => {
    const workbook = addComment(createWorkbook(), 'sheet-1', 'A1', 'Original', stamp('thread'));
    expect(() =>
      editCommentPost(workbook, 'thread', 'thread', 'Changed', morgan.id, stamp('a').time),
    ).toThrow('author');
    expect(() => deleteCommentPost(workbook, 'thread', 'thread', morgan.id)).toThrow('author');
    const next = editCommentPost(workbook, 'thread', 'thread', 'Changed', alex.id, stamp('a').time);
    expect(next.comments?.[0].posts[0].text).toBe('Changed');
    expect(workbook.comments?.[0].posts[0].text).toBe('Original');
    expect(deleteCommentPost(next, 'thread', 'thread', alex.id).comments).toEqual([]);
  });
  it('validates imported comment authors, dates, identities and reaction lists', () => {
    const original = addComment(createWorkbook(), 'sheet-1', 'A1', 'Text', stamp('thread'));
    const duplicate = structuredClone(original);
    duplicate.comments!.push(structuredClone(duplicate.comments![0]));
    expect(() => copyWorkbook(duplicate)).toThrow('comment ids');
    const malformed = structuredClone(original);
    malformed.comments![0].posts[0].createdAt = 'invalid';
    expect(() => copyWorkbook(malformed)).toThrow('comment post');
  });
});

describe('SE-REVIEW-2: tracked suggestions', () => {
  it('keeps canonical values unchanged until acceptance and retains the decision record', () => {
    const source = createWorkbook([['1E-25', '=1-A1']]);
    const proposed = writeCells(source, 'sheet-1', 0, 0, [['2E-25']]);
    const review = proposeWorkbookChange(source, proposed, 'Edit A1', stamp('suggestion', morgan));
    expect(review.worksheets).toEqual(source.worksheets);
    const accepted = decideSuggestion(review, 'suggestion', 'accepted', stamp('decision'));
    expect(accepted.worksheets[0].data[0][0]).toBe('2E-25');
    expect(calculateWorkbook(accepted).worksheets[0].values[0][1]).toMatchObject({
      numerator: '4999999999999999999999999',
      denominator: '5000000000000000000000000',
    });
    expect(accepted.suggestions?.[0]).toMatchObject({ status: 'accepted', decidedBy: alex });
    expect(review.suggestions?.[0].status).toBe('pending');
  });
  it('rejects without applying and refuses stale or already decided proposals', () => {
    const source = createWorkbook([['1']]);
    const review = proposeWorkbookChange(
      source,
      writeCells(source, 'sheet-1', 0, 0, [['2']]),
      'Edit',
      stamp('suggestion'),
    );
    const rejected = decideSuggestion(review, 'suggestion', 'rejected', stamp('decision'));
    expect(rejected.worksheets[0].data).toEqual([['1']]);
    expect(() => decideSuggestion(rejected, 'suggestion', 'accepted', stamp('again'))).toThrow(
      'pending',
    );
    const conflicting = writeCells(review, 'sheet-1', 0, 0, [['3']]);
    expect(() =>
      decideSuggestion(conflicting, 'suggestion', 'accepted', stamp('decision')),
    ).toThrow('not been overwritten');
    expect(conflicting.worksheets[0].data).toEqual([['3']]);
  });
  it('tracks formatting, including blank cells, separately from canonical numeric values', () => {
    const source = createWorkbook([['1E-25']]);
    const next = formatCells(
      source,
      'sheet-1',
      { x1: 0, y1: 0, x2: 1, y2: 0 },
      { 'font-weight': 'bold' },
      { kind: 'percent', decimals: 30 },
    );
    const review = proposeWorkbookChange(source, next, 'Bold percent', stamp('format'));
    expect(review.worksheets[0].style).toBeUndefined();
    const accepted = decideSuggestion(review, 'format', 'accepted', stamp('decision'));
    expect(accepted.worksheets[0].style?.B1).toBe('font-weight: bold');
    expect(accepted.worksheets[0].formats?.A1).toEqual({ kind: 'percent', decimals: 30 });
    expect(accepted.worksheets[0].data[0][0]).toBe('1E-25');
  });
  it('applies structural proposals and moves comments added after the proposal', () => {
    const source = createWorkbook([['1'], ['=A1']]);
    const operation = {
      sheetId: 'sheet-1',
      change: { axis: 'row', kind: 'insert', index: 0, count: 1 },
    } as const;
    let review = proposeWorkbookChange(
      source,
      changeWorkbookStructure(source, 'sheet-1', operation.change),
      'Insert row',
      stamp('structure'),
      operation,
    );
    review = addComment(review, 'sheet-1', 'A2', 'Move this anchor too', stamp('comment'));
    const accepted = decideSuggestion(review, 'structure', 'accepted', stamp('decision'));
    expect(accepted.worksheets[0].data[2][0]).toBe('=A2');
    expect(accepted.comments?.[0].range).toBe('A3');
    expect(accepted.suggestions?.[0].status).toBe('accepted');
  });
});
