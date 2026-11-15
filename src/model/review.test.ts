import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { formatCells, writeCells } from './editing.js';
import { requireDefined } from './require-defined.js';
import {
  addComment,
  addWorkbookComment,
  decideSuggestion,
  deleteCommentEntry,
  editCommentEntry,
  findOpenComment,
  previewSuggestions,
  proposeWorkbookChange,
  replyToComment,
  type ReviewStamp,
  setCommentState,
  toggleCommentReaction,
} from './review.js';
import { changeWorkbookStructure } from './structure.js';
import { copyWorkbook, createWorkbook } from './workbook.js';

const alex = { id: 'alex', name: 'Alex' },
  morgan = { id: 'morgan', name: 'Morgan' };
const stamp = (id: string, actor = alex): ReviewStamp => ({
  id,
  actor,
  time: '2026-11-10T12:00:00.000Z',
});

describe('SE-REVIEW-1.1: persisted threaded comments', () => {
  it('round-trips anchors, authors, replies, resolution and reactions without changing cells', () => {
    const source = createWorkbook([['1E-25', '=1-A1']]);
    let next = addComment(
      source,
      { sheetId: 'sheet-1', range: 'B1', text: 'Check the precision' },
      stamp('thread'),
    );
    next = replyToComment(next, { id: 'thread', text: 'Confirmed' }, stamp('reply', morgan));
    next = toggleCommentReaction(
      next,
      { threadId: 'thread', postId: 'reply' },
      { emoji: '👍', actorId: alex.id },
    );
    next = setCommentState(next, { id: 'thread', state: 'resolved' }, stamp('resolve', morgan));
    const stored = copyWorkbook(JSON.parse(JSON.stringify(next)));
    expect(stored.comments?.[0]).toMatchObject({
      range: 'B1',
      state: 'resolved',
      resolvedBy: 'morgan',
      resolvedAt: '2026-11-10T12:00:00.000Z',
      entries: [
        { author: 'alex', createdAt: '2026-11-10T12:00:00.000Z' },
        { author: 'morgan', reactions: { '👍': ['alex'] } },
      ],
    });
    expect(stored.worksheets).toEqual(source.worksheets);
    expect(source.comments).toBeUndefined();
    expect(() =>
      replyToComment(stored, { id: 'thread', text: 'New reply' }, stamp('later')),
    ).toThrow('Reopen');
    next = setCommentState(stored, { id: 'thread', state: 'open' }, stamp('reopen'));
    expect(next.comments?.[0].resolvedBy).toBeUndefined();
    expect(next.comments?.[0].resolvedAt).toBeUndefined();
    expect(
      replyToComment(next, { id: 'thread', text: 'Reopened' }, stamp('later')).comments?.[0]
        .entries,
    ).toHaveLength(3);
  });
  it('restricts edit/delete to the post author and preserves immutable source snapshots', () => {
    const workbook = addComment(
      createWorkbook(),
      { sheetId: 'sheet-1', range: 'A1', text: 'Original' },
      stamp('thread'),
    );
    expect(() =>
      editCommentEntry(
        workbook,
        { threadId: 'thread', postId: 'thread' },
        { text: 'Changed', actorId: morgan.id, time: stamp('a').time },
      ),
    ).toThrow('author');
    expect(() =>
      deleteCommentEntry(workbook, { threadId: 'thread', postId: 'thread' }, morgan.id),
    ).toThrow('author');
    const next = editCommentEntry(
      workbook,
      { threadId: 'thread', postId: 'thread' },
      { text: 'Changed', actorId: alex.id, time: stamp('a').time },
    );
    expect(next.comments?.[0].entries[0]).toMatchObject({
      text: 'Changed',
      editedAt: stamp('a').time,
    });
    expect(workbook.comments?.[0].entries[0].text).toBe('Original');
    expect(
      deleteCommentEntry(next, { threadId: 'thread', postId: 'thread' }, alex.id).comments,
    ).toEqual([]);
  });
  it('validates imported comment authors, dates, identities and reaction lists', () => {
    const original = addComment(
      createWorkbook(),
      { sheetId: 'sheet-1', range: 'A1', text: 'Text' },
      stamp('thread'),
    );
    const duplicate = structuredClone(original);
    requireDefined(duplicate.comments).push(structuredClone(requireDefined(duplicate.comments)[0]));
    expect(() => copyWorkbook(duplicate)).toThrow('comment ids');
    const malformed = structuredClone(original);
    requireDefined(malformed.comments)[0].entries[0].createdAt = 'invalid';
    expect(() => copyWorkbook(malformed)).toThrow('comment entry');
    const earlierShape = structuredClone(original) as unknown as {
      comments: Record<string, unknown>[];
    };
    const thread = earlierShape.comments[0];
    thread.posts = thread.entries;
    thread.resolved = false;
    delete thread.entries;
    delete thread.state;
    expect(() => copyWorkbook(earlierShape as never)).toThrow('comment thread');
  });
  it('SE-REVIEW-1.1: records who resolved a thread and clears it on reopening', () => {
    const open = addComment(
      createWorkbook(),
      { sheetId: 'sheet-1', range: 'A1', text: 'Text' },
      stamp('thread'),
    );
    const resolved = setCommentState(open, { id: 'thread', state: 'resolved' }, stamp('r', morgan));
    expect(resolved.comments?.[0]).toMatchObject({
      state: 'resolved',
      resolvedBy: 'morgan',
      resolvedAt: '2026-11-10T12:00:00.000Z',
    });
    expect(open.comments?.[0].state).toBe('open');
    const reopened = setCommentState(resolved, { id: 'thread', state: 'open' }, stamp('o'));
    expect(reopened.comments?.[0]).not.toHaveProperty('resolvedBy');
    expect(reopened.comments?.[0]).not.toHaveProperty('resolvedAt');
  });
});

describe('SE-REVIEW-2.1: tracked suggestions', () => {
  it('keeps canonical values unchanged until acceptance and retains the decision record', () => {
    const source = createWorkbook([['1E-25', '=1-A1']]);
    const proposed = writeCells(source, [['2E-25']], { sheetId: 'sheet-1', x: 0, y: 0 });
    const review = proposeWorkbookChange(
      source,
      { proposed, summary: 'Edit A1' },
      stamp('suggestion', morgan),
    );
    expect(review.worksheets).toEqual(source.worksheets);
    const accepted = decideSuggestion(
      review,
      { id: 'suggestion', decision: 'accepted' },
      stamp('decision'),
    );
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
      {
        proposed: writeCells(source, [['2']], { sheetId: 'sheet-1', x: 0, y: 0 }),
        summary: 'Edit',
      },
      stamp('suggestion'),
    );
    const rejected = decideSuggestion(
      review,
      { id: 'suggestion', decision: 'rejected' },
      stamp('decision'),
    );
    expect(rejected.worksheets[0].data).toEqual([['1']]);
    expect(() =>
      decideSuggestion(rejected, { id: 'suggestion', decision: 'accepted' }, stamp('again')),
    ).toThrow('pending');
    const conflicting = writeCells(review, [['3']], { sheetId: 'sheet-1', x: 0, y: 0 });
    expect(() =>
      decideSuggestion(conflicting, { id: 'suggestion', decision: 'accepted' }, stamp('decision')),
    ).toThrow('not been overwritten');
    expect(conflicting.worksheets[0].data).toEqual([['3']]);
  });
  it('tracks formatting, including blank cells, separately from canonical numeric values', () => {
    const source = createWorkbook([['1E-25']]);
    const next = formatCells(
      source,
      { sheetId: 'sheet-1', range: { x1: 0, y1: 0, x2: 1, y2: 0 } },
      { properties: { 'font-weight': 'bold' }, format: { kind: 'percent', decimals: 30 } },
    );
    const review = proposeWorkbookChange(
      source,
      { proposed: next, summary: 'Bold percent' },
      stamp('format'),
    );
    expect(review.worksheets[0].style).toBeUndefined();
    const accepted = decideSuggestion(
      review,
      { id: 'format', decision: 'accepted' },
      stamp('decision'),
    );
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
      {
        proposed: changeWorkbookStructure(source, 'sheet-1', operation.change),
        summary: 'Insert row',
        operation,
      },
      stamp('structure'),
    );
    review = addComment(
      review,
      { sheetId: 'sheet-1', range: 'A2', text: 'Move this anchor too' },
      stamp('comment'),
    );
    const accepted = decideSuggestion(
      review,
      { id: 'structure', decision: 'accepted' },
      stamp('decision'),
    );
    expect(accepted.worksheets[0].data[2][0]).toBe('=A2');
    expect(accepted.comments?.[0].range).toBe('A3');
    expect(accepted.suggestions?.[0].status).toBe('accepted');
  });
});

describe('SE-REVIEW-7.1: conversation and reviewer overlay contracts', () => {
  it('prevents overlapping open threads, keeps workbook scope, and reopens a resolved range', () => {
    let workbook = addComment(
      createWorkbook(),
      { sheetId: 'sheet-1', range: 'A1:B2', text: 'Range' },
      stamp('range'),
    );
    expect(findOpenComment(workbook, 'sheet-1', 'B2:C3')?.id).toBe('range');
    expect(() =>
      addComment(
        workbook,
        { sheetId: 'sheet-1', range: 'B2', text: 'Duplicate' },
        stamp('duplicate'),
      ),
    ).toThrow('already');
    workbook = setCommentState(workbook, { id: 'range', state: 'resolved' }, stamp('resolve'));
    workbook = addComment(
      workbook,
      { sheetId: 'sheet-1', range: 'B2', text: 'Another' },
      stamp('next'),
    );
    workbook = addWorkbookComment(workbook, 'Whole workbook', stamp('whole'));
    expect(requireDefined(workbook.comments).at(-1)).toMatchObject({ scope: 'workbook' });
    expect(requireDefined(requireDefined(workbook.comments).at(-1)).range).toBeUndefined();
  });
  it('coalesces exact targets by author and shares a discussion without touching canonical calculations', () => {
    const original = createWorkbook([['1', '=A1*2']]);
    let workbook = proposeWorkbookChange(
      original,
      {
        proposed: writeCells(original, [['2']], { sheetId: 'sheet-1', x: 0, y: 0 }),
        summary: 'Edit',
      },
      stamp('first', morgan),
    );
    workbook = proposeWorkbookChange(
      workbook,
      {
        proposed: writeCells(workbook, [['3']], { sheetId: 'sheet-1', x: 0, y: 0 }),
        summary: 'Edit again',
      },
      stamp('second', morgan),
    );
    expect(workbook.suggestions).toHaveLength(1);
    expect(workbook.comments).toHaveLength(1);
    expect(requireDefined(workbook.suggestions)[0].changes[0]).toMatchObject({
      before: '1',
      after: '3',
    });
    expect(previewSuggestions(workbook, alex.id)).toBe(workbook);
    const preview = previewSuggestions(workbook, morgan.id);
    expect(calculateWorkbook(preview).worksheets[0].values[0][1]).toMatchObject({ numerator: '6' });
    expect(calculateWorkbook(workbook).worksheets[0].values[0][1]).toMatchObject({
      numerator: '2',
    });
    expect(findOpenComment(workbook, 'sheet-1', 'A1')).toBeUndefined();
    const accepted = decideSuggestion(
      workbook,
      { id: 'first', decision: 'accepted' },
      stamp('decision'),
    );
    expect(requireDefined(accepted.comments)[0]).toMatchObject({
      suggestion: 'first',
      state: 'resolved',
      resolvedBy: 'alex',
      decision: { decision: 'accept', suggestion: 'first', by: 'alex' },
    });
    expect(original.comments).toBeUndefined();
  });
});
