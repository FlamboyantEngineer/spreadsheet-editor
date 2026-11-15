'use client';

import { Attribution } from '@flamboyantengineer/ui-components/attribution';
import { Button } from '@flamboyantengineer/ui-components/button';
import { CommentComposer } from '@flamboyantengineer/ui-components/comment-composer';
import { ChatBubblePlusGlyph } from '@flamboyantengineer/ui-components/core';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { useState } from 'react';
import { selectionAddress } from '../model/editing.js';
import { requireDefined } from '../model/require-defined.js';
import type { SpreadsheetComment } from '../model/review.js';
import { ThreadCard } from './comment-thread-card.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type EditorProps = { editor: SpreadsheetEditorController };

const FILTERS = [
  { id: 'open', label: 'Open comments' },
  { id: 'resolved', label: 'Resolved comments' },
  { id: 'all', label: 'All comments' },
  { id: 'mine', label: 'My conversations' },
];

const noComments = (filter: string) => <p>No {filter === 'all' ? '' : filter + ' '}comments.</p>;

/** The composer of a new comment; cancelling closes it and the draft thread. */
function NewComment({ editor: e, label }: EditorProps & { label: string }) {
  return (
    <CommentComposer
      label={label}
      submitLabel="Post comment"
      onSubmit={e.postComment}
      onCancel={() => {
        e.setCommentComposer(false);
        e.setActiveCommentId(null);
      }}
      autoFocus
      isDisabled={!e.canComment}
    />
  );
}

/** One conversation in the list: where it is, how it starts, and its reply count. */
function ConversationSummary({ editor: e, thread }: EditorProps & { thread: SpreadsheetComment }) {
  const proposal = e.workbook.suggestions?.find((item) => item.id === thread.suggestion);
  const sheet = e.workbook.worksheets.find((sheet) => sheet.id === thread.sheetId);
  const first = thread.entries[0];
  const title = proposal
    ? `Suggestion · ${proposal.status}`
    : thread.scope === 'workbook'
      ? 'Workbook comment'
      : `${sheet?.name ?? 'Deleted worksheet'}!${thread.range}`;
  return (
    <div className="se-comments__conversation">
      <Button variant="ghost" content="parts" onPress={() => e.openComment(thread.id)}>
        <strong>{title}</strong>
        <span>{proposal?.summary ?? first.text}</span>
        <Attribution
          person={{ name: first.author === e.actor.id ? 'You' : e.personName(first.author) }}
          time={requireDefined(thread.entries.at(-1)).createdAt}
        />
        <span>
          {thread.entries.length - 1} replies
          {thread.state === 'resolved' ? ' · Resolved' : ''}
        </span>
      </Button>
    </div>
  );
}

/** The workbook's conversations matching `filter`, most recently active first. */
function conversationsOf(e: SpreadsheetEditorController, filter: string) {
  return (e.workbook.comments ?? [])
    .filter((thread) =>
      filter === 'mine'
        ? thread.entries.some((entry) => entry.author === e.actor.id)
        : filter === 'all' || thread.state === filter,
    )
    .sort((a, b) =>
      requireDefined(b.entries.at(-1)).createdAt.localeCompare(
        requireDefined(a.entries.at(-1)).createdAt,
      ),
    );
}

function CommentList({
  editor: e,
  filter,
  setFilter,
}: EditorProps & { filter: string; setFilter(value: string): void }) {
  const conversations = conversationsOf(e, filter);
  return (
    <div className="se-comments" aria-label="Workbook comments">
      <div className="se-comments__actions">
        <SelectField
          label="Comment filter"
          value={filter}
          onChange={(value) => setFilter(String(value))}
          options={FILTERS}
        />
        <Button
          startIcon={ChatBubblePlusGlyph}
          onPress={e.beginWorkbookComment}
          isDisabled={!e.canComment}
        >
          Add comment
        </Button>
      </div>
      {e.commentComposer && e.commentScope === 'workbook' && (
        <NewComment editor={e} label="Comment on workbook" />
      )}
      {!conversations.length && noComments(filter)}
      {conversations.map((thread) => (
        <ConversationSummary key={thread.id} editor={e} thread={thread} />
      ))}
    </div>
  );
}

function OpenThread({
  editor: e,
  threadId,
  filter,
}: EditorProps & { threadId: string; filter: string }) {
  const threads = (e.workbook.comments ?? []).filter((thread) => thread.id === threadId);
  return (
    <div className="se-comments" aria-label="Workbook comments">
      {e.commentComposer && (threadId === 'draft' || e.commentScope === 'workbook') && (
        <NewComment
          editor={e}
          label={
            e.commentScope === 'workbook'
              ? 'Comment on workbook'
              : `Comment on ${e.sheet.name}!${selectionAddress(e.selection)}`
          }
        />
      )}
      {!threads.length && threadId !== 'draft' && noComments(filter)}
      {threads.map((thread) => (
        <ThreadCard key={thread.id} editor={e} thread={thread} threadId={threadId} />
      ))}
    </div>
  );
}

export function CommentsPanel({ editor: e, threadId }: EditorProps & { threadId?: string }) {
  const [filter, setFilter] = useState('open');
  if (!threadId) return <CommentList editor={e} filter={filter} setFilter={setFilter} />;
  return <OpenThread editor={e} threadId={threadId} filter={filter} />;
}
