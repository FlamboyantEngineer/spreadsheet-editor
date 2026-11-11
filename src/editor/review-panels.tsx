'use client';

import { useState } from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { Spinner } from '@flamboyantengineer/ui-components/progress-bar';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import {
  CommentThread,
  CommentThreadReply,
} from '@flamboyantengineer/ui-components/comment-thread';
import { CommentComposer } from '@flamboyantengineer/ui-components/comment-composer';
import { ReactionBar } from '@flamboyantengineer/ui-components/reaction-bar';
import { ContextMenuTrigger } from '@flamboyantengineer/ui-components/context-menu';
import { Menu } from '@flamboyantengineer/ui-components/menu';
import { Attribution } from '@flamboyantengineer/ui-components/attribution';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import { selectionAddress } from '../model/editing.js';
import {
  editCommentPost,
  deleteCommentPost,
  replyToComment,
  toggleCommentReaction,
  updateComment,
} from '../model/review.js';
import type { SpreadsheetEditorController } from './use-editor.js';
import { editorCommandMetadata } from './editor-commands.js';

export function CommentsPanel({
  editor: e,
  threadId,
}: {
  editor: SpreadsheetEditorController;
  threadId?: string;
}) {
  const [filter, setFilter] = useState('open');
  const threads = (e.workbook.comments ?? []).filter((thread) =>
    threadId
      ? thread.id === threadId
      : !thread.suggestionId &&
        (filter === 'mine'
          ? thread.posts.some((post) => post.author.id === e.actor.id)
          : filter === 'all' || thread.resolved === (filter === 'resolved')),
  );
  if (!threadId) {
    const conversations = (e.workbook.comments ?? [])
      .filter((thread) =>
        filter === 'mine'
          ? thread.posts.some((post) => post.author.id === e.actor.id)
          : filter === 'all' || thread.resolved === (filter === 'resolved'),
      )
      .sort((a, b) => b.posts.at(-1)!.createdAt.localeCompare(a.posts.at(-1)!.createdAt));
    return (
      <div className="se-comments" aria-label="Workbook comments">
        <div className="se-comments__actions">
          <SelectField
            label="Comment filter"
            value={filter}
            onChange={(value) => setFilter(String(value))}
            options={[
              { id: 'open', label: 'Open comments' },
              { id: 'resolved', label: 'Resolved comments' },
              { id: 'all', label: 'All comments' },
              { id: 'mine', label: 'My conversations' },
            ]}
          />
          <Button onPress={e.beginWorkbookComment} isDisabled={!e.canComment}>
            Add comment
          </Button>
        </div>
        {e.commentComposer && e.commentScope === 'workbook' && (
          <CommentComposer
            label="Comment on workbook"
            submitLabel="Post comment"
            onSubmit={e.postComment}
            onCancel={() => {
              e.setCommentComposer(false);
              e.setActiveCommentId(null);
            }}
            autoFocus
            isDisabled={!e.canComment}
          />
        )}
        {!conversations.length && <p>No {filter === 'all' ? '' : filter + ' '}comments.</p>}
        {conversations.map((thread) => {
          const proposal = e.workbook.suggestions?.find((item) => item.id === thread.suggestionId);
          const sheet = e.workbook.worksheets.find((sheet) => sheet.id === thread.sheetId);
          return (
            <div key={thread.id} className="se-comments__conversation">
              <Button variant="ghost" content="parts" onPress={() => e.openComment(thread.id)}>
                <strong>
                  {proposal
                    ? `Suggestion · ${proposal.status}`
                    : thread.scope === 'workbook'
                      ? 'Workbook comment'
                      : `${sheet?.name ?? 'Deleted worksheet'}!${thread.range}`}
                </strong>
                <span>{proposal?.summary ?? thread.posts[0].text}</span>
                <Attribution
                  person={{
                    name:
                      thread.posts[0].author.id === e.actor.id
                        ? 'You'
                        : thread.posts[0].author.name,
                  }}
                  time={thread.posts.at(-1)!.createdAt}
                />
                <span>
                  {thread.posts.length - 1} replies{thread.resolved ? ' · Resolved' : ''}
                </span>
              </Button>
            </div>
          );
        })}
      </div>
    );
  }
  return (
    <div className="se-comments" aria-label="Workbook comments">
      {!threadId && (
        <div className="se-comments__actions">
          <SelectField
            label="Comment filter"
            value={filter}
            onChange={(value) => setFilter(String(value))}
            options={[
              { id: 'open', label: 'Open comments' },
              { id: 'resolved', label: 'Resolved comments' },
              { id: 'all', label: 'All comments' },
              { id: 'mine', label: 'My conversations' },
            ]}
          />
          <Button onPress={e.beginWorkbookComment} isDisabled={!e.canComment}>
            Add comment
          </Button>
        </div>
      )}
      {e.commentComposer && (threadId === 'draft' || e.commentScope === 'workbook') && (
        <CommentComposer
          label={
            e.commentScope === 'workbook'
              ? 'Comment on workbook'
              : `Comment on ${e.sheet.name}!${selectionAddress(e.selection)}`
          }
          submitLabel="Post comment"
          onSubmit={e.postComment}
          onCancel={() => {
            e.setCommentComposer(false);
            e.setActiveCommentId(null);
          }}
          autoFocus
          isDisabled={!e.canComment}
        />
      )}
      {!threads.length && threadId !== 'draft' && (
        <p>No {filter === 'all' ? '' : filter + ' '}comments.</p>
      )}
      {threads.map((thread) => {
        const sheet = e.workbook.worksheets.find((sheet) => sheet.id === thread.sheetId);
        const proposal = e.workbook.suggestions?.find((item) => item.id === thread.suggestionId);
        return (
          <div key={thread.id}>
            {e.props.getCommentLink && (
              <Button
                size="sm"
                onPress={() =>
                  e.attempt(async () => {
                    const text = e.props.getCommentLink!(thread.id);
                    await (e.props.clipboard?.writeText(text) ??
                      navigator.clipboard.writeText(text));
                  })
                }
              >
                Copy comment link
              </Button>
            )}
            {proposal && (
              <article className="se-suggestion" aria-label={`Suggestion: ${proposal.summary}`}>
                <h3>{proposal.summary}</h3>
                <p>{proposal.status}</p>
                {proposal.changes.map((change) => (
                  <p key={change.sheetId + change.address}>
                    {change.address}: <del>{String(change.before ?? '(empty)')}</del> →{' '}
                    <ins>{String(change.after ?? '(empty)')}</ins>
                  </p>
                ))}
                {proposal.status === 'pending' && (
                  <div className="se-suggestion__actions">
                    <Button
                      isDisabled={!e.canDecide || proposal.outdated}
                      onPress={() => e.attempt(() => e.decide(proposal.id, 'accepted'))}
                    >
                      Accept suggestion
                    </Button>
                    <Button
                      isDisabled={!e.canDecide}
                      onPress={() => e.attempt(() => e.decide(proposal.id, 'rejected'))}
                    >
                      Reject suggestion
                    </Button>
                  </div>
                )}
              </article>
            )}
            <CommentThread
              key={thread.id}
              aria-label={
                thread.scope === 'workbook'
                  ? 'Workbook comment thread'
                  : `Comment thread ${sheet?.name ?? 'Deleted worksheet'}!${thread.range}`
              }
              quote={
                thread.scope === 'workbook' ? undefined : thread.orphaned ? (
                  `Deleted cells · ${thread.range}`
                ) : (
                  <Button onPress={() => e.openComment(thread.id)}>
                    {sheet?.name}!{thread.range}
                  </Button>
                )
              }
              isResolved={thread.resolved}
              isReadOnly={!e.canComment}
              onResolvedChange={(resolved) =>
                e.attempt(() =>
                  e.commit(
                    updateComment(e.workbook, thread.id, (thread) => {
                      thread.resolved = resolved;
                    }),
                    resolved ? 'Resolve comment' : 'Reopen comment',
                  ),
                )
              }
              composer={
                !thread.resolved && e.canComment ? (
                  <CommentComposer
                    autoFocus={!!threadId}
                    value={e.replyDrafts[thread.id] ?? ''}
                    onChange={(value) =>
                      e.setReplyDrafts((drafts) => ({ ...drafts, [thread.id]: value }))
                    }
                    onCancel={() => {
                      e.setReplyDrafts((drafts) => ({ ...drafts, [thread.id]: '' }));
                      if (threadId) e.setActiveCommentId(null);
                    }}
                    label="Reply to comment"
                    submitLabel="Post reply"
                    onSubmit={(text) => {
                      e.commit(
                        replyToComment(e.workbook, thread.id, text, e.stamp()),
                        'Reply to comment',
                      );
                      e.setReplyDrafts((drafts) => ({ ...drafts, [thread.id]: '' }));
                    }}
                  />
                ) : null
              }
            >
              {thread.posts.map((post) => (
                <CommentThreadReply
                  key={post.id}
                  author={{ name: post.author.name, colorKey: post.author.id }}
                  time={post.createdAt}
                  text={post.text}
                  isEdited={!!post.updatedAt}
                  onEdit={
                    e.canComment && post.author.id === e.actor.id
                      ? (text) =>
                          e.commit(
                            editCommentPost(
                              e.workbook,
                              thread.id,
                              post.id,
                              text,
                              e.actor.id,
                              new Date().toISOString(),
                            ),
                            'Edit comment',
                          )
                      : undefined
                  }
                  onDelete={
                    e.canComment && post.author.id === e.actor.id
                      ? () =>
                          e.commit(
                            deleteCommentPost(e.workbook, thread.id, post.id, e.actor.id),
                            'Delete comment',
                          )
                      : undefined
                  }
                  reactions={
                    <ReactionBar
                      aria-label="Comment reactions"
                      reactions={Object.entries(post.reactions ?? {}).map(([emoji, people]) => ({
                        emoji,
                        count: people.length,
                        isReacted: people.includes(e.actor.id),
                        people: people.map((id) =>
                          id === e.actor.id
                            ? e.actor.name
                            : (thread.posts.find((post) => post.author.id === id)?.author.name ??
                              id),
                        ),
                      }))}
                      onToggle={
                        e.canComment
                          ? (emoji) =>
                              e.attempt(() =>
                                e.commit(
                                  toggleCommentReaction(
                                    e.workbook,
                                    thread.id,
                                    post.id,
                                    emoji,
                                    e.actor.id,
                                  ),
                                  'React to comment',
                                ),
                              )
                          : undefined
                      }
                    />
                  }
                />
              ))}
            </CommentThread>
          </div>
        );
      })}
    </div>
  );
}

export function SuggestionsPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  const [filter, setFilter] = useState('pending');
  const suggestions = (e.workbook.suggestions ?? []).filter(
    (item) => filter === 'all' || item.status === filter,
  );
  return (
    <div className="se-suggestions" aria-label="Workbook suggestions">
      <SelectField
        label="Suggestion filter"
        value={filter}
        onChange={(value) => setFilter(String(value))}
        options={['pending', 'accepted', 'rejected', 'all'].map((id) => ({
          id,
          label: id[0].toUpperCase() + id.slice(1),
        }))}
      />
      <p>
        In Reviewing mode, cell edits and formatting become suggestions. Canonical values change
        after acceptance.
      </p>
      {!suggestions.length && <p>No {filter === 'all' ? '' : filter + ' '}suggestions.</p>}
      {suggestions.map((suggestion) => (
        <article
          key={suggestion.id}
          className="se-suggestion"
          aria-label={`Suggestion: ${suggestion.summary}`}
          data-status={suggestion.status}
        >
          <h3>{suggestion.summary}</h3>
          <Attribution person={{ name: suggestion.author.name }} time={suggestion.createdAt} />
          <p className="se-suggestion__status">
            {suggestion.status}
            {suggestion.outdated ? ' · deleted cells' : ''}
          </p>
          {suggestion.structure && (
            <p>Worksheet structure change, including any affected formulas.</p>
          )}
          {suggestion.changes.map((change) => (
            <div key={`${change.sheetId}:${change.address}`} className="se-suggestion__change">
              <Button
                onPress={() => {
                  const cell = parseSpreadsheetRange(change.address)!;
                  e.select(
                    { x1: cell.start.x, x2: cell.end.x, y1: cell.start.y, y2: cell.end.y },
                    change.sheetId,
                  );
                }}
                isDisabled={suggestion.outdated}
              >
                {e.workbook.worksheets.find((sheet) => sheet.id === change.sheetId)?.name}!
                {change.address}
              </Button>
              {change.before !== change.after ? (
                <p>
                  <del>{String(change.before ?? '(empty)')}</del>
                  <span aria-hidden="true"> → </span>
                  <ins>{String(change.after ?? '(empty)')}</ins>
                </p>
              ) : (
                <p>Cell formatting change</p>
              )}
            </div>
          ))}
          {suggestion.commentThreadId && (
            <Button size="sm" onPress={() => e.openSuggestion(suggestion.id)}>
              Open discussion
            </Button>
          )}
          {suggestion.status === 'pending' && (
            <div className="se-suggestion__actions">
              <Button
                isDisabled={!e.canDecide || suggestion.outdated}
                onPress={() => e.attempt(() => e.decide(suggestion.id, 'accepted'))}
              >
                Accept suggestion
              </Button>
              <Button
                isDisabled={!e.canDecide}
                onPress={() => e.attempt(() => e.decide(suggestion.id, 'rejected'))}
              >
                Reject suggestion
              </Button>
            </div>
          )}
          {suggestion.decidedBy && (
            <Attribution
              person={{ name: suggestion.decidedBy.name }}
              time={suggestion.decidedAt}
              template="Decided {time} by {person}"
            />
          )}
        </article>
      ))}
    </div>
  );
}

export function PropertiesPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  const [title, setTitle] = useState(e.workbook.name ?? 'Untitled spreadsheet');
  return (
    <div className="se-properties">
      <h3>Workbook properties</h3>
      <TextField label="Workbook title" value={title} onChange={setTitle} isReadOnly={!e.canEdit} />
      <Button
        isDisabled={!e.canEdit || !title.trim()}
        onPress={() => e.commit({ ...e.workbook, name: title.trim() }, 'Rename workbook')}
      >
        Update title
      </Button>
      <dl>
        <dt>Worksheet</dt>
        <dd>{e.sheet.name}</dd>
        <dt>Rows</dt>
        <dd>{e.sheet.data.length}</dd>
        <dt>Columns</dt>
        <dd>{Math.max(0, ...e.sheet.data.map((row) => row.length))}</dd>
        <dt>Selection</dt>
        <dd>{selectionAddress(e.selection)}</dd>
        <dt>Calculations</dt>
        <dd>{e.props.calculation?.mode ?? 'client'}</dd>
      </dl>
    </div>
  );
}

export function HistoryPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  const stored = e.props.history;
  if (stored)
    return (
      <div className="se-history" aria-label="Saved workbook history">
        <h3>Version history</h3>
        {stored.onRefresh && (
          <Button size="sm" onPress={() => e.attempt(() => stored.onRefresh!())}>
            Refresh history
          </Button>
        )}
        {stored.loading && <Spinner label="Loading history" />}
        {stored.error && <p role="alert">{stored.error}</p>}
        {!stored.loading && !stored.entries.length && <p>No saved versions.</p>}
        <ol>
          {[...stored.entries]
            .sort((a, b) => Date.parse(b.timestamp) - Date.parse(a.timestamp))
            .map((entry) => (
              <li key={entry.id}>
                <ContextMenuTrigger>
                  <Button variant="ghost" content="parts">
                    <Attribution person={{ name: entry.author.name }} time={entry.timestamp} />
                    {entry.state && <span>State: {entry.state}</span>}
                    {entry.revision && <span>Revision: {entry.revision}</span>}
                  </Button>
                  <Menu
                    aria-label="Snapshot actions"
                    items={[
                      {
                        id: 'open',
                        label: 'View historical version',
                        ...editorCommandMetadata('open-history'),
                        onAction: () => e.attempt(() => e.openHistorical(entry)),
                      },
                    ]}
                  />
                </ContextMenuTrigger>
              </li>
            ))}
        </ol>
      </div>
    );
  return (
    <div className="se-history">
      <h3>Editing history</h3>
      <p>Recent changes in this editing session.</p>
      {!e.history.past.length && <p>No changes in this session.</p>}
      <ol>
        {[...e.history.past].reverse().map((entry, index) => (
          <li key={index}>{entry.label}</li>
        ))}
      </ol>
      <Button isDisabled={!e.canWrite || !e.history.past.length} onPress={e.undo}>
        Undo last change
      </Button>
      <Button isDisabled={!e.canWrite || !e.history.future.length} onPress={e.redo}>
        Redo last change
      </Button>
    </div>
  );
}
