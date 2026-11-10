'use client';

import { useState } from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import {
  CommentThread,
  CommentThreadReply,
} from '@flamboyantengineer/ui-components/comment-thread';
import { CommentComposer } from '@flamboyantengineer/ui-components/comment-composer';
import { ReactionBar } from '@flamboyantengineer/ui-components/reaction-bar';
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

export function CommentsPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  const [filter, setFilter] = useState('open');
  const threads = (e.workbook.comments ?? []).filter(
    (thread) => filter === 'all' || thread.resolved === (filter === 'resolved'),
  );
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
          ]}
        />
        <Button onPress={e.beginComment} isDisabled={!e.canComment}>
          Add comment
        </Button>
      </div>
      {e.commentComposer && (
        <CommentComposer
          label={`Comment on ${e.sheet.name}!${selectionAddress(e.selection)}`}
          submitLabel="Post comment"
          onSubmit={e.postComment}
          onCancel={() => e.setCommentComposer(false)}
          autoFocus
          isDisabled={!e.canComment}
        />
      )}
      {!threads.length && <p>No {filter === 'all' ? '' : filter + ' '}comments.</p>}
      {threads.map((thread) => {
        const sheet = e.workbook.worksheets.find((sheet) => sheet.id === thread.sheetId),
          range = parseSpreadsheetRange(thread.range)!;
        return (
          <CommentThread
            key={thread.id}
            aria-label={`Comment thread ${sheet?.name ?? 'Deleted worksheet'}!${thread.range}`}
            quote={
              thread.orphaned ? (
                `Deleted cells · ${thread.range}`
              ) : (
                <Button
                  onPress={() =>
                    e.select(
                      { x1: range.start.x, y1: range.start.y, x2: range.end.x, y2: range.end.y },
                      thread.sheetId,
                    )
                  }
                >
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
                  label="Reply to comment"
                  submitLabel="Post reply"
                  onSubmit={(text) =>
                    e.commit(
                      replyToComment(e.workbook, thread.id, text, e.stamp()),
                      'Reply to comment',
                    )
                  }
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
                          : (thread.posts.find((post) => post.author.id === id)?.author.name ?? id),
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
