'use client';

import { Attribution } from '@flamboyantengineer/ui-components/attribution';
import { Button } from '@flamboyantengineer/ui-components/button';
import { CommentComposer } from '@flamboyantengineer/ui-components/comment-composer';
import {
  CommentThread,
  CommentThreadReply,
} from '@flamboyantengineer/ui-components/comment-thread';
import { ContextMenuTrigger } from '@flamboyantengineer/ui-components/context-menu';
import {
  ChatBubbleCheckGlyph,
  ChatBubblePlusGlyph,
  ChatBubbleXMarkGlyph,
} from '@flamboyantengineer/ui-components/core';
import { InlineError } from '@flamboyantengineer/ui-components/inline-error';
import { Menu } from '@flamboyantengineer/ui-components/menu';
import { Spinner } from '@flamboyantengineer/ui-components/progress-bar';
import { ReactionBar } from '@flamboyantengineer/ui-components/reaction-bar';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { useState } from 'react';
import { selectionAddress } from '../model/editing.js';
import { requireDefined } from '../model/require-defined.js';
import {
  deleteCommentEntry,
  editCommentEntry,
  replyToComment,
  setCommentState,
  toggleCommentReaction,
} from '../model/review.js';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import { editorCommandMetadata } from './editor-commands.js';
import type { SpreadsheetEditorController } from './use-editor.js';

export function CommentsPanel({
  editor: e,
  threadId,
}: {
  editor: SpreadsheetEditorController;
  threadId?: string;
}) {
  const [filter, setFilter] = useState('open');
  if (!threadId) {
    const conversations = (e.workbook.comments ?? [])
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
          <Button
            startIcon={ChatBubblePlusGlyph}
            onPress={e.beginWorkbookComment}
            isDisabled={!e.canComment}
          >
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
          const proposal = e.workbook.suggestions?.find((item) => item.id === thread.suggestion);
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
                <span>{proposal?.summary ?? thread.entries[0].text}</span>
                <Attribution
                  person={{
                    name:
                      thread.entries[0].author === e.actor.id
                        ? 'You'
                        : e.personName(thread.entries[0].author),
                  }}
                  time={requireDefined(thread.entries.at(-1)).createdAt}
                />
                <span>
                  {thread.entries.length - 1} replies
                  {thread.state === 'resolved' ? ' · Resolved' : ''}
                </span>
              </Button>
            </div>
          );
        })}
      </div>
    );
  }
  const threads = (e.workbook.comments ?? []).filter((thread) => thread.id === threadId);
  return (
    <div className="se-comments" aria-label="Workbook comments">
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
        const proposal = e.workbook.suggestions?.find((item) => item.id === thread.suggestion);
        return (
          <div key={thread.id}>
            {e.props.getCommentLink && (
              <Button
                size="sm"
                onPress={() =>
                  e.attempt(async () => {
                    const text = requireDefined(e.props.getCommentLink)(thread.id);
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
                      startIcon={ChatBubbleCheckGlyph}
                      isDisabled={!e.canDecide || proposal.outdated}
                      onPress={() => e.attempt(() => e.decide(proposal.id, 'accepted'))}
                    >
                      Accept suggestion
                    </Button>
                    <Button
                      startIcon={ChatBubbleXMarkGlyph}
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
              isResolved={thread.state === 'resolved'}
              isReadOnly={!e.canComment}
              onResolvedChange={(resolved) =>
                e.attempt(() =>
                  e.commit(
                    setCommentState(
                      e.workbook,
                      thread.id,
                      resolved ? 'resolved' : 'open',
                      e.stamp(),
                    ),
                    resolved ? 'Resolve comment' : 'Reopen comment',
                  ),
                )
              }
              composer={
                thread.state === 'open' && e.canComment ? (
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
              {thread.entries.map((post) => (
                <CommentThreadReply
                  key={post.id}
                  author={{ name: e.personName(post.author), colorKey: post.author }}
                  time={post.createdAt}
                  text={post.text}
                  isEdited={!!post.editedAt}
                  onEdit={
                    e.canComment && post.author === e.actor.id
                      ? (text) =>
                          e.commit(
                            editCommentEntry(
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
                    e.canComment && post.author === e.actor.id
                      ? () =>
                          e.commit(
                            deleteCommentEntry(e.workbook, thread.id, post.id, e.actor.id),
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
                        people: people.map(e.personName),
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
                  const cell = requireDefined(parseSpreadsheetRange(change.address));
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
                startIcon={ChatBubbleCheckGlyph}
                isDisabled={!e.canDecide || suggestion.outdated}
                onPress={() => e.attempt(() => e.decide(suggestion.id, 'accepted'))}
              >
                Accept suggestion
              </Button>
              <Button
                startIcon={ChatBubbleXMarkGlyph}
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
  const [titleDraft, setTitleDraft] = useState<string | null>(null);
  const title = titleDraft ?? e.workbook.name ?? 'Untitled spreadsheet';
  return (
    <div className="se-properties">
      <h3>Workbook properties</h3>
      <TextField
        label="Workbook title"
        value={title}
        onChange={setTitleDraft}
        isReadOnly={!e.canEdit}
      />
      <Button
        isDisabled={!e.canEdit || !title.trim()}
        onPress={() =>
          e.attempt(() => {
            e.commit({ ...e.workbook, name: title.trim() }, 'Rename workbook');
            setTitleDraft(null);
          })
        }
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
  if (e.props.versions && e.props.workbookId !== undefined) return <VersionsPanel editor={e} />;
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
      <Button isDisabled={!e.canWrite || !e.canUndo} onPress={e.undo}>
        Undo last change
      </Button>
      <Button isDisabled={!e.canWrite || !e.canRedo} onPress={e.redo}>
        Redo last change
      </Button>
    </div>
  );
}

/** The workbook's saved versions, from the host's version store (SE-ADAPTERS-5.1). */
function VersionsPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  const { versions, loading, error, refresh } = e.versionList;
  const [label, setLabel] = useState('');
  const [renaming, setRenaming] = useState<{ revision: string; label: string } | null>(null);
  return (
    <div className="se-history" aria-label="Saved workbook history">
      <h3>Version history</h3>
      {e.canSaveVersion && (
        <form
          className="se-history__save"
          onSubmit={(event) => {
            event.preventDefault();
            void e.saveVersion(label.trim() || null).then((saved) => saved && setLabel(''));
          }}
        >
          <TextField label="Version name" value={label} onChange={setLabel} />
          <Button type="submit" size="sm">
            Save version
          </Button>
        </form>
      )}
      <Button size="sm" onPress={refresh}>
        Refresh history
      </Button>
      {loading && <Spinner label="Loading history" />}
      {error && <InlineError>{error}</InlineError>}
      {!loading && !error && !versions.length && <p>No saved versions.</p>}
      <ol>
        {versions.map((version) => (
          <li
            key={version.revision}
            data-current={e.viewedVersion?.revision === version.revision || undefined}
          >
            {renaming?.revision === version.revision ? (
              <form
                className="se-history__rename"
                onSubmit={(event) => {
                  event.preventDefault();
                  void e
                    .renameVersion(version, renaming.label.trim() || null)
                    .then((renamed) => renamed && setRenaming(null));
                }}
              >
                <TextField
                  label="Version name"
                  value={renaming.label}
                  onChange={(value) => setRenaming({ revision: version.revision, label: value })}
                  autoFocus
                />
                <Button type="submit" size="sm">
                  Rename
                </Button>
                <Button size="sm" variant="ghost" onPress={() => setRenaming(null)}>
                  Cancel
                </Button>
              </form>
            ) : (
              <ContextMenuTrigger>
                <Button variant="ghost" content="parts">
                  <Attribution
                    person={{ name: version.author ? e.personName(version.author) : 'Unknown' }}
                    time={version.savedAt}
                  />
                  <span>{version.label ?? 'Saved automatically'}</span>
                </Button>
                <Menu
                  aria-label="Version actions"
                  items={[
                    {
                      id: 'open',
                      label: 'View this version',
                      ...editorCommandMetadata('open-history'),
                      onAction: () => void e.openVersion(version),
                    },
                    ...(e.canRenameVersion
                      ? [
                          {
                            id: 'rename',
                            label: 'Rename version',
                            onAction: () =>
                              setRenaming({
                                revision: version.revision,
                                label: version.label ?? '',
                              }),
                          },
                        ]
                      : []),
                    ...(e.canRestoreVersion
                      ? [
                          {
                            id: 'restore',
                            label: 'Restore this version',
                            onAction: () => void e.restoreVersion(version),
                          },
                        ]
                      : []),
                  ]}
                />
              </ContextMenuTrigger>
            )}
          </li>
        ))}
      </ol>
    </div>
  );
}
