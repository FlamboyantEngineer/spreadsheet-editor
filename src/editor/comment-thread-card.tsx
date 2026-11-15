'use client';

import { Button } from '@flamboyantengineer/ui-components/button';
import { CommentComposer } from '@flamboyantengineer/ui-components/comment-composer';
import {
  CommentThread,
  CommentThreadReply,
} from '@flamboyantengineer/ui-components/comment-thread';
import { ReactionBar } from '@flamboyantengineer/ui-components/reaction-bar';
import { requireDefined } from '../model/require-defined.js';
import {
  type CommentEntry,
  deleteCommentEntry,
  editCommentEntry,
  replyToComment,
  setCommentState,
  type SpreadsheetComment,
  type SpreadsheetSuggestion,
  toggleCommentReaction,
} from '../model/review.js';
import { SuggestionDecisionButtons } from './suggestions-panel.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type ThreadProps = { editor: SpreadsheetEditorController; thread: SpreadsheetComment };

/** The suggestion a thread discusses, with its changes and decision buttons. */
function ThreadProposal({
  editor: e,
  proposal,
}: {
  editor: SpreadsheetEditorController;
  proposal: SpreadsheetSuggestion;
}) {
  return (
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
        <SuggestionDecisionButtons editor={e} suggestion={proposal} />
      )}
    </article>
  );
}

function PostReactions({ editor: e, thread, post }: ThreadProps & { post: CommentEntry }) {
  return (
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
                    { threadId: thread.id, postId: post.id },
                    { emoji, actorId: e.actor.id },
                  ),
                  'React to comment',
                ),
              )
          : undefined
      }
    />
  );
}

/** One post: its own author may edit or delete it; everyone who can comment may react. */
function ThreadPost({ editor: e, thread, post }: ThreadProps & { post: CommentEntry }) {
  const own = e.canComment && post.author === e.actor.id;
  const entry = { threadId: thread.id, postId: post.id };
  return (
    <CommentThreadReply
      author={{ name: e.personName(post.author), colorKey: post.author }}
      time={post.createdAt}
      text={post.text}
      isEdited={!!post.editedAt}
      onEdit={
        own
          ? (text) =>
              e.commit(
                editCommentEntry(e.workbook, entry, {
                  text,
                  actorId: e.actor.id,
                  time: new Date().toISOString(),
                }),
                'Edit comment',
              )
          : undefined
      }
      onDelete={
        own
          ? () => e.commit(deleteCommentEntry(e.workbook, entry, e.actor.id), 'Delete comment')
          : undefined
      }
      reactions={<PostReactions editor={e} thread={thread} post={post} />}
    />
  );
}

function ReplyComposer({ editor: e, thread, threadId }: ThreadProps & { threadId: string }) {
  const clearDraft = () => e.setReplyDrafts((drafts) => ({ ...drafts, [thread.id]: '' }));
  return (
    <CommentComposer
      autoFocus={!!threadId}
      value={e.replyDrafts[thread.id] ?? ''}
      onChange={(value) => e.setReplyDrafts((drafts) => ({ ...drafts, [thread.id]: value }))}
      onCancel={() => {
        clearDraft();
        if (threadId) e.setActiveCommentId(null);
      }}
      label="Reply to comment"
      submitLabel="Post reply"
      onSubmit={(text) => {
        e.commit(
          replyToComment(e.workbook, { id: thread.id, text }, e.stamp()),
          'Reply to comment',
        );
        clearDraft();
      }}
    />
  );
}

function CopyLinkButton({ editor: e, thread }: ThreadProps) {
  if (!e.props.getCommentLink) return null;
  return (
    <Button
      size="sm"
      onPress={() =>
        e.attempt(async () => {
          const text = requireDefined(e.props.getCommentLink)(thread.id);
          await (e.props.clipboard?.writeText(text) ?? navigator.clipboard.writeText(text));
        })
      }
    >
      Copy comment link
    </Button>
  );
}

/** The quoted range a thread is about: a button that selects it, or its deleted cells. */
function threadQuote({ editor: e, thread }: ThreadProps, sheetName: string | undefined) {
  if (thread.scope === 'workbook') return undefined;
  if (thread.orphaned) return `Deleted cells · ${thread.range}`;
  return (
    <Button onPress={() => e.openComment(thread.id)}>
      {sheetName}!{thread.range}
    </Button>
  );
}

/** An open thread, with the suggestion it discusses and its reply composer. */
export function ThreadCard(props: ThreadProps & { threadId: string }) {
  const { editor: e, thread } = props;
  const sheet = e.workbook.worksheets.find((sheet) => sheet.id === thread.sheetId);
  const proposal = e.workbook.suggestions?.find((item) => item.id === thread.suggestion);
  return (
    <div>
      <CopyLinkButton {...props} />
      {proposal && <ThreadProposal editor={e} proposal={proposal} />}
      <CommentThread
        key={thread.id}
        aria-label={
          thread.scope === 'workbook'
            ? 'Workbook comment thread'
            : `Comment thread ${sheet?.name ?? 'Deleted worksheet'}!${thread.range}`
        }
        quote={threadQuote(props, sheet?.name)}
        isResolved={thread.state === 'resolved'}
        isReadOnly={!e.canComment}
        onResolvedChange={(resolved) =>
          e.attempt(() =>
            e.commit(
              setCommentState(
                e.workbook,
                { id: thread.id, state: resolved ? 'resolved' : 'open' },
                e.stamp(),
              ),
              resolved ? 'Resolve comment' : 'Reopen comment',
            ),
          )
        }
        composer={thread.state === 'open' && e.canComment ? <ReplyComposer {...props} /> : null}
      >
        {thread.entries.map((post) => (
          <ThreadPost key={post.id} editor={e} thread={thread} post={post} />
        ))}
      </CommentThread>
    </div>
  );
}
