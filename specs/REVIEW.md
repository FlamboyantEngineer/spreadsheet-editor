# Comments and suggestions

## SE-REVIEW-1 - Workbook comments

### SE-REVIEW-1.1 - Workbook comments

Version-1 workbooks may carry comment threads. Their records take the document editor's comment shape, so a host that embeds both editors stores and lists comments the same way:

- A thread has a stable `id`, a range anchor (`sheetId`, `range`) or `scope: 'workbook'`, a `state` of `'open'` or `'resolved'`, and its `entries`, the opening comment first and replies after it, oldest first.
- An entry has an `id`, an `author` that is a person id, a `createdAt` time, its plain `text`, an `editedAt` time once its text changed, and emoji `reactions` that map each emoji to the ids of the people who reacted.
- A resolved thread records `resolvedBy` (a person id) and `resolvedAt`; reopening clears both.
- A thread that discusses a suggestion names it in `suggestion`. Once the suggestion is decided, the thread is resolved and records `decision` as `{ decision: 'accept' | 'reject', suggestion, by, at }`.

Comments stay in the workbook JSON; there is no separate comment store. Records store person ids only. The editor names people through the host's `identity` adapter (`resolvePeople`), and an id the adapter does not know is shown as the id itself (SE-REVIEW-1.1). Suggestions keep their recorded author name, since they become cell provenance on acceptance.

The UI shall use shared CommentThread, CommentThreadReply, CommentComposer, Attribution and ReactionBar components. Threads support creation, replies, author-only edit/delete, reactions, resolve/reopen, open/resolved/all filtering, navigation and cell markers. Invalid imported ids, states, dates, decisions or reaction lists are rejected, and so are threads in the earlier `posts`/`resolved` shape.

The workbook and its annotations are immutable inputs; mutations return independent snapshots through `onChange`. Save/reload, JSON export and live insertion retain the data. Storage and server authorization belong to the host; no application comment API, fixed URL or browser storage is required.

Implementation: `src/model/review.ts`, `src/editor/review-panels.tsx`, `src/editor/use-person-names.ts`, `src/engine/cell-renderer.ts`.

Verification: `src/model/review.test.ts`, `src/editor/host-integration.test.tsx`, `e2e/full-editor.spec.ts`.

## SE-REVIEW-2 - Proposals and decisions

### SE-REVIEW-2.1 - Proposals and decisions

Reviewing shall stage cell, paste, formula and formatting changes as suggestions, preserving canonical data and calculations until acceptance while privately calculating the local reviewer’s preview. Each proposal carries its author/time, original/proposed values or structural snapshots and a pending/accepted/rejected status. Decisions retain actor/time metadata. Rejection leaves data unchanged; acceptance verifies every affected baseline before applying, and reports a conflict without partial writes if data or formatting diverged. Structural acceptance moves current comments, including comments added after the proposal, by the recorded structural operation.

Deleted anchors shall mark proposals outdated and disable acceptance. Already decided proposals cannot be applied again. The side pane shows before/after values or formatting/structure summaries, filters by decision state and navigates to affected cells. Pending proposals mark their cells. Normal document tables project canonical calculated values; live spreadsheet insertions retain pending review state.

Implementation: `src/model/review.ts`, `src/model/structure.ts`, `src/editor/use-editor.ts`, `src/editor/review-panels.tsx`, `src/engine/workbook-grid.tsx`.

Verification: `src/model/review.test.ts`, `src/model/structure.test.ts`, `e2e/full-editor.spec.ts` (SE-REVIEW-2.1).

## SE-REVIEW-5 - Permissions and transport

### SE-REVIEW-5.1 - Permissions and transport

Reading and host read-only mode shall disable comment mutation, review decisions and canonical edits in the toolbar, formula field, context menu, panels and native grid paths. Host flags independently restrict edit/comment/suggest/decide controls; the host must enforce the same policy when persisting changes. Server calculation shall accept annotation-bearing workbooks and return exact results after accepted proposals without discarding annotations. Separate editor instances retain independent review state.

Implementation: `src/editor/contracts.ts`, `src/editor/use-editor.ts`, `src/engine/workbook-grid.tsx`, `src/model/workbook.ts`.

Verification: `e2e/full-editor.spec.ts` (SE-REVIEW-5.1, SE-REVIEW-7.1), existing independent-embedding browser test.

## SE-REVIEW-7 - Conversations and capabilities

### SE-REVIEW-7.1 - Conversations and capabilities

Ordinary comments support range or workbook scope; a new range thread must not duplicate an overlapping unresolved ordinary comment. The shared Popover renders cell discussions; the Comments pane provides compact newest-first comment/suggestion navigation with status and mine filters. `initialCommentId` and `getCommentLink` connect host deep links without fixed package URLs. Replies share drafts between views; cancel clears them. All cells in a comment range receive presentation markers; suggestion discussions do not count as ordinary comment corners.

A proposal creates one linked discussion. Same-author pending proposals with identical changed targets coalesce, preserving the original baseline. Authoring displays canonical cells; Reviewing displays only the current actor’s pending overlay. Acceptance/rejection resolves the linked thread with a decision and retains original-author provenance for accepted values. Structural proposal serialization remains readable for earlier version-1 workbooks, but the UI requires Authoring for structure, resize, merge, Save and decisions. Review-only host permissions clamp the mode accordingly. Server authorization remains the host’s responsibility.

Built-in comment entries are plain text. Host organizational mentions, group expansion, notifications, rich composers and centralized comment APIs remain integrations; hosts may replace the Comments pane.

Implementation: `src/model/review.ts`, `src/editor/review-panels.tsx`, `src/editor/use-comments.ts`.

Verification: `src/model/review.test.ts`, `src/editor/editor-lifecycle.test.tsx`, `e2e/full-editor.spec.ts`.

## SE-REVIEW-8 - Canonical input baselines

### SE-REVIEW-8.1 - Canonical input baselines

Suggested cells optionally record `beforeInputValue` and `afterInputValue`. These use the same bounded numeric metadata validation as canonical authored inputs and cannot override formulas. A canonical-only change is a cell proposal rather than a structural proposal. Coalescing retains its original canonical baseline, local preview uses the proposed override, and acceptance rejects a changed canonical baseline before applying any proposed cell. Replacement removes an override when the proposal contains none; formatting-only suggestions retain it. Structural changes move available suggestion anchors and canonical metadata with their cells, including decided discussions, while deleted anchors become unavailable.

Formats in authored cells and suggestion baselines share ingress validation, including supported locale/currency metadata.

Implementation: `src/model/review.ts`, `src/model/cell-format.ts`, `src/model/structure.ts`.

Verification: `src/model/host-model.test.ts` (SE-HOST-MODEL-3.1), existing `src/model/review.test.ts`.
