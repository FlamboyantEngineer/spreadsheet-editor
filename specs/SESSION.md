# Workbook session

## SE-SESSION-1 - Workbook session

A workbook session opens one stored workbook through the host's `workbooks` store and saves it against the revision it was edited from, as document-editor's store sessions do. It needs no framework or DOM: `createWorkbookSession(options)` lives in `./host`, and `useWorkbookSession(options)` in the main entry opens one for as long as a component is mounted. The editor takes it as its `session` prop in place of the `onSave`, `onFlush` and `autosaveMs` props it had before.

### SE-SESSION-1.1 - Opening

The session shall load the workbook (`workbooks.load`) and report `loading`, then `saved` at the stored revision, or `unavailable` with the host's error. While its session is opening, the editor renders a status in place of the grid (`[data-session="loading"]`), or the error (`[data-session="unavailable"]`); a null `session` prop also shows it as opening. When another save of the workbook lands (`workbooks.subscribe`) and the session holds no unsaved change, it loads and shows that save; otherwise it keeps the reader's changes.

`getSnapshot()` returns the workbook, status, draft offer and local revision from one update. Its identity stays stable between notifications. Server rendering shows the opening state for a configured session and renders local or controlled editors normally.

Implementation: `src/host/workbook-session.ts`, `src/editor/use-session-state.ts`, `src/editor/spreadsheet-editor.tsx`
Verification: `src/host/workbook-session.test.ts`, `src/editor/host-integration.test.tsx`

### SE-SESSION-1.2 - Causal coverage

Each accepted local change shall be recorded with its `WorkbookChangeContext.transactionId` and a local revision. A save covers the changes accepted before it started, sends their content with the stored revision as `baseRevision`, and on its answer marks exactly those as saved: a change accepted while it was in flight stays unsaved and saves next. A save asked for while another is in flight joins it, and saves again only when the first did not cover what was accepted before it was asked for. `acknowledge(transactionId)` marks a transaction and every earlier one as saved for a host that keeps changes by other means, such as an operation feed; an unknown, repeated or already covered id changes nothing. Comment and other annotation changes are transactions like any other. A remote sequence number never counts as saving a local change.

Autosave runs `autosaveDelay` milliseconds after the last change (1000 by default; `Infinity` turns it off), and when the `lifecycle` adapter reports the page hidden or left. An autosave that finds the store `unavailable` retries with backoff from half a second to four seconds, for up to 30 seconds; a save the reader asks for, and every other failure, stops at `error` with the host's message and keeps the changes unsaved. A terminal save failure remains at `error` through newer edits, with its message and unsaved changes retained; autosave and lifecycle events do not restart it until the reader explicitly saves or flushes. A save refused as `conflict` stops at `conflict`: the editor shows a banner (`.se-editor__conflict`) to keep the reader's workbook over the stored one (`keepMine`, which saves against the stored revision) or take the stored one and drop the reader's changes (`takeStored`). Saving does not resume until one is chosen.

Taking the stored copy waits for a save already in progress. A save requested while that choice is pending waits for its outcome and submits only changes still pending afterward. A local edit during the reload keeps the local workbook and rejects the choice.

Implementation: `src/host/workbook-session.ts`, `src/host/transaction-receipts.ts`
Verification: `src/host/workbook-session.test.ts`, `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`

### SE-SESSION-1.3 - Drafts

Given a `drafts` adapter, the session shall keep unsaved content as a draft `draftDelay` milliseconds after the last change (300 by default), with the revision it was edited from, and clear it once a save covers every change. A session that opens a workbook with a kept draft offers it (`draft()`), marked `stale` when the stored revision moved on since; the editor shows a banner (`.se-editor__draft`) to restore it as unsaved changes (`restoreDraft`) or discard it (`discardDraft`).

Draft puts and clears run in order within a session. A failed Flush retains pending draft work. A conflicting remote revision is the target for an explicit overwrite and leaves the recovery draft's edit base unchanged. Restoring a stale draft uses its recorded edit base for the next save.

Implementation: `src/host/workbook-session.ts`, `src/editor/spreadsheet-editor.tsx`
Verification: `src/host/workbook-session.test.ts`

### SE-SESSION-1.4 - The editor's saving

With a session, the editor shows the session's workbook unless the host supplies `value`. A host that supplies both `value` and `session` uses `acceptChange(next, context)` to synchronously return the accepted workbook. The session records that exact workbook, including rebased changes. The editor waits for a configured session to load before showing editable content. Save (button, menu and Ctrl+S) commits the active draft and saves through the session. The handle's `flush` commits the draft and waits until every change accepted before that call is saved. `flushing` stays true until the last active Flush completes. The status and handle report the same session save phase and failures. Causal acknowledgements cover accepted transactions; an uncommitted draft stays dirty. Opening or restoring a saved version protects newer edits and drafts during the request (SE-ADAPTERS-5.1). Without a session the editor has no Save control.

Replacing the session object invalidates the previous editor handle and starts a fresh editing instance, including when both sessions address the same workbook ID. `useWorkbookSession` exposes only the session for its current store and configuration, and returns `null` while it switches sessions.

Implementation: `src/editor/use-editor.ts`, `src/editor/editor-handle.ts`, `src/editor/use-session-state.ts`, `src/editor/use-session-actions.ts`, `src/editor/use-workbook-session.ts`, `src/editor/editor-toolbar.tsx`
Verification: `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`, `src/editor/editor-lifecycle.test.tsx`, `e2e/host-integration.spec.ts`, `e2e/full-editor.spec.ts`
