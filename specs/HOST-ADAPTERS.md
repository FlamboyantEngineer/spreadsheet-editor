# Host adapters

The editor reaches what lies outside it through adapters the host passes in: who the reader is, what they may do, where workbooks and their versions are stored, where the other people in a workbook are, and the page's lifecycle. The adapters follow the document editor's host adapters by name and shape, so a host that embeds both editors passes the same identity, permissions and versions adapters to each. The `./host` entry holds the types, an in-memory host and pure causal transaction receipt bookkeeping; it needs no DOM.

The editor is switched to these adapters one concern at a time. It takes the reader's identity and capabilities (SE-ADAPTERS-2.3), presence (SE-ADAPTERS-4.3), a workbook session for saving (SE-SESSION) and the version store (SE-ADAPTERS-5.1).

## SE-ADAPTERS-1 - Contract

### SE-ADAPTERS-1.1 - One host object

The host shall be one `SpreadsheetHost` object. `identity` and `permissions` shall be required; `workbooks`, `versions`, `presence`, `drafts` and `lifecycle` shall be optional, and a feature whose adapter is absent stays off rather than failing.

Implementation: `src/host/adapters.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-1.2 - Results instead of exceptions

Adapter requests that can fail for an expected reason shall resolve to a `HostResult`: `{ ok: true, value }` or `{ ok: false, error }`, where `error.kind` is `not-found`, `denied`, `conflict` or `unavailable` and `error.message` is optional text for the reader.

Implementation: `src/host/adapters.ts`, `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-1.3 - Revisions guard saves

A stored workbook shall come with its saved state: `revision`, the host's own identifier, which the editor only compares for equality, and `sequence`, an integer that orders saved states. A save shall name the revision its content was edited from (`baseRevision`). The workbook store shall refuse a save whose base revision is not the stored revision with a `conflict` error, which may carry the workbook as it is stored now (`error.current`), and return the new saved state when it accepts one. It may tell a subscriber each new saved state, by anyone, until it unsubscribes (`subscribe`).

Implementation: `src/host/adapters.ts`, `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

The pure `createTransactionReceiptTracker` in `./host` records the existing editor transaction IDs with local coverage revisions and releases acknowledged IDs. It creates no transaction identity, provider or saving state. [HOST-INTEGRATION.md](HOST-INTEGRATION.md) specifies causal coverage and the pending workbook-session wiring.

### SE-ADAPTERS-1.4 - Identity

The `identity` adapter shall give the reader (`currentPerson`), find people whose names match a query (`findPeople`) and resolve people by id, leaving out unknown ids (`resolvePeople`). A person is an id, a name and optionally an avatar URL.

Implementation: `src/host/adapters.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-1.5 - Versions, drafts and lifecycle

The `versions` adapter shall list a workbook's saved versions, newest first, each with its saved state, when it was saved, the id of who saved it and its name, and load a version's content by its revision; it may save content as a named version (`create`), name a version (`rename`), make a version the current content as a new saved state (`restore`) and tell a subscriber when the versions change. The `drafts` adapter shall keep at most one draft per workbook: its content, the revision it was edited from and when it was kept. The `lifecycle` adapter shall tell the editor when the page is `hidden`, `visible` again or being left (`leaving`).

Implementation: `src/host/adapters.ts`
Verification: `src/host/memory-host.test.ts`

## SE-ADAPTERS-2 - Capabilities

### SE-ADAPTERS-2.1 - Capabilities, not roles

The host shall describe what the reader may do with a workbook as `EditorCapabilities`, the document editor's: `edit`, `suggest`, `comment`, `decide` on others' suggestions, `saveVersion` and open the `history`. `resolveSpreadsheetMode(capabilities, { historical })` shall give the most the reader may do: `read` for a historical version, otherwise `author` when they may edit, `review` when they may suggest or comment, and `read` otherwise.

Implementation: `src/host/capabilities.ts`
Verification: `src/host/capabilities.test.ts`

### SE-ADAPTERS-2.2 - Following capabilities

The `permissions` adapter shall give the reader's capabilities on a workbook and may tell a subscriber whenever they change, until it unsubscribes.

Implementation: `src/host/adapters.ts`, `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-2.3 - The editor's identity and capabilities

The editor shall take the host's `identity` adapter, whose reader (`currentPerson`, by id and name) is the actor of the comments, suggestions and decisions they make, and `capabilities`, all of them by default. It shall offer at most the mode `resolveSpreadsheetMode` gives: Authoring and every mode below it for a reader who may edit; Reviewing, and no other mode, for one who may only suggest or comment, and Reading when the features they may use there are turned off; Reading alone otherwise, with no mode to choose. Within a mode, editing needs `edit`, suggesting `suggest`, commenting `comment`, and deciding suggestions `decide` in Authoring. A historical version is read-only whatever the capabilities.

Implementation: `src/editor/use-editor.ts`, `src/editor/spreadsheet-editor.tsx`, `src/editor/editor-toolbar.tsx`, `src/editor/editor-menus.ts`
Verification: `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`

## SE-ADAPTERS-3 - In-memory host

### SE-ADAPTERS-3.1 - Complete standalone host

`createMemoryHost(options)` shall return a host with in-memory identity, permissions, workbooks, versions, presence, drafts and lifecycle. It shall default to a reader named Reader, with id `reader`, with every capability and no workbooks.

Implementation: `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-3.2 - Isolated values

The in-memory stores shall keep and return copies, so a caller that mutates a workbook it passed in or loaded does not change what is stored, and two hosts shall share nothing but a presence room they are given.

Implementation: `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-3.3 - People lookup

The in-memory identity adapter shall know the reader and `options.people`, find people whose names contain the query, ignoring case and surrounding spaces, and resolve people by id in the order asked, leaving out unknown ids.

Implementation: `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-3.4 - In-memory versions

The in-memory host shall keep every saved state of a workbook as a version, from the state it started with, attributed to the reader at `options.now`, with revisions `r1`, `r2` and so on.

Implementation: `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-3.5 - In-memory drafts

The in-memory host shall keep one draft per workbook, as a copy.

Implementation: `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-3.6 - Driving an in-memory host

The in-memory host's `memory` controls shall fail the next requests of a kind (`load`, `save`, `versions`, `drafts`) with a given error (`failNext`), give the reader other capabilities and tell the subscribers (`setCapabilities`), and tell the lifecycle subscribers of a lifecycle change (`emitLifecycle`), for labs and tests.

Implementation: `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

## SE-ADAPTERS-4 - Presence

### SE-ADAPTERS-4.1 - Presence adapter

Workbooks have no shared room, so the host carries presence. The `presence` adapter shall share the reader's presence in a workbook, or withdraw it with null (`publish`), and tell a subscriber the other people in the workbook, without the reader, whenever they change (`subscribe`). The reader shares where they are (`at`: a worksheet id and an A1 range, or null), the range they are writing in (`editLock`), which others treat as an advisory lock, and whether they are `active`, `idle` or `away`. A peer adds their id, name and, optionally, picture (`avatar`) and colour (`color`), as the document editor's peers do.

Implementation: `src/host/presence-adapters.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-4.2 - In-memory presence

`createMemoryPresenceRoom()` shall keep presence in memory for everyone who joins it (`join(person)`), each seeing the others sorted by name, and the subscriber shall be told the current peers when it subscribes. A memory host joins the room `options.presence` names, or a room of its own.

Implementation: `src/host/memory-presence.ts`, `src/host/memory-host.ts`
Verification: `src/host/memory-host.test.ts`

### SE-ADAPTERS-4.3 - The editor's presence

Given a `presence` adapter and a `workbookId`, the editor shall subscribe to the other people in that workbook and publish the reader's place: `at` is the active worksheet and the selected range, and `editLock` is the same place while the reader is writing in a cell, otherwise null. It shall publish again when the place, the lock or the status changes, and withdraw (`publish(workbookId, null)`) when it unmounts or the adapter or workbook id changes. The reader is `away` while the page is hidden and `active` otherwise. Without either prop the editor has no peers and publishes nothing.

Peers are placed on the workbook before they are shown: the reader and peers without an id or a name are left out, and a place on a worksheet the workbook lacks or in a malformed range is dropped rather than trusted. The toolbar lists the peers with their status, following one moves to their selection, their selections and locks are drawn on the grid in a colour slot from their id, and a lock refuses the reader's writes to those cells (advisory; the host enforces its own policy). A new `workbookId` also starts a new editor session, as a new `sessionKey` does.

Implementation: `src/editor/use-workbook-presence.ts`, `src/editor/use-editor.ts`, `src/model/presence.ts`, `src/editor/editor-toolbar.tsx`, `src/engine/workbook-grid.tsx`
Verification: `src/editor/host-integration.test.tsx`, `src/model/presence.test.ts`, `src/editor/use-editor.test.ts`, `e2e/interaction-parity.spec.ts` (collaboration)

## SE-SESSION - Workbook session

A workbook session opens one stored workbook through the host's `workbooks` store and saves it against the revision it was edited from, as document-editor's store sessions do. It needs no framework or DOM: `createWorkbookSession(options)` lives in `./host`, and `useWorkbookSession(options)` in the main entry opens one for as long as a component is mounted. The editor takes it as its `session` prop in place of the `onSave`, `onFlush` and `autosaveMs` props it had before.

### SE-SESSION-1 - Opening

The session shall load the workbook (`workbooks.load`) and report `loading`, then `saved` at the stored revision, or `unavailable` with the host's error. While its session is opening, the editor renders a status in place of the grid (`[data-session="loading"]`), or the error (`[data-session="unavailable"]`); a null `session` prop also shows it as opening. When another save of the workbook lands (`workbooks.subscribe`) and the session holds no unsaved change, it loads and shows that save; otherwise it keeps the reader's changes.

Implementation: `src/host/workbook-session.ts`, `src/editor/use-session-state.ts`, `src/editor/spreadsheet-editor.tsx`
Verification: `src/host/workbook-session.test.ts`

### SE-SESSION-2 - Causal coverage

Each accepted local change shall be recorded with its `WorkbookChangeContext.transactionId` and a local revision. A save covers the changes accepted before it started, sends their content with the stored revision as `baseRevision`, and on its answer marks exactly those as saved: a change accepted while it was in flight stays unsaved and saves next. A save asked for while another is in flight joins it, and saves again only when the first did not cover what was accepted before it was asked for. `acknowledge(transactionId)` marks a transaction and every earlier one as saved for a host that keeps changes by other means, such as an operation feed; an unknown, repeated or already covered id changes nothing. Comment and other annotation changes are transactions like any other. A remote sequence number never counts as saving a local change.

Autosave runs `autosaveDelay` milliseconds after the last change (1000 by default; `Infinity` turns it off), and when the `lifecycle` adapter reports the page hidden or left. An autosave that finds the store `unavailable` retries with backoff from half a second to four seconds, for up to 30 seconds; a save the reader asks for, and every other failure, stops at `error` with the host's message and keeps the changes unsaved. A save refused as `conflict` stops at `conflict`: the editor shows a banner (`.se-editor__conflict`) to keep the reader's workbook over the stored one (`keepMine`, which saves against the stored revision) or take the stored one and drop the reader's changes (`takeStored`). Saving does not resume until one is chosen.

Implementation: `src/host/workbook-session.ts`, `src/host/transaction-receipts.ts`
Verification: `src/host/workbook-session.test.ts`, `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`

### SE-SESSION-3 - Drafts

Given a `drafts` adapter, the session shall keep unsaved content as a draft `draftDelay` milliseconds after the last change (300 by default), with the revision it was edited from, and clear it once a save covers every change. A session that opens a workbook with a kept draft offers it (`draft()`), marked `stale` when the stored revision moved on since; the editor shows a banner (`.se-editor__draft`) to restore it as unsaved changes (`restoreDraft`) or discard it (`discardDraft`).

Implementation: `src/host/workbook-session.ts`, `src/editor/spreadsheet-editor.tsx`
Verification: `src/host/workbook-session.test.ts`

### SE-SESSION-4 - The editor's saving

With a session, the editor shall show the session's workbook unless `value` is given, record every accepted change in the session after the host's `onChange` accepts it, and refuse a change while the session has not loaded. Save (button, menu and Ctrl+S) commits the active draft and saves through the session; the handle's `flush` commits the draft and waits until every change accepted before it is saved. The save status shows the session's phase, and the handle's state reports `dirty` and `saving` from it. Opening a saved version first flushes, and stays closed when the flush fails or a newer change is unsaved (SE-ADAPTERS-5.1). Without a session the editor has no Save control and saves nothing.

Implementation: `src/editor/use-editor.ts`, `src/editor/use-workbook-session.ts`, `src/editor/editor-toolbar.tsx`
Verification: `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`, `src/editor/editor-lifecycle.test.tsx`, `e2e/host-integration.spec.ts`, `e2e/full-editor.spec.ts`

## SE-ADAPTERS-5 - Versions

### SE-ADAPTERS-5.1 - The editor's versions

Given a `versions` adapter, a `workbookId` and the `history` capability, the history pane shall list the workbook's versions while it is open, newest first, each with who saved it (named through the identity adapter, the id otherwise), when, and its name or "Saved automatically". It lists them again when the store reports a change and on Refresh, shows only the latest answer, and shows a failed list as an error. Without the adapter or the capability the pane shows the session's editing history instead.

Viewing a version (its context menu) shall commit the active draft and flush the session first, and stay on the current workbook when that fails or a newer change is unsaved. The editor then loads the version's content and shows it in place of the workbook, read-only whatever the capabilities, with a `.se-editor__historical` banner naming the version and a button back to the current workbook. A failed load is shown as a notice. The handle's `getSnapshot` returns the workbook shown.

Where the store allows it: with `saveVersion` and `create`, the pane saves the workbook as a version under an optional name, after flushing the session; with `saveVersion` and `rename`, a version's context menu names it again, and an empty name takes its name away; with `edit` and `restore`, a version's context menu and the banner make it the current content, after which the editor takes the stored workbook from the session and returns to it.

Implementation: `src/editor/use-workbook-versions.ts`, `src/editor/use-editor.ts`, `src/editor/review-panels.tsx`, `src/editor/spreadsheet-editor.tsx`
Verification: `src/editor/editor-versions.test.ts`, `src/editor/editor-lifecycle.test.tsx`, `src/editor/use-editor.test.ts`, `e2e/full-editor.spec.ts` (SE-CHROME-1), `e2e/interaction-parity.spec.ts` (SE-VR-1)
