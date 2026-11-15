# Host integration

## SE-HOST-1 - Semantic accepted changes

### SE-HOST-1.1 - Semantic accepted changes

`onChange(next, context)` emits one local transaction per accepted command. The context includes a detached baseline, transaction ID, source and ordered `WorkbookOperation` patches. Line commands carry explicit stable identities and neighbors even for blank rows. Ordinary writes beyond existing identities emit the appended line insertions before cell patches, using the host's ID factory. Workbook-title changes emit invertible `workbook-metadata` operations. Remote controlled values never emit a local change. Synchronous host rejection retains accepted data, history and drafts and fails navigation/save/flush. Hosts preflight before mutating shared state; asynchronous durability failures belong to the workbook session. External undo owns its lifetime and every undo surface, including the history pane, delegates when configured.

Implementation: `src/model/operations.ts`, `src/editor/command-operations.ts`, `src/editor/use-editor.ts`.

Verification: `src/model/host-model.test.ts`, `src/editor/host-integration.test.tsx`, `src/editor/review-panels.test.tsx`.

## SE-HOST-2 - Session lifecycle

### SE-HOST-2.1 - Session lifecycle

`editorRef` exposes a stable session handle for detached snapshots, draft commit/cancel, selection, state subscription, Save and Flush. Commit synchronously publishes native/formula-bar input once. Save and Flush go through the editor's workbook session ([HOST-ADAPTERS.md](HOST-ADAPTERS.md), SE-SESSION): Flush commits the active draft and waits until every change accepted before it is saved, without waiting for later ones. A failed save leaves local changes dirty and retains the error through newer edits until explicit retry or success. The handle reports session failures even when a background autosave produced them. Remote object replacement is not a local change; an earlier save cannot cover a later local edit. Unmount/session replacement invalidates handles and subscriptions. Providers remain host-owned. `restoreDraft` restores retained uncommitted input without changing accepted data, and `setPanel` opens or collapses a host-selected panel. `panelRef`, `panelId`, `panelData` and `onPanelChange` allow the host to connect its existing dock lifecycle without discovering or modifying package markup. `onPanelChange` is called once on mount and then only when the panel id or its expanded state changes, never because the host passed a new callback, so a host may store the state from an inline callback.

A flush requested while another is active first commits the current draft, and waits for a further save when the active one does not cover its draft. Opening a saved version (SE-ADAPTERS-5.1) commits native and formula-bar drafts, flushes the session, and stays on the current workbook after a failed save or when a change newer than that save is unsaved. Opening a version from a pristine session saves nothing.

Implementation: `src/editor/contracts.ts`, `src/editor/public-host.ts`, `src/editor/use-editor.ts`, `src/editor/editor-handle.ts`, `src/editor/use-session-state.ts`, `src/editor/use-session-actions.ts`.

Verification: `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`, `src/editor/editor-lifecycle.test.tsx`.

## SE-HOST-3 - Authoritative publications

### SE-HOST-3.1 - Authoritative publications

External mode accepts bounded version-one results tied to `workbookCalculationKey` and a monotonic session revision. The key is a short digest (`wck1-` and 28 hexadecimal digits) of the canonical calculation inputs, so equal inputs share a key across transports whatever the workbook's size; it is a freshness token, not an authorization signature. The editor computes it once per accepted workbook. Reusing a revision for different content is rejected. It never invokes local calculation. `reviewCalculation` supplies an independently keyed publication for the current actor’s suggestion projection; missing or stale review publications remain pending rather than calculating locally. Per-cell pending/unavailable states cannot retain cached values as current. Exports and ordinary document insertion require current ready values for all requested authored cells. Native calculation records actual dependencies; protected drafts invalidate their cell and transitive dependents without evaluation.

Client/server requests use authored calculation identity rather than workbook object identity. Renames, visual styles and equivalent controlled snapshots retain pending/ready results without additional calculation. Changed calculation inputs, provider/options, disabling or re-enabling calculation invalidate a pass and reject late replies.

Implementation: `src/calculation/publication.ts`, `src/editor/use-publication.ts`, `src/editor/use-editor.ts`.

Verification: `src/calculation/publication.test.ts`, `src/editor/host-integration.test.tsx`, `src/editor/use-calculation.test.ts`.

## SE-HOST-4 - Catalog and safe presentation

### SE-HOST-4.1 - Catalog and safe presentation

Host descriptors extend both formula entry surfaces without overriding built-ins or authorizing execution. `includeBuiltinFormulaFunctions={false}` makes the supplied descriptors the complete catalog in both formula entry surfaces; a host can then advertise only its authorized evaluator capabilities, including names also present in the native catalog. Host `formatting.kinds` and `formatting.decimalPlaces` constrain both visible formatting controls and accepted formatting commands to the host’s storage capabilities. `features.newWorkbook` and `features.import` can hide replacement actions when the host owns workbook creation and import. Comment-only permissions retain Reviewing mode for comment actions without permitting cell changes, suggestions or decisions. Matching input/revision presentation may supply bounded text, accessible descriptions, default units and semantic decorations. Explicit formats take precedence. Current ready cells may mount React content through engine-owned portals; pending drafts and stale publications clear external content. Search includes authored expressions and current formatted results; replacement only changes authored matches.

Previous starts at the last result when none is active. Query changes reset the active result; failed draft commitment cannot advance it. Result-only matches disable replacement controls. Copy uses ready calculated text, falling back to authored inputs for pending/unavailable cells rather than producing blank or stale clipboard data. Selection statistics use quantity-aware arithmetic and retain units; incompatible dimensions, kinds or provenance suppress aggregates independently while preserving nonempty counts.

Implementation: `src/editor/public-host.ts`, `src/editor/editor-formula-bar.tsx`, `src/engine/native-formula.tsx`, `src/engine/workbook-grid.tsx`.

Verification: `src/editor/host-integration.test.tsx`, `src/editor/editor-search.test.tsx`, `e2e/core-engine.spec.ts`, `e2e/host-integration.spec.ts`.

See [the implementation plan](../docs/CORE-EDITOR-PLAN.md) for compatibility and acceptance requirements. Host application storage, authorization, transport, immutable history and rich-content sanitation remain outside the package.

## SE-HOST-5 - Atomic operation application

### SE-HOST-5.1 - Atomic operation application

The headless `/model` exports `applyWorkbookTransaction(current, context, { appliedTransactionIds? })`, `applyWorkbookOperations` and `createWorkbookOperationBinding`. The binding reads the latest accepted host snapshot, applies the full semantic transaction in memory, validates it, and calls the host's synchronous atomic `commit` exactly once. It retains up to `rememberTransactions` accepted transaction IDs (default 10,000) across transport/provider replacement and rejects duplicates in that window. The host’s synchronous `hasTransaction(id)` lookup extends deduplication to older receipts or session restart using durable history. No full transaction contexts are retained. No global document, provider, undo manager or calculator is created. Hosts authorize actor ownership and permissions before their atomic commit; these functions are not an authorization boundary.

Cells resolve stable row/column identities. Threads, posts and suggestions merge by entity ID, individual object fields merge independently, and reaction actors merge by membership. Canonical numeric values remain atomic exact rational JSON. Changed scalar values and conflicting suggestion decisions reject rather than overwrite accepted data. An inverse from `invertWorkbookOperations` undoes an actor's own fields while retaining independent edits. Deleted cell targets reject. Transactional line deletion also rejects concurrent modifications to cells or geometry in the targeted line, preserving accepted changes for explicit conflict resolution. Structure commands address existing line identities and insert between surviving neighbors in the host's serialized acceptance order; this layer does not implement a CRDT or promise transport-independent ordering.

Every structure operation (inserting or deleting rows and columns with their heights, widths, merges, formats, styles, comments and frozen columns; adding, removing, renaming and reordering worksheets), applied and then inverted, restores the workbook exactly; maps and lists a workbook did not have stay absent. A rename's inverse restores only the formulas the rename rewrote. A formula written after the rename that names the new worksheet is not part of those operations: `applyWorkbookTransaction` rejects such a stale inverse (`referential-rename-baseline`), while plain `applyWorkbookOperations` leaves that formula naming the missing worksheet. Rewriting references the operations do not name is not part of this requirement.

Use `initializeWorkbookIdentities` and persist that shared baseline once before opening concurrent structural editing sessions. Existing identities are preserved. Missing identities reject structural application. `applyWorkbookTransaction` compares the baseline's line identities to the latest snapshot. After concurrent structure, ambiguous positional geometry/merges, formula writes, changed suggestion records, or new/moved comment anchors reject atomically; ordinary identified cells and replies/reactions still rebase. Hosts may retry an explicitly refreshed command. Worksheet reordering rejects a competing reorder while preserving independently added sheets. Worksheet renames reject a stale baseline rather than leaving concurrently introduced references unresolved; stale formula/suggestion inputs also reject after concurrent worksheet renaming. Inverse line commands transform their layout baseline before restoring deleted geometry/merges. `applyWorkbookOperations` is the lower-level application function for coordinates already serialized by the caller; use the transaction API for stale host transactions.

The pure `/host` `createTransactionReceiptTracker` is causal bookkeeping for the upcoming host workbook session: record the existing editor transaction IDs with their monotonic local revisions (coverage order, not a second identity generator), acknowledge only a known ID's coverage, and release covered IDs after successful Save/Flush. It owns no provider, save scheduler or UI phase; it retains only pending ID/revision pairs. A later cumulative acknowledgement releases earlier IDs; old/unknown receipts cannot acknowledge later edits. Callers keep transaction IDs unique within the session.

The current callback-based editor still marks controlled local changes dirty until its configured Save/Flush succeeds. External receipt projection is deliberately deferred to the approved `SpreadsheetHost` workbook-session migration rather than adding a transient parallel prop. That session must capture the local transaction/revision coverage when submitting a save, associate the returned host `SavedState.revision` with that submitted coverage, and acknowledge only those local edits. A subscription's remote revision or sequence alone cannot prove local coverage. Edits and comment/review-only transactions accepted after submission and uncommitted drafts must remain dirty. The same existing session state must drive subscriptions, toolbar and public handle saving/retrying/error projection. A receipt must not schedule saves, recompute calculations, or replace the explicit Flush durability barrier.

Implementation: `src/model/apply-operations.ts`, `src/host/transaction-receipts.ts`.

Verification: `src/model/apply-operations.test.ts`, `src/host/transaction-receipts.test.ts`, and the packed public consumers that the package check imports.
