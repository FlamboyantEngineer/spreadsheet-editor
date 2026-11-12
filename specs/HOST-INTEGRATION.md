# Host integration

## SE-HOST-1 — Semantic accepted changes

`onChange(next, context)` emits one local transaction per accepted command. The context includes a detached baseline, transaction ID, source and ordered `WorkbookOperation` patches. Line commands carry explicit stable identities and neighbors even for blank rows. Ordinary writes beyond existing identities emit the appended line insertions before cell patches, using the host's ID factory. Workbook-title changes emit invertible `workbook-metadata` operations. Remote controlled values never emit a local change. Synchronous host rejection retains accepted data, history and drafts and fails navigation/save/flush. Hosts preflight before mutating shared state; asynchronous durability failures belong to persistence callbacks. External undo owns its lifetime and every undo surface, including the history pane, delegates when configured.

Implementation: `src/model/operations.ts`, `src/editor/command-operations.ts`, `src/editor/use-editor.ts`.

Verification: `src/model/host-model.test.ts`, `src/editor/host-integration.test.tsx`, `src/editor/review-panels.test.tsx`.

## SE-HOST-2 — Session lifecycle

`editorRef` exposes a stable session handle for detached snapshots, draft commit/cancel, selection, state subscription, Save and Flush. Commit synchronously publishes native/formula-bar input once. Flush drains scheduled saves and calls the optional host durability barrier without creating a manual history snapshot. Persistence rejection leaves local changes dirty. Local revision tracking excludes remote object replacement; an earlier save cannot cover a later local edit. Unmount/session replacement aborts requests and invalidates handles and subscriptions. Providers remain host-owned.

A flush requested while another is active first commits the current draft. It may share the pending receipt only when that receipt covers its requested local revision; otherwise it queues a new barrier. Permission loss aborts active barriers. Manual Save cancels a redundant scheduled autosave. Opening a historical version commits native and formula-bar drafts, uses configured save and flush adapters, and prevents navigation after failed persistence or changes newer than its receipt. A pristine history navigation performs no save or receipt request.

Implementation: `src/editor/contracts.ts`, `src/editor/public-host.ts`, `src/editor/use-editor.ts`.

Verification: `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`, `src/editor/editor-lifecycle.test.tsx`.

## SE-HOST-3 — Authoritative publications

External mode accepts bounded version-one results tied to `workbookCalculationKey` and a monotonic session revision. Reusing a revision for different content is rejected. It never invokes local calculation. Per-cell pending/unavailable states cannot retain cached values as current. Exports and ordinary document insertion require current ready values for all requested authored cells. Native calculation records actual dependencies; protected drafts invalidate their cell and transitive dependents without evaluation.

Client/server requests use authored calculation identity rather than workbook object identity. Renames, visual styles and equivalent controlled snapshots retain pending/ready results without additional calculation. Changed calculation inputs, provider/options, disabling or re-enabling calculation invalidate a pass and reject late replies.

Implementation: `src/calculation/publication.ts`, `src/editor/use-publication.ts`, `src/editor/use-editor.ts`.

Verification: `src/calculation/publication.test.ts`, `src/editor/host-integration.test.tsx`, `src/editor/use-calculation.test.ts`.

## SE-HOST-4 — Catalog and safe presentation

Host descriptors extend both formula entry surfaces without overriding built-ins or authorizing execution. Matching input/revision presentation may supply bounded text, accessible descriptions, default units and semantic decorations. Explicit formats take precedence. Current ready cells may mount React content through engine-owned portals; pending drafts and stale publications clear external content. Search includes authored expressions and current formatted results; replacement only changes authored matches.

Previous starts at the last result when none is active. Query changes reset the active result; failed draft commitment cannot advance it. Result-only matches disable replacement controls. Copy uses ready calculated text, falling back to authored inputs for pending/unavailable cells rather than producing blank or stale clipboard data. Selection statistics use quantity-aware arithmetic and retain units; incompatible dimensions, kinds or provenance suppress aggregates independently while preserving nonempty counts.

Implementation: `src/editor/public-host.ts`, `src/editor/editor-formula-bar.tsx`, `src/engine/native-formula.tsx`, `src/engine/workbook-grid.tsx`.

Verification: `src/editor/host-integration.test.tsx`, `src/editor/editor-search.test.tsx`, `e2e/core-engine.spec.ts`, `e2e/host-integration.spec.ts`.

See [the implementation plan](../docs/CORE-EDITOR-PLAN.md) for compatibility and acceptance requirements. Host application storage, authorization, transport, immutable history and rich-content sanitation remain outside the package.
