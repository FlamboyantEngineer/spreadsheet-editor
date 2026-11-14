# Host integration

## SE-HOST-1 — Semantic accepted changes

`onChange(next, context)` emits one local transaction per accepted command. The context includes a detached baseline, transaction ID, source and ordered `WorkbookOperation` patches. Line commands carry explicit stable identities and neighbors even for blank rows. Remote controlled values never emit a local change. Synchronous host rejection retains accepted data, history and drafts and fails navigation/save/flush. Hosts preflight before mutating shared state; asynchronous durability failures belong to persistence callbacks. External undo owns its lifetime and every undo surface delegates when configured.

Implementation: `src/model/operations.ts`, `src/editor/command-operations.ts`, `src/editor/use-editor.ts`.

Verification: `src/model/operations.test.ts`, `src/editor/host-integration.test.tsx`.

## SE-HOST-2 — Session lifecycle

`editorRef` exposes a stable session handle for detached snapshots, draft commit/cancel, selection, state subscription, Save and Flush. Commit synchronously publishes native/formula-bar input once. Flush drains scheduled saves and calls the optional host durability barrier without creating a manual history snapshot. Persistence rejection leaves local changes dirty. Local revision tracking excludes remote object replacement; an earlier save cannot cover a later local edit. Unmount/session replacement aborts requests and invalidates handles and subscriptions. Providers remain host-owned. `restoreDraft` restores retained uncommitted input without changing accepted data, and `setPanel` opens or collapses a host-selected panel. `panelRef`, `panelId`, `panelData` and `onPanelChange` allow the host to connect its existing dock lifecycle without discovering or modifying package markup.

Implementation: `src/editor/contracts.ts`, `src/editor/public-host.ts`, `src/editor/use-editor.ts`.

Verification: `src/editor/host-integration.test.tsx`, `src/editor/use-editor.test.ts`.

## SE-HOST-3 — Authoritative publications

External mode accepts bounded version-one results tied to `workbookCalculationKey` and a monotonic session revision. Reusing a revision for different content is rejected. It never invokes local calculation. `reviewCalculation` supplies an independently keyed publication for the current actor’s suggestion projection; missing or stale review publications remain pending rather than calculating locally. Per-cell pending/unavailable states cannot retain cached values as current. Exports and ordinary document insertion require current ready values for all requested authored cells. Native calculation records actual dependencies; protected drafts invalidate their cell and transitive dependents without evaluation.

Implementation: `src/calculation/publication.ts`, `src/editor/use-publication.ts`, `src/editor/use-editor.ts`.

Verification: `src/calculation/publication.test.ts`, `src/editor/host-integration.test.tsx`.

## SE-HOST-4 — Catalog and safe presentation

Host descriptors extend both formula entry surfaces without overriding built-ins or authorizing execution. `includeBuiltinFormulaFunctions={false}` makes the supplied descriptors the complete catalog in both formula entry surfaces; a host can then advertise only its authorized evaluator capabilities, including names also present in the native catalog. Host `formatting.kinds` and `formatting.decimalPlaces` constrain both visible formatting controls and accepted formatting commands to the host’s storage capabilities. `features.newWorkbook` and `features.import` can hide replacement actions when the host owns artifact creation and import. Comment-only permissions retain Reviewing mode for comment actions without permitting cell changes, suggestions or decisions. Matching input/revision presentation may supply bounded text, accessible descriptions, default units and semantic decorations. Explicit formats take precedence. Current ready cells may mount React content through engine-owned portals; pending drafts and stale publications clear external content. Search includes authored expressions and current formatted results; replacement only changes authored matches.

Implementation: `src/editor/public-host.ts`, `src/editor/editor-formula-bar.tsx`, `src/engine/native-formula.tsx`, `src/engine/workbook-grid.tsx`.

Verification: `src/editor/host-integration.test.tsx`, `e2e/core-engine.spec.ts`.

See [the implementation plan](../docs/CORE-EDITOR-PLAN.md) for compatibility and acceptance requirements. Host application storage, authorization, transport, immutable history and rich-content sanitation remain outside the package.
