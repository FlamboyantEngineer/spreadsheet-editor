# Standalone editor

## SE-UI-1 - React editor and shared controls

### SE-UI-1.1 - React editor and shared controls

The editor renders the full reusable editor chrome using the base component library and a canonical React Jspreadsheet host. It supports controlled or local workbook state, the menu bar, formatting tools, raw-value formula bar, sheet actions, search/replace, right panels, save status, accessible selection statistics, sheet-bar zoom and a bounded 50-step undo/redo history. Read-only mode locks content and review mutation in every input path. A host-controlled workbook replacement resets available history. [EDITOR-CHROME.md](EDITOR-CHROME.md) and [REVIEW.md](REVIEW.md) specify the controls, comments and tracked suggestions.

Implementation: `src/editor/spreadsheet-editor.tsx`, `src/editor/editor-toolbar.tsx`, `src/editor/use-editor.ts`, `src/engine/workbook-grid.tsx`, `src/engine/jspreadsheet-host.tsx`.

Verification: `e2e/editor.spec.ts` (SE-UI-1.1, SE-UI-3.1), `e2e/full-editor.spec.ts`, `src/model/workbook.test.ts`.

## SE-UI-2 - Interchangeable calculation location

### SE-UI-2.1 - Interchangeable calculation location

Each editor instance selects client calculation or an explicit asynchronous server provider. Pending results and failures are visible. Formula cells show a placeholder while pending and Unavailable after a provider failure. Exporting a normal document table requires current results. Failure must not fall back to local calculation. Changing workbook, provider or options invalidates old results immediately; aborted or stale responses cannot replace a newer result. Unmount cancels pending work.

Implementation: `src/editor/use-calculation.ts`, `src/calculation/provider.ts`, `src/calculation/http.ts`, `src/engine/cell-renderer.ts`.

Verification: `src/editor/use-calculation.test.ts`, `src/calculation/http.test.ts`, `e2e/editor.spec.ts` (SE-UI-2.1).

## SE-UI-3 - Instance isolation and authored data

### SE-UI-3.1 - Instance isolation and authored data

The grid receives a copy of authored data and preserves formula text and decimal/scientific literals. Rendering calculated text never changes raw data. Multiple instances may coexist; live document insertions and undo snapshots do not mutate their source workbooks.

Native multi-cell commands publish one accepted workbook transaction and one undo step, including Delete, Backspace and fill operations. Capturing those commands preserves sparse authored data and untouched canonical inputs. A host refusal restores the vendor display to accepted state. Active single-cell drafts retain preflight before the native editor closes. Finishing an unchanged native edit preserves the authored primitive type and canonical input; opening a logical cell and pressing Enter cannot turn it into text.

The initial asynchronous grid mount does not move focus away from editor or host controls. Later worksheet mounts may restore grid focus after a deliberate switch. Mount focus intent is captured for that grid instance so later parent renders cannot turn initial adoption into an autofocus request.

Implementation: `src/engine/workbook-grid.tsx`, `src/engine/cell-renderer.ts`, `src/document/insertion.ts`.

Verification: `e2e/editor.spec.ts` (SE-UI-3.1, SE-DOC-2.1), `e2e/core-engine.spec.ts`, `src/engine/worksheet-data.test.ts`, `src/document/insertion.test.ts`, `src/model/workbook.test.ts`.

## Extraction boundary

This stage extracts reusable grid and parser modules and provides an independent rational calculation engine, complete editor chrome, comments, tracked suggestions and host-neutral integration APIs. Existing consumers are migrated separately. Storage, authorization, collaboration transports, durable history and application lifecycle services stay outside this package. Hosts may supply extra workflow/assistant panes. Curated rational unit/quantity evaluation and optional Node XLSX conversion are included; application-specific external formulas remain host-provided. [HOST-COMPATIBILITY.md](HOST-COMPATIBILITY.md) records the reviewed requirements and host boundaries. Row/column insertion/deletion rewrites references and moves review anchors within the package.
