# Standalone editor

## SE-UI-1 — React editor and shared controls

The editor renders the full reusable editor chrome using the base component library and a canonical React Jspreadsheet host. It supports controlled or local workbook state, the menu bar, formatting tools, raw-value formula bar, sheet actions, search/replace, right panels, save status, selection statistics, zoom and a bounded 50-step undo/redo history. Read-only mode locks content and review mutation in every input path. A host-controlled workbook replacement resets available history. [EDITOR-CHROME.md](EDITOR-CHROME.md) and [REVIEW.md](REVIEW.md) specify the controls, comments and tracked suggestions.

Implementation: `src/editor/spreadsheet-editor.tsx`, `src/editor/editor-toolbar.tsx`, `src/editor/use-editor.ts`, `src/engine/workbook-grid.tsx`, `src/engine/jspreadsheet-host.tsx`.

Verification: `e2e/editor.spec.ts` (SE-UI-1, SE-UI-3, SE-UI-7), `e2e/full-editor.spec.ts`, `src/model/workbook.test.ts`.

## SE-UI-2 — Interchangeable calculation location

Each editor instance selects client calculation or an explicit asynchronous server provider. Pending results and failures are visible. Formula cells show a placeholder while pending and Unavailable after a provider failure. Exporting a normal document table requires current results. Failure must not fall back to local calculation. Changing workbook, provider or options invalidates old results immediately; aborted or stale responses cannot replace a newer result. Unmount cancels pending work.

Implementation: `src/editor/use-calculation.ts`, `src/calculation/provider.ts`, `src/calculation/http.ts`, `src/engine/cell-renderer.ts`.

Verification: `src/editor/use-calculation.test.ts`, `src/calculation/http.test.ts`, `e2e/editor.spec.ts` (SE-UI-2, SE-UI-5).

## SE-UI-3 — Instance isolation and authored data

The grid receives a copy of authored data and preserves formula text and decimal/scientific literals. Rendering calculated text never changes raw data. Multiple instances may coexist; live document insertions and undo snapshots do not mutate their source workbooks.

Implementation: `src/engine/workbook-grid.tsx`, `src/engine/cell-renderer.ts`, `src/document/insertion.ts`.

Verification: `e2e/editor.spec.ts` (SE-UI-3, SE-DOC-2), `src/document/insertion.test.ts`, `src/model/workbook.test.ts`.

## SE-LABS-1 — One-command labs

After installing dependencies, `npm run labs` builds and serves labs plus a real server calculation endpoint. It requires no application server, database or consumer repository checkout. Host and port are configurable. The harness serves only browser assets and bounds request bodies. CI must exercise actual browser edits, HTTP calculation and both insertion modes.

Implementation: `scripts/labs.mjs`, `scripts/build-labs.mjs`, `labs/src/main.tsx`.

Verification: `e2e/editor.spec.ts`, launched by the Playwright web-server configuration through `npm run labs`.

## SE-LABS-2 — Installation diagnostics and portable startup

Labs require Node 22 or newer and a completed development dependency installation (`npm ci --include=dev`). The launcher and standalone labs build must report a missing esbuild package with the exact installation command and private-library access guidance, without an uncaught module-resolution traceback. Actual build errors and missing transitive dependencies must retain their original diagnostics. Git package preparation launches the current Node executable directly, including from Windows paths with spaces, without relying on a shell or npm.cmd.

Implementation: `scripts/lib/labs-setup.mjs`, `scripts/labs.mjs`, `scripts/build-labs.mjs`, `scripts/prepare.mjs`.

Verification: `src/tooling/labs-startup.test.ts`, existing browser labs tests.

## Extraction boundary

This stage extracts reusable grid and parser modules and provides an independent rational calculation engine, complete editor chrome, comments, tracked suggestions and host-neutral integration APIs. Existing consumers are migrated separately. Storage, authorization, collaboration transports, durable history and application lifecycle services stay outside this package. Hosts may supply extra workflow/assistant panes. Built-in unit/quantity evaluation, application-specific external formulas and XLSX import/export are outside this contract. Row/column insertion/deletion rewrites references and moves review anchors within the package.
