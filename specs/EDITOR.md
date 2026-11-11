# Standalone editor

## SE-UI-1 — React editor and shared controls

The editor renders the full reusable editor chrome using the base component library and a canonical React Jspreadsheet host. It supports controlled or local workbook state, the menu bar, formatting tools, raw-value formula bar, sheet actions, search/replace, right panels, save status, accessible selection statistics, worksheet-strip zoom and a bounded 50-step undo/redo history. Read-only mode locks content and review mutation in every input path. A host-controlled workbook replacement resets available history. [EDITOR-CHROME.md](EDITOR-CHROME.md) and [REVIEW.md](REVIEW.md) specify the controls, comments and tracked suggestions.

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

`npm run labs` installs any missing or stale locked dependencies, then builds and serves labs plus a real server calculation endpoint. It requires no application server, database or consumer repository checkout. Host and port are configurable. The harness serves only browser assets and bounds request bodies. CI must exercise actual browser edits, HTTP calculation and both insertion modes.

Implementation: `scripts/labs.mjs`, `scripts/build-labs.mjs`, `labs/src/main.tsx`.

Verification: `e2e/editor.spec.ts`, launched by the Playwright web-server configuration through `npm run labs`.

## SE-LABS-2 — Installation diagnostics and portable startup

A fresh source checkout shall start with `npm run labs`, without a preceding install command. Node 22.13+ within Node 22 or Node 24+, npm and Git remain system prerequisites. `.nvmrc` selects Node 24 locally and in CI, and `.npmrc` rejects unsupported runtimes during installation. The launcher and standalone labs build shall inspect mandatory installed packages and npm's installed lockfile against the committed lockfile, verify prepared UI assets and exercise esbuild's native binary in a separate process. Missing, stale, incomplete or wrong-platform installs shall trigger one `npm ci --include=dev --include=optional --ignore-scripts=false` before building; production/omit/ignore-scripts defaults must not prevent setup. Healthy installs shall be reused without network installation. The installer shall run the current Node executable with npm's JavaScript CLI path, including on Windows and in paths containing spaces, without a shell or npm.cmd.

Installation errors retain npm's output and exit code and stop startup without retry loops. A successful install must pass readiness checks before build or serving. Actual source/build failures remain visible and do not trigger dependency reinstall. The public UI library requires no private token. Git package preparation continues to launch the current Node executable directly.

Implementation: `scripts/lib/labs-dependencies.mjs`, `scripts/lib/labs-setup.mjs`, `scripts/labs.mjs`, `scripts/build-labs.mjs`, `scripts/prepare.mjs`.

Verification: `src/tooling/labs-startup.test.ts`, existing browser labs tests.

## Extraction boundary

This stage extracts reusable grid and parser modules and provides an independent rational calculation engine, complete editor chrome, comments, tracked suggestions and host-neutral integration APIs. Existing consumers are migrated separately. Storage, authorization, collaboration transports, durable history and application lifecycle services stay outside this package. Hosts may supply extra workflow/assistant panes. Curated rational unit/quantity evaluation and optional Node XLSX conversion are included; application-specific external formulas remain host-provided. [HOST-PARITY.md](HOST-PARITY.md) records the reviewed requirements and host boundaries. Row/column insertion/deletion rewrites references and moves review anchors within the package.

## SE-LABS-3 — Independent configurations

The labs index shall link to minimal, embedded and full editor pages. Each page shares ui-components-style Stylesheet (Reference/Unstyled) and Theme (System/Light/Dark) selectors; selections follow navigation. Minimal provides the editable grid and worksheet strip, embedded provides compact tools inside an application-owned workspace, and full provides all menus and review commands. Reference presentation must keep formula entry, address, Apply/Cancel and formatting controls usable at desktop/tablet/mobile widths without horizontal page overflow. Generic field labels remain accessible while compact chrome omits visible labels and empty feedback rows.

Verification: `e2e/labs.spec.ts` at 1440, 768 and 390 pixels, plus `e2e/editor.spec.ts`.
