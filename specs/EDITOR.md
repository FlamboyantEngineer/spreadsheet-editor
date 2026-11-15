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

Native multi-cell commands publish one accepted workbook transaction and one undo step, including Delete, Backspace and fill operations. Capturing those commands preserves sparse authored data and untouched canonical inputs. A host refusal restores the vendor display to accepted state. Active single-cell drafts retain preflight before the native editor closes. Finishing an unchanged native edit preserves the authored primitive type and canonical input; opening a logical cell and pressing Enter cannot turn it into text.

The initial asynchronous grid mount does not move focus away from editor or host controls. Later worksheet mounts may restore grid focus after a deliberate switch. Mount focus intent is captured for that grid instance so later parent renders cannot turn initial adoption into an autofocus request.

Implementation: `src/engine/workbook-grid.tsx`, `src/engine/cell-renderer.ts`, `src/document/insertion.ts`.

Verification: `e2e/editor.spec.ts` (SE-UI-3, SE-DOC-2), `e2e/core-engine.spec.ts`, `src/engine/worksheet-data.test.ts`, `src/document/insertion.test.ts`, `src/model/workbook.test.ts`.

## SE-LABS-1 — One-command labs

`npm run labs` installs any missing or stale locked dependencies, then builds and serves labs plus a real server calculation endpoint. It requires no application server, database or consumer repository checkout. Host and port are configurable. The harness serves only browser assets and bounds request bodies. CI must exercise actual browser edits, HTTP calculation and both insertion modes.

Following ui-components' lab policy, the launcher shall try the requested port and up to 19 subsequent ports on `EADDRINUSE`, without exceeding port 65535. It shall bind the actual API-capable server rather than probing and releasing a port, and print the selected URL. Other listen errors fail immediately. `--strict-port` shall disable fallback, including in the Playwright web-server command; `--port 0` shall use one operating-system-assigned port. CLI host/port flags override `LABS_HOST`/`LABS_PORT` defaults. Retrying must preserve request handlers and remove temporary listen/error listeners on every attempt. The static ui-components launcher is not used here because the harness also owns calculation/Excel endpoints and must keep server bundles private.

Implementation: `scripts/labs.mjs`, `scripts/lib/labs-server.mjs`, `scripts/build-labs.mjs`, `labs/src/main.tsx`.

Verification: `src/tooling/labs-server.test.ts`, `e2e/editor.spec.ts`, launched by the Playwright web-server configuration through `npm run labs -- --strict-port`.

## SE-LABS-2 — Installation diagnostics and portable startup

A fresh source checkout shall start with `npm run labs`, without a preceding install command. Node 22.13+ within Node 22 or Node 24+, npm and Git remain system prerequisites. `.nvmrc` selects Node 24 locally and in CI, and `.npmrc` rejects unsupported runtimes during installation. The launcher and standalone labs build shall inspect mandatory installed packages and npm's installed lockfile against the committed lockfile, verify prepared UI assets and exercise esbuild's native binary in a separate process. Missing, stale, incomplete or wrong-platform installs shall trigger one `npm ci --include=dev --include=optional --ignore-scripts=false` before building; production/omit/ignore-scripts defaults must not prevent setup. Healthy installs shall be reused without network installation. The installer shall run the current Node executable with npm's JavaScript CLI path, including on Windows and in paths containing spaces, without a shell or npm.cmd.

Installation errors retain npm's output and exit code and stop startup without retry loops. A successful install must pass readiness checks before build or serving. Actual source/build failures remain visible and do not trigger dependency reinstall. The public UI library requires no private token. Git package preparation continues to launch the current Node executable directly.

Implementation: `scripts/lib/labs-dependencies.mjs`, `scripts/lib/labs-setup.mjs`, `scripts/labs.mjs`, `scripts/build-labs.mjs`, `scripts/prepare.mjs`.

Verification: `src/tooling/labs-startup.test.ts`, existing browser labs tests.

## Extraction boundary

This stage extracts reusable grid and parser modules and provides an independent rational calculation engine, complete editor chrome, comments, tracked suggestions and host-neutral integration APIs. Existing consumers are migrated separately. Storage, authorization, collaboration transports, durable history and application lifecycle services stay outside this package. Hosts may supply extra workflow/assistant panes. Curated rational unit/quantity evaluation and optional Node XLSX conversion are included; application-specific external formulas remain host-provided. [HOST-COMPATIBILITY.md](HOST-COMPATIBILITY.md) records the reviewed requirements and host boundaries. Row/column insertion/deletion rewrites references and moves review anchors within the package.

## SE-LABS-3 — Independent configurations

The labs index shall link to minimal, embedded, full editor, host integration and large worksheet pages. Each page uses the public ui-components lab bar and settings controls; selections follow navigation. Minimal provides the editable grid and worksheet strip, embedded provides compact tools inside an application-owned workspace, full provides all menus and review commands, host integration demonstrates lifecycle and authoritative result contracts, and the large worksheet holds 10,000 cells (or `?rows=` rows of 25 columns) for the grid budgets in SE-PERF-1. Reference presentation must keep formula entry, address and formatting controls usable at desktop/tablet/mobile widths without horizontal page overflow. Generic field labels remain accessible while compact chrome omits visible labels and empty feedback rows.

Verification: `e2e/labs.spec.ts` at 1440, 768 and 390 pixels, plus `e2e/editor.spec.ts`.

## SE-PERF-1 — Grid work scales with what changed

A worksheet of 10,000 cells stays smooth to type into and edit, and one of 50,000 cells stays usable. The grid's per-edit work is limited to the cells that changed:

- Typing in a cell rewrites only the draft cell and the cells whose value or readiness changed; unchanged cells keep their text, presentation and DOM nodes.
- A row whose values, inputs, formats and presentation are the same objects as at the last paint, and that holds no draft or open cell editor, is not visited, unless the engine has rewritten one of its cells. The editor gives a values row a new identity whenever a cell in it changes value or readiness.
- Committing an edit updates changed values in place. The grid rebuilds its rows only when its extent changes or more than 2,000 cells change at once.
- Zoom and authored styles are applied to every cell only when zoom changes; otherwise only to cells whose authored style changed and cells the engine has newly created. At 100% the stylesheet's own text metrics apply without inline copies.
- Shared-border reconciliation visits only cells with an inline border or a clipboard edge and their right and bottom neighbors. Unchanged column, row and frozen-column measures are not rewritten.
- Native change capture compares cell data directly instead of serializing the worksheet.

The grid does not window its rows; that is decided once the model's own per-edit work is reduced and these budgets are measured again.

Implementation: `src/engine/cell-renderer.ts`, `src/engine/worksheet-presentation.ts`, `src/engine/shared-grid-borders.ts`, `src/engine/worksheet-data.ts`, `src/engine/grid-geometry.ts`, `src/engine/workbook-grid.tsx`, `labs/src/large-lab.tsx`.

Verification: `e2e/grid-budget.spec.ts` (mutation and time budgets on the large worksheet), `src/engine/grid-budget.test.ts`, `src/engine/cell-renderer.test.ts`, `src/engine/shared-grid-borders.test.ts`, `src/engine/worksheet-data.test.ts`, `src/engine/grid-geometry.test.ts`.

## SE-LABS-4 — Shared lab shell

The labs shall use `mountLabBar` and `bootLabSettings` from `@flamboyantengineer/ui-components/labs`. The public bar renders in its own shadow root, keeping its controls usable when custom page styles change component appearance. The head bootstrap applies stylesheet and theme settings before the editor bundle runs. Reference, Unstyled and Custom URL stylesheets; System, Light, Dark and custom theme values; and configurable root theme attributes use the upstream settings contract. The legacy `stylesheet=unstyled` query maps to `none`.

The lab host supplies a safe, namespaced storage adapter. Missing or refused browser storage never prevents startup or editing. Settings persist across navigation/reload, explicit query parameters take precedence, and System follows color-scheme changes. Navigation links carry current settings, including custom URLs and theme attributes. Package editor source remains independent of the lab shell and browser storage. One-command dependency readiness checks include the prepared `/labs` export.

Implementation: `labs/src/main.tsx`, `labs/src/boot.ts`, `labs/src/settings.ts`, `scripts/build-labs.mjs`, `scripts/lib/labs-dependencies.mjs`.

Verification: `e2e/labs.spec.ts` (SE-LABS-4), `src/tooling/labs-startup.test.ts`.

## SE-HOST-MODEL-1 — Canonical authored values and logical identities

Version-1 worksheets optionally carry `inputValues`, `rowIds` and `columnIds`. Canonical inputs are validated rational/quantity values attached to existing nonformula authored cells. Their primitive `data` remains the edit representation, including text in a different input unit. Replacement and clearing discard the canonical override; formatting preserves it. JSON, internal clipboard, structure, review baselines, snapshots and live document insertion retain canonical values. Metadata addresses count once toward the cell budget.

Line identity lists, when supplied, contain unique nonempty bounded strings and cover the effective logical extent. `worksheetExtent(sheet, workbook?)` includes data, explicit minimum dimensions, geometry, styles, formats, canonical inputs, provenance, merges and available review anchors. Structural changes update the saved extent even when the affected final line is blank. Insertions receive explicit generated IDs; deletion removes only affected IDs and retains at least one line. `materializeWorksheetExtent` gives editor commands their initial 26-by-100 viewport without forcing every headless model to allocate a blank matrix; when that padding would take the sheet past the cell limit (a tall, narrow sheet such as 10,000 by 20), the command works on the logical extent instead. Hosts supply `WorkbookIdFactory` for deterministic entity generation.

Implementation: `src/model/workbook.ts`, `src/model/extent.ts`, `src/model/structure.ts`, `src/model/editing.ts`.

Verification: `src/model/host-model.test.ts` (SE-HOST-MODEL-1), `src/model/structure.test.ts`.

## SE-HOST-MODEL-2 — Semantic command records

`WorkbookChangeContext` records a transaction ID, source, detached baseline and ordered `WorkbookOperation` list. Operations explicitly identify cells, lines, sheet addition/deletion/rename/order, geometry, merges, annotations and deliberate import/new/reset replacement. A line operation comes from its structural command, including when two blank matrices are identical; it is never inferred from cell differences. Insert/delete records carry line IDs and surviving neighbors. Cell records carry authored value, canonical input, style, format and contributor provenance, with stable line IDs when available. Rewritten cross-sheet references and annotation anchors belong to the same command result. Removed authored cells and sparse geometry remain available to `invertWorkbookOperations` for standalone undo.

`collectCellOperations` visits authored or metadata coordinates, or a caller-supplied bounded target set. It does not construct an implicit grid of blank cells. `changeWorkbookStructureWithOperations` and `writeCellsWithOperations` return the workbook and its command records together. The editor owns publication and rollback, and the host owns collaboration preflight and durable storage.

Implementation: `src/model/operations.ts`.

Verification: `src/model/host-model.test.ts` (SE-HOST-MODEL-2, SE-HOST-MODEL-4).

## SE-HOST-MODEL-3 — Authored clipboard blocks

Internal copy snapshots primitive expressions and supported style, format, canonical input and contributor provenance. A copied block tiles a destination only when both destination dimensions are exact multiples of the source. Smaller and nonmultiple destinations receive one complete block from their top-left. Every copied formula translates against its original source coordinates, preserving absolute axes. A cut moves its block once, clears the source atomically and rejects a stale source before mutation. Foreign delimited text is bounded by transport, cell count, column count and field length limits before workbook mutation.

Implementation: `src/model/clipboard.ts`, `src/model/editing.ts`.

Verification: `src/model/host-model.test.ts` (SE-HOST-MODEL-3, SE-HOST-MODEL-4), `src/model/clipboard.test.ts`.

## SE-PERF-1 — Work per edit and per keystroke

A workbook of 10,000 cells stays smooth to edit and one of 50,000 cells stays usable, with client calculation on the main thread. The editor's own work for one committed edit, and for one formula-bar keystroke, does not grow with every cell of the workbook:

- A command names the cells it writes; its cell patches compare only those cells. A command that names none is compared by value in place, without copying or serializing unchanged cells.
- What the editor produced itself is not validated or copied again: a client calculation result is trusted, the host's frozen baseline is copied only when the host reads it, and a draft is attributed to its contributor once.
- What depends only on a calculation result (the reverse dependencies, the ready values) is computed once per result. A keystroke walks only the cells its draft cell reaches, and output rows without a pending cell keep their identity, while a row with one is always new.

While the client or server recalculates after an edit, the editor keeps showing the last ready value of every cell that does not read an edited cell, directly or through other formulas; only those readers show as pending. The calculation itself stays pending with no current result, so exports and document insertion still wait for it. A structural or naming change shows every cell as pending.

Calculation itself still recalculates the whole workbook; a calculation worker and incremental recalculation are not part of this requirement.

Implementation: `src/editor/use-editor.ts`, `src/editor/command-operations.ts`, `src/editor/use-publication.ts`, `src/editor/use-calculation.ts`, `src/model/operations.ts`, `src/calculation/publication.ts`.

Verification: `src/editor/editor-performance.test.ts` (SE-PERF-1) keeps a Node budget for a 10,000-cell and a 50,000-cell workbook.
