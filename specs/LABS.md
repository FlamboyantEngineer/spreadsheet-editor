# Labs

## SE-LABS-1 - One-command labs

### SE-LABS-1.1 - One-command labs

`npm run labs` installs any missing or stale locked dependencies, then builds and serves labs plus a real server calculation endpoint. It requires no application server, database or consumer repository checkout. Host and port are configurable. The harness serves only browser assets and bounds request bodies. CI must exercise actual browser edits, HTTP calculation and both insertion modes.

Following ui-components' lab policy, the launcher shall try the requested port and up to 19 subsequent ports on `EADDRINUSE`, without exceeding port 65535. It shall bind the actual API-capable server rather than probing and releasing a port, and print the selected URL. Other listen errors fail immediately. `--strict-port` shall disable fallback, including in the Playwright web-server command; `--port 0` shall use one operating-system-assigned port. CLI host/port flags override `LABS_HOST`/`LABS_PORT` defaults. Retrying must preserve request handlers and remove temporary listen/error listeners on every attempt. The static ui-components launcher is not used here because the harness also owns calculation/Excel endpoints and must keep server bundles private.

Implementation: `scripts/labs.mjs`, `scripts/lib/labs-server.mjs`, `scripts/build-labs.mjs`, `labs/src/main.tsx`.

Verification: `src/tooling/labs-server.test.ts`, `e2e/editor.spec.ts`, launched by the Playwright web-server configuration through npm run labs with --strict-port.

## SE-LABS-2 - Installation diagnostics and portable startup

### SE-LABS-2.1 - Installation diagnostics and portable startup

A fresh source checkout shall start with `npm run labs`, without a preceding install command. Node 22.13+ within Node 22 or Node 24+, npm and Git remain system prerequisites. `.nvmrc` selects Node 24 locally and in CI, and `.npmrc` rejects unsupported runtimes during installation. The launcher and standalone labs build shall inspect mandatory installed packages and npm's installed lockfile against the committed lockfile, verify prepared UI assets and exercise esbuild's native binary in a separate process. Missing, stale, incomplete or wrong-platform installs shall trigger one `npm ci --include=dev --include=optional --ignore-scripts=false` before building; production/omit/ignore-scripts defaults must not prevent setup. Healthy installs shall be reused without network installation. The installer shall run the current Node executable with npm's JavaScript CLI path, including on Windows and in paths containing spaces, without a shell or npm.cmd.

Installation errors retain npm's output and exit code and stop startup without retry loops. A successful install must pass readiness checks before build or serving. Actual source/build failures remain visible and do not trigger dependency reinstall. The public UI library requires no private token. Git package preparation continues to launch the current Node executable directly.

Implementation: `scripts/lib/labs-dependencies.mjs`, `scripts/lib/labs-setup.mjs`, `scripts/labs.mjs`, `scripts/build-labs.mjs`, `scripts/prepare.mjs`.

Verification: `src/tooling/labs-startup.test.ts`, existing browser labs tests.

## SE-LABS-3 - Independent configurations

### SE-LABS-3.1 - Independent configurations

The labs index shall link to minimal, embedded, full editor, host integration and large worksheet pages. Each page uses the public ui-components lab bar and settings controls; selections follow navigation. Minimal provides the editable grid and sheet bar, embedded provides compact tools inside an application-owned workspace, full provides all menus and review commands, host integration demonstrates lifecycle and authoritative result contracts, and the large worksheet holds 10,000 cells (or `?rows=` rows of 25 columns) for the grid budgets in SE-PERF-1.1. Reference presentation must keep formula entry, address and formatting controls usable at desktop/tablet/mobile widths without horizontal page overflow. Generic field labels remain accessible while compact chrome omits visible labels and empty feedback rows.

Implementation: `labs/src/main.tsx`, `labs/src/full-lab.tsx`, `labs/src/host-lab.tsx`, `labs/src/large-lab.tsx`, `labs/src/excel-lab.tsx`.

Verification: `e2e/labs.spec.ts` at 1440, 768 and 390 pixels, plus `e2e/editor.spec.ts`.

## SE-LABS-4 - Shared lab shell

### SE-LABS-4.1 - Shared lab shell

The labs shall use `mountLabBar` and `bootLabSettings` from `@flamboyantengineer/ui-components/labs`. The public bar renders in its own shadow root, keeping its controls usable when custom page styles change component appearance. The head bootstrap applies stylesheet and theme settings before the editor bundle runs. Reference, Unstyled and Custom URL stylesheets; System, Light, Dark and custom theme values; and configurable root theme attributes use the upstream settings contract. The legacy `stylesheet=unstyled` query maps to `none`.

The lab host supplies a safe, namespaced storage adapter. Missing or refused browser storage never prevents startup or editing. Settings persist across navigation/reload, explicit query parameters take precedence, and System follows color-scheme changes. Navigation links carry current settings, including custom URLs and theme attributes. Package editor source remains independent of the lab shell and browser storage. One-command dependency readiness checks include the prepared `/labs` export.

Implementation: `labs/src/main.tsx`, `labs/src/boot.ts`, `labs/src/settings.ts`, `scripts/build-labs.mjs`, `scripts/lib/labs-dependencies.mjs`.

Verification: `e2e/labs.spec.ts` (SE-LABS-4.1), `src/tooling/labs-startup.test.ts`.

### SE-LABS-4.2 - shadcn/ui stylesheet

The labs shall offer the shared settings' shadcn/ui stylesheet beside the reference one (`?stylesheet=shadcn`). It loads the package's `shadcn.css`: ui-components' shadcn/ui sheet, the editor's own rules (`labs/src/editor.css`, the same file the reference presentation loads after ui-components' reference sheet) unchanged on the tokens that sheet's bridge points at shadcn/ui's, then a short retouch of the editor chrome (`labs/src/shadcn-retouch.css`): the editor frame as a rounded hairline card on the page background, a hairline under the chrome, the worksheet tabs as a pill list flush under the grid without a rule under them, side panels on the sidebar color and suggestion cards on the shadcn/ui radius. Light and dark follow `data-theme` and shadcn/ui's `.dark` class. Every lab page carries `data-ui-scale="compact"` on its root element, so the editor chrome and its menus, popovers and dialogs take the compact control scale under the shadcn/ui sheet; the reference sheet ignores the attribute. The retouch never sizes grid cells, rows or columns: the engine's measurements are the same under both sheets.

Implementation: `labs/src/shadcn.css`, `labs/src/shadcn-retouch.css`, `labs/src/editor.css`, `labs/src/reference.css`, `labs/src/boot.ts`, `labs/src/main.tsx`, `scripts/build-labs.mjs`, `scripts/build.mjs`.

Verification: `e2e/labs.spec.ts` (SE-LABS-4.2); the package check imports and bundles the shadcn/ui stylesheet from the packed package.
