# Agent Instructions

This repository is a standalone, embeddable React spreadsheet editor with exact arithmetic, built on `@flamboyantengineer/ui-components`. Read [README.md](README.md), [CONTRIBUTING.md](CONTRIBUTING.md) and the specs of the area you change before changing it.

## Gates

Before every commit, run the checks that cover your change. All of them must pass:

```sh
npm run lint && npm run format:check && npm run typecheck
npm run check:references && npm run check:specs && npm run check:structure && npm run check:exports
npm run test:changed   # unit tests related to your changes since origin/dev
```

The reference, spec, structure and export checks and the shared ESLint and Prettier configuration come from ui-components' tooling (its `specs/TOOLING.md`), at the pinned commit; `tooling.config.mjs` holds this repository's settings. The shared size limits (complexity, depth, parameters, function and file length) are being adopted one area at a time: `SIZE_LIMITS_PENDING` in `eslint.config.js` lists the areas still to refactor, and a change only ever removes entries from it.

- `npm run test:changed` runs `vitest --changed` against `origin/dev`. A change to `package.json` or the Vitest config reruns everything; for those changes run the tests of what you touched by path instead (`npx vitest run src/model/workbook.test.ts`).
- Changes to exports, public types, entry points or package contents also run `npm run build && npm run check:package`.
- Changes to the grid engine, the editor chrome, a lab or a browser test (`src/engine/`, `src/editor/`, `labs/`, `e2e/`) also run `npm run build:labs` and the specs they touch (`npx playwright test e2e/<name>.spec.ts`).
- `npm run check` runs every gate, as CI does. Run it locally only when a change reaches across most of the library or when asked.

Run `npm run format` to fix formatting. Never weaken a lint rule, the denylist or a check to get green; a disable comment names its rule and says why.

## Rules

- Work on a short-lived branch from `dev` and open a pull request to `dev`. Merge it once CI passes and the branch is up to date with `dev`.
- Generic controls come from `@flamboyantengineer/ui-components`. A control it lacks is requested there, not built here. React owns the editor chrome; imperative grid DOM belongs only in `src/engine/`.
- Components emit hooks only: `se-` hook classes, data attributes and the properties documented in `specs/HOOK-CONTRACT.md`. Package source loads no CSS, fixed application URLs, global storage or application services. Reference presentation lives in `labs/`.
- Everything outside the editor is reached through explicit host callbacks or calculation providers (`specs/HOST-INTEGRATION.md`): no globals, no singletons, no direct browser storage.
- Never name a host application or its domain terms anywhere in the repository, including commit messages and pull request text. Use neutral sample data in labs and tests.
- Every behaviour change updates its spec (`specs/`) and its TypeScript tests, which name the `SE-` requirement ids; [specs/README.md](specs/README.md) gives the format.
- The versioned workbook and rational JSON formats are compatibility contracts: keep them unchanged unless a spec change says otherwise. Numerical operations never pass authored values through JavaScript `Number` arithmetic.
- APIs and styling hooks are experimental until `1.0.0`: use minor changesets for API additions, removals and reorganizations, and patches for fixes that preserve the API. Remove obsolete APIs without deprecation periods or compatibility shims. Add a changeset (`npx changeset`) to every pull request that changes the package.
- Keep generated builds, browser test output, scratch files and credentials out of commits.
