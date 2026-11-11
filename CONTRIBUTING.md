# Contributing

Use Node 22 or newer, `npm ci`, and a scoped branch from `dev`. The public `@flamboyantengineer/ui-components` library from [FlamboyantEngineer/ui-components](https://github.com/FlamboyantEngineer/ui-components) is pinned by Git SHA in the development manifest and lockfile; hosts supply it as a peer. Keep React as a peer to avoid multiple renderer instances.

Run `npm run labs` to install or repair its locked dependencies, build and serve the standalone demos, including the server evaluator. Changes require `npm run check` before commit. Use TypeScript tests beside source and browser tests under `e2e/`; test observable behavior and name the owning `SE-*` requirement. Update specifications and public type documentation with contract changes.

Keep the model, calculation and document entry points usable without React or DOM globals. Exact values use reduced numerator/denominator strings; scientific input is parsed directly into integers. JavaScript numeric arithmetic is for indexes, budgets and presentation dimensions only. Decimal approximation contexts are local to a calculation. Do not use dynamic code generation to evaluate formulas.

All generic UI controls use the base component library. React owns the toolbar, formula bar, sheet buttons and status. Only `src/engine/` may touch Jspreadsheet-owned DOM. Library source loads no stylesheet and contains no application routes or persistence. Style hooks are documented in `specs/HOOK-CONTRACT.md`; reference CSS belongs in labs.

The package is experimental until 1.0. Keep serialized model versions explicit and document incompatible changes. The initial extraction deliberately leaves consumer migration to separate changes. Source modules and their origin commit are recorded in NOTICE.md.

The base component repository is public. Local installs and CI need no private-library token. Never put credentials in manifests or lockfiles.
