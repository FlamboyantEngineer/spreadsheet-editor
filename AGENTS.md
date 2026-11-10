# Agent instructions

Read README.md, CONTRIBUTING.md and the relevant specs before changing behavior.

- Keep application services behind explicit host callbacks or calculation providers.
- Generic controls come from `@flamboyantengineer/ui-components`; React owns editor chrome. Imperative grid DOM belongs only in `src/engine/`.
- Package source emits `se-*` hooks and loads no CSS, fixed application URLs, global storage or application services. Reference presentation lives in `labs/`.
- Preserve the versioned workbook and rational JSON contracts. Numerical operations must never pass authored values through JavaScript `Number` arithmetic.
- Update the owning spec and TypeScript tests when behavior changes. Keep source provenance for extracted modules.
- Work on a scoped branch from `dev`; open a draft PR to `dev`. Do not merge without a user request.
- Before committing, run `npm run check`. Keep generated builds, browser artifacts, scratch files and credentials out of commits.
