# Generic spreadsheet editor implementation plan

The library provides exact calculations, a reusable workbook model, editor interactions and explicit host integration contracts. Application adapters and consumer migrations are outside this repository.

## 1. Compatibility and numeric values

- Preserve version-one workbook/result schemas and primitive CellInput. Add optional fields; legacy callbacks/providers remain valid.
- Keep exact rationals, scientific inputs such as 1E-25, cancellation, 32-digit default display, isolated 80-digit approximations and intentional #LIMIT! behavior.
- Represent canonical inputs separately from edit text; validate rational limits, temperature tags, physical kinds and opaque provenance at every ingress.
- Support absolute/difference temperatures, scalar kind propagation, complete unit aliases and prefixes, explicit word-size context, and case-insensitive expected-dimension categories.
- Verify function behavior through portable fixtures, including TEXT masks, CONVERT aliases, INDEX shapes, Windows-1252, calendar day 60 and exact locale/currency formatting.

## 2. Model and semantic transactions

- Preserve canonical values through formatting, clipboard, structural changes, suggestions, undo and document insertion. Replacement clears them.
- Keep stable row/column identities covering effective logical extent; permit deterministic ID factories.
- Compute semantic cell, line, sheet, geometry, merge, annotation and explicit replacement operations at command time. Blank line commands retain their intent.
- Emit one local callback per command with transaction ID, immutable baseline and ordered operations. Remote controlled values emit no local changes.
- Reject synchronous host exceptions before accepting the workbook/history/draft. Restore presentation and propagate failure through lifecycle barriers.
- Delegate undo/redo when supplied; retain standalone history otherwise.

## 3. Public lifecycle and calculation publication

- Expose a stable session handle for draft commit/cancel, detached snapshots, selection, state subscriptions, Save and Flush.
- Commit active native/formula-bar input once before persistence. Flush drains scheduled saves and calls the host barrier without creating a manual history snapshot.
- Preserve dirty state on failure and after edits newer than a captured receipt. Remote replacement alone is clean. Disposal invalidates handles and late completions; providers remain host-owned.
- Add external calculation mode with current input key, monotonic revision, global/per-cell readiness, bounded host JSON and no local fallback.
- Export a deterministic calculation key excluding visual/annotation changes. Validate result shape and metadata before display or export.
- Record actual cell/extent dependencies. Drafts invalidate their cell and transitive dependents without evaluation; unknown dependencies invalidate conservatively.

## 4. Generic presentation and editing

- Merge bounded host function descriptors with built-ins in both entry surfaces. Reject duplicate names/aliases and preserve completion/IME priority.
- Accept presentation only for matching key/revision and current ready cells. Honor explicit authored formats before suggested units.
- Mount optional React cell content through engine-owned portals; clear stale titles, descriptions, decorations and content.
- Synchronize vendor logical extents for structure, undo and controlled values without replacing ordinary editor roots.
- Tile copied blocks only for exact destination multiples; offset each formula against the original source. Cut moves once and clears atomically.
- Show protected drafts immediately, preserve unrelated results, and search authored/current formatted text. Replace authored matches only.
- Coalesce pointer selection to one publication per frame plus final release/cancel/blur; avoid frozen-style work on selection. Restore per-sheet selection, scroll and focus.

## 5. Excel and host boundaries

- Use serializable host-function capabilities, defaulting to none. Bound worker requests, names, policies, captured text and cancellation.
- Export immutable captured operands without live resolution. Use a maintained HTML parser for explicit HTML-to-text policy.
- Preserve exact numeric text and supported locale/currency metadata. Report unrepresentable canonical/quantity/provenance metadata and missing contexts explicitly.
- Keep React/ui-components responsible for chrome, imperative grid DOM inside src/engine, CSS in host/labs, and all application services behind callbacks.

## 6. Delivery and verification

- Keep owning specifications, portable fixtures, public types and a generic host example current.
- Verify model/calculation/document/root imports without DOM globals and compile a built-package modern/legacy consumer.
- Exercise native/formula entry, drafts, external publications, presentation cleanup, search, IME, permissions, logical extents and per-sheet restoration in browser tests.
- Run npm run check before committing; it covers lint, formatting, types, units, build, package smoke/type fixtures and browser suites.
- Deliver one scoped library draft PR to dev. Consumer repositories are unchanged. Do not merge.
