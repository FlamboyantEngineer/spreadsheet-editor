# Specs

Each file in this directory specifies one area of the editor or one shared contract. Specs are the source of truth for behaviour and for styling hooks; tests and labs verify them.

| Document                                       | Owns                                                                 |
| ---------------------------------------------- | -------------------------------------------------------------------- |
| [EDITOR.md](EDITOR.md)                         | The React editor, calculation location and authored data.            |
| [EDITOR-CHROME.md](EDITOR-CHROME.md)           | Menus, toolbar, formula bar, sheet bar, zoom and the context menu.   |
| [CALCULATION.md](CALCULATION.md)               | Exact values, evaluation, server transport, quantities and formats.  |
| [PERFORMANCE.md](PERFORMANCE.md)               | Grid and editor work budgets.                                        |
| [REVIEW.md](REVIEW.md)                         | Comments and suggestions.                                            |
| [HOST-ADAPTERS.md](HOST-ADAPTERS.md)           | The host object: identity, capabilities, stores, presence, versions. |
| [SESSION.md](SESSION.md)                       | The workbook session that opens and saves a stored workbook.         |
| [HOST-INTEGRATION.md](HOST-INTEGRATION.md)     | Accepted changes, session lifecycle, publications and the catalog.   |
| [HOST-MODEL.md](HOST-MODEL.md)                 | Canonical values, line identities, command records and clipboard.    |
| [HOST-COMPATIBILITY.md](HOST-COMPATIBILITY.md) | Implemented differences and integration boundaries.                  |
| [DOCUMENT-INSERTION.md](DOCUMENT-INSERTION.md) | Document table payloads and linked capture.                          |
| [EXCEL.md](EXCEL.md)                           | The optional Excel import and export entry.                          |
| [LABS.md](LABS.md)                             | The labs launcher, configurations and lab shell.                     |
| [HOOK-CONTRACT.md](HOOK-CONTRACT.md)           | The `se-*` hooks and what hosts may style.                           |

## File format

- File names are `UPPER-KEBAB.md`, and the first line is a `# Title` heading.
- Requirements are grouped under `## SE-<AREA>-<n> - <Group title>` headings. Each requirement is a `### SE-<AREA>-<n>.<m> - <Title>` heading followed by its statement, then two lines:

  ```markdown
  Implementation: `src/calculation/rational.ts`

  Verification: `src/calculation/evaluator.test.ts`, `e2e/editor.spec.ts`
  ```

  - Each file uses one `<AREA>`.
  - `Implementation:` lists the files that implement the requirement, each in backticks.
  - `Verification:` lists the tests that verify it, each in backticks. At least one of those tests names the requirement ID in a test title or comment. Checks that are not test files, such as the package check, are described in plain text.
  - Within a group, requirements are numbered from 1 without gaps and appear in number order, and groups appear in ascending order (gaps between group numbers are allowed).
  - IDs are never reused. A withdrawn requirement keeps its heading with the statement `Withdrawn.` and its last verification.

`npm run check:specs` enforces this format, the existence of every listed path, that the listed tests name each ID, and that every relative link in `specs/` and `docs/` reaches its file and heading.
