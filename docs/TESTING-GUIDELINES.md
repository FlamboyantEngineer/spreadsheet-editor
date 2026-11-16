# Testing Guidelines

Every unit test in `src/` and every browser test in `e2e/` follows these rules, on top of the [coding guidelines](CODING-GUIDELINES.md). Tests verify the specs: they drive the library through its public functions, its host contracts and what a person sees and does, and they name the requirement they check. Rules marked **(enforced)** fail `npm run check`; the others are checked in review.

## Files and names

**T1. TypeScript next to the source.** A unit test is `<module>.test.ts`, or `.test.tsx` when it renders, beside the module it tests (`operations.ts` → `operations.test.ts`). Browser tests are `e2e/<area>.spec.ts` and run against the labs. **(enforced: Vitest only collects `src/**/*.test.ts(x)`)**

**T2. Test names start with the requirement IDs they verify**, then say what a caller or person observes: `it('SE-CALC-1.1 keeps one third exact through a sum', …)`. A `describe` title may carry the IDs its tests share. A test may name several IDs; an ID may be split over several tests. `check:specs` fails when no test listed under a requirement's `Verification:` names its ID ([specs/README.md](../specs/README.md)). **(enforced)**

## Environments

**T3. Tests run in Node unless they render.** A test that renders or needs DOM globals starts with `// @vitest-environment jsdom`; everything else, including the model, calculation, document, excel and host tests, runs in Node, which keeps those entry points free of DOM globals (S2).

**T4. Test hooks with `renderHook` and components with `render`.** Drive the editor hook (`useSpreadsheetEditor`) through `renderHook` and `act` when the behaviour is its state; render the component when the behaviour is what a person sees. The Jspreadsheet grid needs a real browser: grid behaviour is tested in `e2e/`.

## Queries and interaction

**T5. Query through `screen` by role, label or text**, narrowed with `within(element)`. A role and name query also checks that the element is exposed to assistive technology.

**T6. Hooks are public API, so tests may read them.** `se-` classes and data attributes check the styling contract (`specs/HOOK-CONTRACT.md`). Do not reach into markup that is not a hook or an ARIA relationship.

**T7. No test ids.** Query by role, label, text or a documented hook, never `getByTestId`. A `data-testid` may appear only on something the test passes in, such as a host-provided icon, to check that the editor renders it.

**T8. Act through `userEvent`** for pointer and keyboard sequences a person makes (`const user = userEvent.setup()`, then await every call). Use `fireEvent` only for what user-event cannot produce or where the test checks one exact key event, such as an input method's composing key (key code 229).

**T9. Wait for what settles later** with `findBy` queries and `waitFor`, and use plain `getBy` queries for everything synchronous. Never sleep for a fixed time.

## Speed and isolation

**T10. Tests are parallel-safe and order-independent.** Each test builds its own workbook, host and fixtures (`createWorkbook`, `createMemoryHost`). Restore anything a test changes (fake timers, stubbed globals, listeners on `window` or `document`), write temporary files under `mkdtemp`, and never depend on another test's state. The suite must pass with any `--maxWorkers`.

**T11. Control time with Vitest's fake timers,** never with real waits. Debounces, retries and timeouts are driven by advancing the clock; restore real timers after the test.

**T12. Keep each test under the 5 second default on a loaded machine.** Split a long test into focused ones that start from the state they need.

## Browser tests

**T13. Open a lab, then find the editor by its hooks and roles.** Browser tests run against `npm run labs` on port 4173, one worker at a time, with the browser from `CHROMIUM_PATH` when Playwright's own is not installed. Start from the lab page that shows the behaviour, locate the editor by `.se-editor` and its data attributes, and then by role.

**T14. Grid cells are located by the engine's cell attributes** (`td[data-x="0"][data-y="1"]`). The grid renders only the rows near its viewport (`specs/PERFORMANCE.md`), so scroll a target row into view before acting on it.

**T15. Retry values that settle with `toPass`** or a web-first assertion (`toHaveAttribute`, `toBeVisible`): scroll positions, measured sizes, calculated results and anything that follows animation frames.

## Assertions

Adapted from the agent-rules testing set ([coding-rules/README.md](coding-rules/README.md)).

**T21. Assert what a caller or person can observe:** returned workbooks and operations, calculated values, rendered roles, names and hooks, and the calls a host callback receives. Do not assert private function names, call counts or internal call order unless the spec requires them.

**T22. Mock at a module or network boundary,** never the module under test. Prefer the in-memory host and an injected `fetch` to mocks.

**T23. A snapshot is never the only assertion.**

**T24. Build test data with a small factory or a shared fixture** instead of copying a large fixture between files. Exact values are written as the strings the format stores, never computed with `Number` arithmetic (F1).

**T25. Cover the failure paths a change introduces or touches:** rejected edits, read-only and locked cells, conflicting operations, invalid formulas, corrupt files and unavailable calculation, where they exist.

**T26. Never weaken an assertion or skip a test to make it pass.** Fix the code, or change the spec requirement and its test together. A known gap the spec records stays visible as `it.fails` with a comment naming the requirement, never as `.skip`.
