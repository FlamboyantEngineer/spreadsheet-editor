# Coding Guidelines

Every file in `src/`, `labs/src/` and `e2e/` follows these rules; tests also follow [TESTING-GUIDELINES.md](TESTING-GUIDELINES.md). They follow the sibling libraries' guidelines (the base component library's, the document editor's and the collaboration library's), adapted to an editor whose model, calculation and file formats run without React or a DOM. Rules marked **(enforced)** fail `npm run check`; the others are checked in review. The shared ESLint configuration (ui-components' `specs/TOOLING.md`) applies the size limits (C1) and the naming convention (N2, N3) to library source, `src/` without its tests; in labs, tests and browser tests those rules are checked in review.

## Naming

**N1. Files and directories are kebab-case,** named for what the module holds (`operation-apply.ts`, `use-find.ts`, `editor-formula-bar.tsx`). A React hook's module is `use-<thing>.ts`; a component's module is named for the component (`comments-panel.tsx` → `CommentsPanel`). **(enforced: `check:structure`)**

**N2. Functions say what they do.** Hooks are `use<Thing>`; components are PascalCase nouns; factories are `create<Thing>` (`createWorkbook`, `createMemoryHost`); pure functions start with a verb (`applyWorkbookOperations`, `formatWorksheetCell`); predicates start with `is`, `has` or `can`. Module constants are `UPPER_SNAKE_CASE` only for exported primitive defaults; everything else is camelCase. **(enforced: camelCase and PascalCase in library source; the constant rule is checked in review)**

**N3. Types are PascalCase nouns.** Props are `<Component>Props`; a callback or adapter the host implements is named for its role (`CalculationProvider`, `SpreadsheetHost`); the options of a function are `<Thing>Options`; a result a caller branches on is `<Action>Result`. **(enforced: PascalCase)**

**N4. One vocabulary.** The editor speaks of a workbook, its worksheets, cells addressed by `x` (column) and `y` (row) from zero, and a selection `x1`, `y1`, `x2`, `y2`. Authored values are what a person typed; results are what calculation produced. Props follow the base library's vocabulary: states are `is*` booleans, events are `on<Event>`, controlled pairs are `value`/`defaultValue` with `onChange`. Do not introduce synonyms.

**N5. Hooks follow the contract.** Classes are `se-<part>` and `se-<part>__<slot>`, data attributes and values are kebab-case, and every hook is declared in `specs/HOOK-CONTRACT.md`. **(enforced: `check:specs` for the hook prefix)**

## Structure

**S1. Seven places for code.**

- `src/model/` holds the workbook model: cells, formats, operations, structure changes, review and validation.
- `src/calculation/` holds exact arithmetic, formula evaluation, units, formatting and calculation providers.
- `src/document/` holds what a document needs from a workbook: inserted ranges and linked tables.
- `src/excel/` holds reading and writing Excel workbooks, through a worker process.
- `src/host/` holds the host contracts, workbook sessions and the in-memory host (`specs/HOST-ADAPTERS.md`).
- `src/engine/` holds the grid engine: the Jspreadsheet host and every module that touches its DOM.
- `src/editor/` holds the React editor chrome and the hooks behind it.

`src/tooling/` holds tests of the repository's own scripts. Nothing else lives in `src/` but the package's root `index.ts`.

**S2. React and the DOM stay in `engine/` and `editor/`.** The model, calculation, document, excel and host entry points run on a server and in a worker without React or DOM globals. Only `src/engine/` touches Jspreadsheet-owned DOM; React owns the rest of the chrome.

**S3. `index.ts` only re-exports** the public API with named `export … from` statements. No default exports anywhere (TS7). **(enforced: `check:structure`, lint)**

**S4. Every export is used.** An export that only its own file uses is not exported. **(enforced: `check:exports`; `--fix` drops them)**

**S5. No import cycles between modules.** Break a cycle by moving the shared part into its own module. **(enforced: `check:structure`)**

**S6. Imports are sorted and typed.** Relative imports end in `.js`; a module is imported in one statement, with type-only names marked inline (`import { Icon, type IconSource }`) or the whole statement `import type` when every name is a type; packages come first, then relative paths (`npx eslint --fix` sorts them). **(enforced: lint)**

**S7. Controls come from the base component library.** Buttons, fields, menus, popovers, dialogs, panels and every other generic control are imported from `@flamboyantengineer/ui-components` through a component entry, never rebuilt here. A control the editor needs and the base library lacks is requested there first.

**S8. Module order:** imports, exported types, module constants, private helpers, exported functions, components and hooks.

## Inheritance and composition

**I1. Functions, plain objects and components.** Write factory functions, plain data and function components. A class is allowed only where an instance must be one: an `Error` subclass a caller checks with `instanceof`, or private evaluator and grid state that a hot path allocates per cell or per pass. A class is never part of the public API except as an error.

**I2. Host contracts are interfaces.** The editor declares what it needs from a host as an interface (`src/host/adapters.ts`, `src/editor/contracts.ts`) and receives an implementation through props or a session; it never reaches for a global, a singleton or browser storage. **(enforced: lint rejects browser storage globals in library source)**

**I3. Specialise by composition.** A specialised component renders the general one with fixed props; shared behaviour two modules need lives in one helper or hook, never in two copies.

**I4. Large hooks are built from small ones.** A hook that grows past the size limits is split into state hooks that each own one concern, plain module functions that take the state they need (`fn(core, …)`), and thin hooks that bind those functions for the caller. Keep refs out of plain functions during render: pass them to hooks and read them in effects and handlers (the React Compiler rules below).

## React

**RE1. Code the React Compiler can optimise.** Library source and labs run every rule of `eslint-plugin-react-hooks`' recommended set at error: no refs read or written during render, no mutation of values a hook returned, no conditional hooks. **(enforced: lint)**

**RE2. A disable comment names its rule and says why,** in the same comment (`// eslint-disable-next-line react-hooks/refs -- the handle reads the latest render's lifecycle when the host calls it`). Size limits are never disabled; split the code instead.

**RE3. Components emit hooks only:** `se-` classes, data attributes and the properties documented in `specs/HOOK-CONTRACT.md`. Package source loads no stylesheet; reference presentation lives in `labs/`.

## Simplicity

**C1. Small units.** A function has at most 80 lines and a cyclomatic complexity of at most 15, nests at most three levels deep, and takes at most three parameters (pass an options object instead). A module has at most 300 lines; split helpers into sibling modules before that. **(enforced in library source)**

**C2. No speculative surface.** Every export and option is exercised by a test. Remove what nothing uses.

**C3. Comments explain why, not what.** Document public functions and types in one line of TSDoc; comment only non-obvious constraints and workarounds. No commented-out code, no narration, no decorative separators.

**C4. No dead code.** No unused exports, parameters, branches or files. **(enforced: lint for unused variables and imports, `check:exports`)**

## Numbers and formats

**F1. Authored values never pass through JavaScript `Number` arithmetic.** Exact values are reduced numerator and denominator strings; scientific input is parsed directly into integers. `Number` is for indexes, budgets and presentation dimensions only.

**F2. Formulas evaluate through the parsed syntax tree.** No `eval`, `new Function` or generated code. **(enforced: lint)**

**F3. The versioned workbook and rational JSON formats are compatibility contracts.** Change them only with a spec change that says so.

## TypeScript

Adapted from the agent-rules TypeScript set ([coding-rules/README.md](coding-rules/README.md)). They apply to every TypeScript file. The enforced ones run in `npm run lint` through the shared configuration in `eslint.config.js`, which loads the TypeScript program.

**TS1. `strict` stays on.** Do not weaken `tsconfig.json` to make a file compile.

**TS2. No `any`.** Use `unknown` and narrow it; wrap a third-party `any` at the boundary that receives it. **(enforced: `@typescript-eslint/no-explicit-any`)**

**TS3. No non-null assertions (`value!`).** Narrow the value, restructure so the type proves it, or use `requireDefined` where the invariant would break. **(enforced: `@typescript-eslint/no-non-null-assertion`)**

**TS4. No `@ts-ignore` or `@ts-nocheck`.** A `@ts-expect-error` carries a description of what removes it. **(enforced: `@typescript-eslint/ban-ts-comment`)**

**TS5. Import a type-only binding with `import type`** (S6). **(enforced: `@typescript-eslint/consistent-type-imports`, `@typescript-eslint/no-import-type-side-effects`)**

**TS6. No `enum`.** Use a string-literal union, or an `as const` object with a derived type. **(enforced)**

**TS7. Named exports only.** A default export is allowed only where a loader requires it: the Vitest and Playwright configuration files. **(enforced: `no-restricted-exports`)**

**TS8. A `switch` over a union handles every member,** or ends in a `default` that handles the rest on purpose. **(enforced: `@typescript-eslint/switch-exhaustiveness-check`)**

**TS9. No floating promises.** `await` a promise, return it, or mark a deliberately detached one with `void` and handle its rejection. Do not pass an async function where the callee expects a synchronous callback. **(enforced: `@typescript-eslint/no-floating-promises`, `@typescript-eslint/no-misused-promises`)**

**TS10. Throw and reject with `Error` objects,** never strings or plain objects. **(enforced: `@typescript-eslint/only-throw-error`, `@typescript-eslint/prefer-promise-reject-errors`)**

**TS11. When a `catch` block throws a new error, attach the caught one as `{ cause }`.** **(enforced: `preserve-caught-error`)**

**TS12. Return a discriminated result for expected failures** a caller branches on (`{ ok: true, value } | { ok: false, error }`) instead of throwing.

**TS13. Prefer `type` for unions and derived shapes,** and `interface` for object shapes callers extend, such as host contracts (N3, I2).

**TS14. Give an exported function an explicit return type** when the inferred one is a wide object or a promise the caller must branch on.

**TS15. Two compiler programs check the code** (`npm run typecheck`, `npm run build`). `tsconfig.json` checks library source, labs, browser tests and configs; `tsconfig.build.json` compiles library source without its tests. Beyond `strict`, both set `noImplicitOverride`, `noFallthroughCasesInSwitch`, `isolatedModules` and `verbatimModuleSyntax`. **(enforced)**

## Node scripts

Adapted from the agent-rules Node.js set. They apply to `scripts/` and the configuration files; library source reads no `process.env`.

**R1. Target the Node version in `package.json` `engines`** (Node 22.13 or later, or Node 24).

**R2. Read options from arguments,** and environment variables only at the script's entry point, with a default.

**R3. Fail loudly.** A check exits non-zero with a message naming the file and the rule it broke; an unexpected failure stays an `Error` with its original `cause`.

**R4. Scripts that write files can run twice with the same result,** and a script that deletes files deletes only build output.

## Security

Adapted from the agent-rules security set. They apply everywhere.

**SEC1. Every imported file is untrusted.** The Excel reader checks a workbook's size and structure before it parses it, and runs the parse in a worker it can stop. Bound what any input may allocate.

**SEC2. No `eval`, `new Function` or string timers** (F2).

**SEC3. Network access belongs to the host.** The editor loads no fixed application URL; a remote calculation provider receives its endpoint and credentials from the host and never stores or logs them.

**SEC4. Keep secrets and tokens out of the repository, out of logs and out of browser storage.**

**SEC5. Pass shell arguments as arrays** (`execFileSync(command, args)`), never as a concatenated command string.

**SEC6. Add dependencies through `npm`** so the lockfile records them.

## Design

**D1. Design for less complexity.** New and touched code follows [coding-rules/a-philosophy-of-software-design.mini.md](coding-rules/a-philosophy-of-software-design.mini.md), vendored unchanged from agent-rules-books. It is review guidance, not lint. In short: measure a design by the complexity a reader must carry, prefer deep modules with small interfaces over pass-through layers, keep each decision inside the module that owns it, define invalid states away rather than making callers defend against them, and use names and comments to state contracts, not mechanics. Where it overlaps a numbered rule, the numbered rule wins.
