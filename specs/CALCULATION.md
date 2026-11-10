# Calculation contract

## SE-CALC-1 — Exact values

Numeric literals, including exponent notation, parse from strings into reduced rational values. Exact results serialize as `{ kind: 'number', numerator: string, denominator: string, approximate: false }`, with a positive denominator. Arithmetic and comparisons never convert these values through binary floating point. `1E-25`, `1-1E-25`, `1-(1-1E-25)`, `1/3*3` and cancellation at `1E-250` remain exact. JavaScript numeric input is accepted for compatibility but cannot restore digits already lost by its caller.

Display rounds a copy to configurable significant digits (32 by default), with decimal, automatic or fraction notation. Explicit rounding functions change the value using their documented direction; ROUND uses halves away from zero. No display option changes subsequent calculations.

Non-rational operations use a per-calculation decimal context (default 80 significant digits, supported range 16–1000). They serialize an approximate rational decimal with `approximate: true` and `precision`. Dependent arithmetic retains the approximation marker. The display prefixes these results with `≈` unless disabled. Perfect rational square roots and integer powers remain exact.

Implementation: `src/calculation/rational.ts`, `src/calculation/values.ts`, `src/calculation/format.ts`.

Verification: `src/calculation/evaluator.test.ts`, `e2e/editor.spec.ts` (SE-UI-1, SE-UI-4).

## SE-CALC-2 — Workbook evaluation

Formulas parse into a syntax tree; no authored formula reaches JavaScript evaluation or code generation. Cross-sheet references are case-insensitive, quoted names and absolute addresses are supported, and whole-row/column ranges are bounded by the used dimensions. Cells are memoized within one calculation. Cycles, invalid references, numerical errors, unknown functions and work-limit failures produce typed error values. IF and IFERROR evaluate only the selected branch. Hosts may resolve names and synchronous custom functions explicitly.

The exported `SUPPORTED_FUNCTIONS` array is the executable catalog. The initial families below are retained; the executable catalog now also covers all 154 supported worksheet names and aliases from the pinned Host registry in [HOST-FUNCTIONS.json](HOST-FUNCTIONS.json). Additions include conditional aggregates, lookups, calendar/text functions, hyperbolic math and stable EXPM1/LOG1P. Pending registry entries remain unsupported. Families:

- Logic: IF, IFERROR, AND, OR, NOT, TRUE, FALSE, ISERROR, NA.
- Aggregates: SUM, PRODUCT, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, SUMSQ.
- Arithmetic: ABS, SIGN, SQRT, POWER, MOD, ROUND, ROUNDUP, ROUNDDOWN, INT, TRUNC, CEILING, FLOOR, QUOTIENT.
- Approximate math: PI, SIN, COS, TAN, ASIN, ACOS, ATAN, EXP, LN, LOG, LOG10.
- Statistics: VAR.S, VAR.P, STDEV.S, STDEV.P.
- References: ROW, COLUMN, ROWS, COLUMNS.
- Text and type checks: LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM, CONCAT, CONCATENATE, VALUE, ISNUMBER, ISTEXT, ISBLANK.

Default bounds are 250,000 populated cells, 200 worksheets, 10,000 characters per input cell, 10,000 digits per numerator/denominator, 2,000,000 evaluation operations and expression depth 128. Dependency chains use an explicit resumable stack; each formula is separately limited to 10,000 cumulative addressed cells before range allocation. Numeric size and evaluation limits bound resource usage rather than introducing rounding. The synchronous API permits explicit limits within the arithmetic engine's maximums. Over-budget cell computations return `#LIMIT!`. A malformed or oversized workbook is rejected before evaluation.

Implementation: `src/model/formula-parser.ts`, `src/calculation/evaluator.ts`, `src/model/workbook.ts`.

Verification: `src/calculation/evaluator.test.ts`, `src/model/workbook.test.ts`.

## SE-CALC-3 — Server transport and parity

`CalculationProvider.calculate(workbook, options, signal)` returns a versioned result. The HTTP provider posts a versioned request with a per-provider request id, and validates the response id, workbook dimensions, error tokens and rational fields. Exact numerator/denominator strings must survive transport unchanged.

`handleCalculationRequest` accepts Fetch-standard Requests and returns Responses. It accepts POST JSON only, bounds streaming request bodies (2 MiB by default), validates inputs and caps requested limits at server ceilings. Client response bodies are bounded (16 MiB by default). Authentication and routing are supplied by the host. No server state or global arithmetic context is shared between requests.

Implementation: `src/calculation/provider.ts`, `src/calculation/http.ts`.

Verification: `src/calculation/http.test.ts`, `e2e/editor.spec.ts` (SE-UI-2, SE-UI-5).

## SE-CALC-4 — Quantities and calendar values

Numeric values may carry a seven-axis integer SI dimension vector and quantity-kind metadata. The standalone rational kernel converts curated engineering units exactly, checks incompatible dimensions, and preserves input-unit metadata when display units change. Unit parsing is bounded; unknown units return typed errors. The registry is deliberately separate from the full platform constants catalog. Calendar values remain rational serials; date/time display uses UTC calendar fields and preserves Excel day 60. TODAY/NOW capture one clock per pass, in `options.timeZone` (headless default UTC; the editor supplies the viewer zone). Host-defined name/call results are captured with exact parser-span keys and formula text in optional result context for immutable export.

Verification: `units.test.ts`, `catalog.test.ts`, `parity.test.ts`, HTTP/provider tests and Excel stored-context tests. The [Host comparison](HOST-PARITY.md) records intentional precision and error-token differences.
