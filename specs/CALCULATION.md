# Calculation contract

## SE-CALC-1 - Exact values

### SE-CALC-1.1 - Exact values

Numeric literals, including exponent notation, parse from strings into reduced rational values. Exact results serialize as `{ kind: 'number', numerator: string, denominator: string, approximate: false }`, with a positive denominator. Arithmetic and comparisons never convert these values through binary floating point. `1E-25`, `1-1E-25`, `1-(1-1E-25)`, `1/3*3` and cancellation at `1E-250` remain exact. JavaScript numeric input is accepted for compatibility but cannot restore digits already lost by its caller.

Display rounds a copy to configurable significant digits (32 by default), with decimal, automatic or fraction notation. Explicit rounding functions change the value using their documented direction; ROUND uses halves away from zero. No display option changes subsequent calculations.

Non-rational operations use a per-calculation decimal context (default 80 significant digits, supported range 16–1000). They serialize an approximate rational decimal with `approximate: true` and `precision`. Dependent arithmetic retains the approximation marker. The display prefixes these results with `≈` unless disabled. Perfect rational square roots and integer powers remain exact.

Implementation: `src/calculation/rational.ts`, `src/calculation/values.ts`, `src/calculation/format.ts`.

Verification: `src/calculation/evaluator.test.ts`, `e2e/editor.spec.ts` (SE-UI-1.1, SE-CALC-1.1).

## SE-CALC-2 - Workbook evaluation

### SE-CALC-2.1 - Workbook evaluation

Formulas parse into a syntax tree; no authored formula reaches JavaScript evaluation or code generation. Cross-sheet references are case-insensitive, unquoted names may use letters and digits of any script, quoted names and absolute addresses are supported, and whole-row/column ranges are bounded by calculation extents (authored data, canonical inputs, identities, explicit geometry, merges and unit formats). Visual styles, comments, contributor markers and display-only formats can extend the viewport without changing calculation ranges. A bounded reference tokenizer also serves draft highlighting and reference rewrites, preserving worksheet qualifiers, range spacing, absolute axes and host bindings. Error tokens become typed errors and quoted strings retain their literal text. Cells are memoized within one calculation. Cycles, invalid references, numerical errors, unknown functions and work-limit failures produce typed error values. IF and IFERROR evaluate only the selected branch. Text compares case-insensitively with one fixed collation (`en-US`), never the runtime's locale, so client and server calculation agree. Hosts may resolve names and synchronous custom functions explicitly.

The exported `SUPPORTED_FUNCTIONS` array is sorted by name and derived from executable functions and aliases. The initial families below are retained; the executable catalog now also covers all 154 supported worksheet names and aliases from the pinned registry in [FUNCTION-CATALOG.json](FUNCTION-CATALOG.json). Additions include conditional aggregates, lookups, calendar/text functions, hyperbolic math and stable EXPM1/LOG1P. Pending registry entries remain unsupported. Families:

- Logic: IF, IFERROR, AND, OR, NOT, TRUE, FALSE, ISERROR, NA.
- Aggregates: SUM, PRODUCT, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK, SUMSQ.
- Arithmetic: ABS, SIGN, SQRT, POWER, MOD, ROUND, ROUNDUP, ROUNDDOWN, INT, TRUNC, CEILING, FLOOR, QUOTIENT.
- Approximate math: PI, SIN, COS, TAN, ASIN, ACOS, ATAN, EXP, LN, LOG, LOG10.
- Statistics: VAR.S, VAR.P, STDEV.S, STDEV.P.
- References: ROW, COLUMN, ROWS, COLUMNS.
- Text and type checks: LEN, LEFT, RIGHT, MID, LOWER, UPPER, TRIM, CONCAT, CONCATENATE, VALUE, ISNUMBER, ISTEXT, ISBLANK.

Default bounds are 250,000 populated cells, 200 worksheets, 10,000 characters per input cell, 10,000 digits per numerator/denominator, 2,000,000 evaluation operations and expression depth 128. The parser bounds nested syntax at 256 levels. Dependency chains use an explicit resumable stack; each formula is separately limited to 10,000 cumulative addressed cells before range allocation. Numeric size and evaluation limits bound resource usage rather than introducing rounding. The synchronous API permits explicit limits within the arithmetic engine's maximums. Over-budget cell computations return `#LIMIT!`. A malformed or oversized workbook is rejected before evaluation.

Implementation: `src/model/formula-parser.ts`, `src/model/formula-reference-token.ts`, `src/calculation/evaluator.ts`, `src/model/workbook.ts`.

Verification: `src/calculation/evaluator.test.ts`, `src/model/formula-references.test.ts`, `src/model/workbook.test.ts`.

## SE-CALC-3 - Server transport and parity

### SE-CALC-3.1 - Server transport and parity

`CalculationProvider.calculate(workbook, options, signal)` returns a versioned result. The HTTP provider posts a versioned request with a per-provider request id, and validates the response id, workbook dimensions, error tokens and rational fields. Exact numerator/denominator strings, temperature tags, provenance, captured host JSON and dependency keys survive transport unchanged. Response validation checks every numeric value, captured operand and dependency tuple. Cell keys use canonical JSON array serialization. Host JSON contains finite primitives, arrays or plain objects, with a depth limit of 32 and transport byte limits.

`handleCalculationRequest` accepts Fetch-standard Requests and returns Responses. It accepts POST JSON only, bounds streaming request bodies, validates inputs and caps requested limits at server ceilings. The transport has independent byte limits. The default request limit is 64 bytes per cell of the server's `maxCells` and the default response limit is 256 bytes per cell of the client's requested `maxCells`. Each default leaves at least 1 KiB for framing small workbooks. Hosts may set explicit limits for larger serialized values and metadata. A refused request gets `413` with `limit` set to `requestBytes` (the body was too large) or `workbook` (a workbook or calculation limit), and a `detail` message. Both sides validate cells against one list of error codes. Authentication and routing are supplied by the host. No server state or global arithmetic context is shared between requests.

Implementation: `src/calculation/provider.ts`, `src/calculation/http.ts`.

Verification: `src/calculation/http.test.ts`, `e2e/editor.spec.ts` (SE-UI-2.1).

## SE-CALC-4 - Quantities and calendar values

### SE-CALC-4.1 - Quantities and calendar values

Numeric values may carry the unchanged seven-axis SI dimension vector, physical `quantityKind`, `temperature: 'absolute' | 'difference'`, and opaque namespace/revision `provenance`. Angle and information remain distinct even with zero SI vectors; their derived powers remain explicit kind metadata. Scalar operations preserve physical kind. Incompatible explicit kinds (including torque/energy) or conflicting revisions in one provenance namespace return `#VALUE!`. Provenance unions across arithmetic, aggregates, rounding and host-derived dependencies.

The complete pinned unit catalog and decimal/binary prefixes are portable source data in `unit-catalog.ts`. Factors, divisors and affine offsets are strings. Pi-based angle factors use the isolated approximation context. Aliases are case sensitive: literal `C`/`F` mean coulomb/farad; Excel CONVERT has separate temperature-code aliases. Unit-configured cells can interpret C/F as temperatures only when their expected dimension explicitly says Temperature. Expected dimensions accept case-insensitive category names or unit expressions. Word units require the positive bounded integer-string `options.wordSizeBits` (1–100000); no architecture size is assumed.

Standalone pure-temperature literals, including kelvin and authored signed affine literals, are absolute. Absolute minus absolute yields a difference; absolute plus difference and difference plus absolute yield absolute; differences add/subtract as differences. Absolute plus absolute and difference minus absolute are invalid. Calculated absolute temperatures cannot be negated; authored signed temperature literals apply their sign before conversion. AVERAGE/comparison reject mixed absolutes and differences. Averaging absolute values preserves their kind, without illegally summing absolutes. Compound temperatures use scale without affine offset. Difference display converts using scale only. Old version-1 untagged pure temperatures mean absolute at ingress; new outputs carry explicit tags.

Nonformula `Worksheet.inputValues` supply validated authored canonical values before primitive parsing. They bypass `inputUnit` inheritance, preserving stored SI values while the formula bar retains the original edit text. Applying a display format never mutates a canonical input. An authored apostrophe escape preserves literal text under unit formatting. A quantity literal can be followed by an arithmetic operator and a cell reference or function call.

Scalar function arguments (text counts, choice indices, rounding places, calendar fields and serials, and CONVERT magnitudes) reject SI quantities and zero-axis angle/information kinds instead of discarding their units. ATAN2 accepts compatible coordinate dimensions/kinds and retains their unioned provenance; incompatible coordinates return `#VALUE!`. Unit formatting validates normalized dimensions and non-SI kinds, so authored/canonical radian and information values retain their meaning.

Selection statistics reuse the same exact quantity kernel. Count includes nonempty values; sum and average are independently unavailable when their physical-kind, revision or temperature rules prohibit the operation. Absolute temperatures can have a valid average even when their sum is undefined. Mixed dimensions never produce a unitless statistic.

Calendar values remain rational serials. DATE resolves month and day overflow in Excel serial space, including fictitious 1900-02-29; EDATE/EOMONTH preserve day 60 and valid boundaries from serial zero through December 9999. Internal calendar fields never reapply DATE's short-year convention, and out-of-domain month shifts return `#NUM!`. Date presentation formats calendar fields in UTC without viewer-zone shifts. TODAY/NOW capture one clock per pass in `options.timeZone` (headless UTC). WEEKDAY supports modes 1–3 and 11–17.

Implementation: `src/calculation/values.ts`, `src/calculation/units.ts`, `src/calculation/unit-catalog.ts`, `src/calculation/calendar.ts`, `src/calculation/evaluator.ts`, `src/calculation/http.ts`.

Verification: `src/calculation/quantity-metadata.test.ts`, `src/calculation/units.test.ts`, `src/calculation/dimensional-arguments.test.ts`, `src/calculation/calendar-boundaries.test.ts`, `src/calculation/summary.test.ts`, `src/calculation/core-numeric-corpus.test.ts`, HTTP/provider tests and Excel stored-context tests.

## SE-CALC-5 - Portable function semantics and dependencies

### SE-CALC-5.1 - Portable function semantics and dependencies

[NUMERIC-CORPUS.json](NUMERIC-CORPUS.json) records portable formula/quantity cases and their source revision. It supplements catalog membership with semantic aggregate, lookup, text, calendar, quantity and INDEX-shape fixtures. [NUMERIC-INTENTIONAL-DIFFERENCES.json](NUMERIC-INTENTIONAL-DIFFERENCES.json) records individual rational/default-precision/display/limit differences; it grants no family-wide exceptions.

TEXT accepts a bounded explicit set of numeric masks (`0`, `#,##0`, optional 1–100 zero decimals, currency prefix or percent suffix) and calendar masks documented in `text-format.ts`. Unsupported masks and quantities return `#VALUE!`. Windows-1252 CHAR/CODE remain separate from Unicode UNICHAR/UNICODE. INDEX truncates fractional indices, selects either vector axis, and preserves whole row/column/matrix results for zero axes and nested functions. Exact text lookup and wildcard criteria match text cells. Lookup return cells that are blank produce zero, while invalid coordinates keep their typed reference error. Exact integer LOG results are verified with rational powers before returning, so INT(LOG(8,2)) remains 3.

`CalculationResult.dependencies` records actually-read canonical JSON `[sheetId,x,y]` cell tuples and used-extent worksheet IDs. Invalid grid coordinates produce typed errors with no out-of-grid dependency keys. The native calculator records an entry for every formula cell and none for value cells, and marks this with `scope: 'formulas'`; without that mark a cell with no entry has unknown dependencies and a draft treats it as pending. A formula that reads a value cell does not suspend; reading a formula cell that is not calculated yet suspends the reader without capturing a stack, so evaluation cost does not depend on whether a range sits above or below its data. Each worksheet's calculation extent is read once per calculation. Lazy untaken branches add no dependencies. Collection shares addressed-cell budgets and resumable evaluation. A resumed range charges only its unread cells against the remaining operation budget. The evaluator supports 1,200-cell chains within the existing limits. FormulaHost name/call callbacks receive the originating formula and cell context. Optional bounded `hostContext` is captured separately from native values in result context and span-keyed operands, for immutable later export.

Implementation: `src/calculation/evaluator.ts`, `src/calculation/extended-functions.ts`, `src/calculation/text-format.ts`, `src/calculation/rational.ts`.

Verification: `src/calculation/core-numeric-corpus.test.ts`, `src/calculation/quantity-metadata.test.ts`, `src/calculation/parity.test.ts`, `src/calculation/catalog.test.ts`, `src/calculation/evaluator.test.ts` and `src/calculation/publication.test.ts` (SE-4).

## SE-CALC-6 - Exact locale and currency formatting

### SE-CALC-6.1 - Exact locale and currency formatting

Optional `CellFormat.locale` and `.currency` default to en-US/USD. Unsupported locales or currency codes fail workbook ingress with useful errors. Existing default decimal counts remain unchanged. Currency uses Intl currency codes and locale placement/grouping; percentages and serial calendars use the same requested locale. Fixed/scientific numeric formatting never converts authored values through binary Number arithmetic. Exact decimal-string Intl formatting is used on supporting runtimes, with a tested BigInt/string/format-parts fallback. The fictitious leap date is localized using calendar fields, not converted through a viewer time zone.

Implementation: `src/calculation/cell-format.ts`, `src/calculation/calendar.ts`, `src/model/workbook.ts`.

Verification: `src/calculation/cell-format.test.ts`, `src/calculation/quantity-metadata.test.ts` (EUR/de-DE, USD/en-US, integers beyond 2^53, negative currency, percent, leap day, forced fallback).
