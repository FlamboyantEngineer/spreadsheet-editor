# Optional Excel boundary

## SE-XLSX-1 — Import and loss reporting

The Node-only `/excel` entry needs the optional peer dependencies `exceljs`, `fflate` and `parse5`; no other entry imports them. It exposes `importXlsx(bytes, { signal, timeoutMs })`. Conversion runs in a child with a 640 MiB heap, 768 MiB RSS watchdog, two-minute ceiling and 100 MiB output ceiling. ZIP metadata is checked before decompression: 20 MiB input, 128 MiB expansion, 8,192 entries, 32 MiB per part, bounded XML depth, canonical paths, ordinary workbook main type, no encryption/macros/DOCTYPE/entities. Packages are limited to 200 worksheets and 250,000 non-empty cells before ExcelJS parsing.

Conversion retains worksheet order/names, raw numeric XML strings, authored/expanded shared formulas, merges, widths/heights and frozen columns. Excel formula caches are discarded; conversion never calculates. Common fonts, solid fills, per-side borders, wrapping, alignment, theme/indexed colors and number/date/time/unit presets become authored formatting. Imported styles use the editor's one style spelling (`key: value` pairs joined by `; `, SE-CHROME-1), so an imported cell and the same formatting applied in the editor store identical text. Unsupported features produce grouped warnings with a disposition, count and at most five example locations. Hidden content becomes visible and frozen rows/print state are skipped. Rich text and hyperlinks become plain text; comments, conditional formatting, charts, tables and similar unsupported features are reported. Imported authored cells receive system provenance.

The host prepares calculations, presents the loss report and owns preview/commit. `onImportXlsx` returns the explicitly confirmed workbook or null. The labs demonstrate a read-only server-prepared preview with no client reevaluation; cancellation preserves the current workbook.

## SE-XLSX-2 — Captured-result export

`exportXlsx(workbook, result, { signal, timeoutMs })` uses a separate bounded child and allows at most two concurrent exports per Node process (a third rejects with status 503). It never calculates or reads live external services. It verifies matching worksheet/result shapes, writes ordinary formulas with existing caches, inlines recorded named/external operands by parser span, and emits stored values for units and EXPM1/LOG1P. Unsupported formulas retain expressions without caches. Malformed formulas retain authored text. Excel recalculation-on-open/save is disabled. Values become native numbers only after exact decimal round-trip validation; otherwise terminating exact decimals remain full text and non-terminating rational values remain fraction text. Styles, geometry, merges and frozen columns are retained; comments, suggestions and contributor metadata are omitted.

The host supplies authorized immutable stored results, route/access policy, attachment headers, source provenance and download orchestration. The browser callback is an integration seam, not a trusted storage or authorization boundary. See [HOST-COMPATIBILITY.md](HOST-COMPATIBILITY.md).

Verification: `src/excel/excel.test.ts`, `scripts/check-package.mjs` through both packaged child-process paths, and `e2e/interaction-parity.spec.ts` through preview/cancel/replace/download.

The labs worker resolves ExcelJS as an installed Node dependency instead of bundling its optional transport integrations. Local XLSX conversion requires no cloud SDK or application service.

## SE-XLSX-3 — Host formula capabilities and bounded captured text

Import/export accept `options.formulaCapabilities: ExcelFormulaCapabilities`, containing a serializable `hostFunctions` list of `{ name, aliases?, capturedText? }`. The default list is empty; `capturedText` defaults to `plain`. Hosts explicitly declare their external function catalog, including any application-specific names. Declarations cannot override built-ins. Names and aliases normalize to uppercase, are unique ASCII formula identifiers of at most 128 characters, and are bounded to 128 functions and 16 aliases per function. Unknown properties, callbacks and unsupported policies reject before spawning and again inside the child.

Captured external calls use their exact parser span and matching authored formula, including repeated calls and dynamic/nested arguments. Declared aliases follow the same captured-operand path. Names retain their existing captured behavior. Conversion never resolves a missing operand or invokes a host function. Undeclared/unsupported functions remain formulas without caches and receive an explicit warning; unavailable/unrepresentable captured operands use stored-result fallback with a warning.

For a direct declared `html-to-text` host call, its frozen cached result receives the same plain-text projection as its inlined operand. When the plain result exceeds the 255-character formula-literal bound, the exporter retains it as ordinary stored cell text up to 32,767 characters and reports the formula loss. Derived formula results remain the supplied captured results; export does not evaluate an expanded formula.

`html-to-text` uses parse5's HTML parser and entity decoder. Script/style, embedded/interactive and other nontext subtrees are omitted, BR becomes newline, and formula quotes are escaped after extraction. Parsing is bounded to 65,536 input characters, 8,192 tree nodes and depth 128. Formula string literals remain limited to 255 characters, complete formulas to 8,192, and cell strings to 32,767. XML control characters and unpaired surrogates reject literal representation. A representability failure uses the existing captured stored-value fallback and an explicit loss outcome; no markup executes.

The child receives capability data inside the bounded request. Imports use a bounded length-prefixed JSON header followed by the original binary bytes, preserving the 20 MiB XLSX limit without base64 expansion. Export requests remain bounded JSON. Process cancellation, heap/RSS/time/output limits, two-export concurrency and discarded import caches remain unchanged.

Verification: `src/excel/excel.test.ts` (generic FOO, aliases, repeated/dynamic/nested calls, explicitly supplied example descriptors, default absence, HTML/entities/BR/unsafe markup, projected caches and long stored-text fallback, missing context, limits and request validation).

## SE-XLSX-4 — Exact projection, format metadata and conversion reports

`exportXlsx` preserves its byte-returning API. Additive `exportXlsxWithReport` returns `{ bytes, report }` with grouped warnings, bounded examples, and explicit `stored-value`/`not-exported` dispositions. It never computes or performs live lookups.

Currency formats use ISO currency codes and Excel locale-ID tags; recognized locale/currency metadata round trips through import. The supported Excel locale IDs include en-US, de-DE, en-GB, fr-FR, es-ES, it-IT, ja-JP, ko-KR, zh-CN, zh-TW, pt-BR, nl-NL, sv-SE, da-DK, fi-FI and nb-NO. Other valid locales export with the default Excel locale and an explicit loss warning. Calendar serials, including day 60, remain unchanged. Raw decimal XML and text-backed large integers preserve exact digits.

Canonical numeric overrides export their accepted stored values. Units and temperature meaning export as displayed stored text rather than unitless scalars. XLSX does not retain typed canonical, temperature/kind, input-unit, expected-dimension or provenance metadata; the report identifies those losses. Exact numbers not safely representable by an Excel number remain exact decimal or fraction text, never a rounded substitute.

Verification: `src/excel/excel.test.ts` (EUR/de-DE and USD/en-US, percentages, negative values, large integers, leap-day serials and unsupported-locale warnings).
