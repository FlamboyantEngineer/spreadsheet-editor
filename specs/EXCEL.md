# Optional Excel boundary

## SE-XLSX-1 — Import and loss reporting

The Node-only `/excel` entry exposes `importXlsx(bytes, { signal, timeoutMs })`. Conversion runs in a child with a 640 MiB heap, 768 MiB RSS watchdog, two-minute ceiling and 100 MiB output ceiling. ZIP metadata is checked before decompression: 20 MiB input, 128 MiB expansion, 8,192 entries, 32 MiB per part, bounded XML depth, canonical paths, ordinary workbook main type, no encryption/macros/DOCTYPE/entities. Packages are limited to 200 worksheets and 250,000 non-empty cells before ExcelJS parsing.

Conversion retains worksheet order/names, raw numeric XML strings, authored/expanded shared formulas, merges, widths/heights and frozen columns. Excel formula caches are discarded; conversion never calculates. Common fonts, solid fills, per-side borders, wrapping, alignment, theme/indexed colors and number/date/time/unit presets become authored formatting. Unsupported features produce grouped warnings with a disposition, count and at most five example locations. Hidden content becomes visible and frozen rows/print state are skipped. Rich text and hyperlinks become plain text; comments, conditional formatting, charts, tables and similar unsupported features are reported. Imported authored cells receive system provenance.

The host prepares calculations, presents the loss report and owns preview/commit. `onImportXlsx` returns the explicitly confirmed workbook or null. The labs demonstrate a read-only server-prepared preview with no client reevaluation; cancellation preserves the current workbook.

## SE-XLSX-2 — Captured-result export

`exportXlsx(workbook, result, { signal, timeoutMs })` uses a separate bounded child and allows at most two concurrent exports per Node process (a third rejects with status 503). It never calculates or reads live external services. It verifies matching worksheet/result shapes, writes ordinary formulas with existing caches, inlines recorded named/external operands by parser span, and emits stored values for units and EXPM1/LOG1P. Unsupported formulas retain expressions without caches. Malformed formulas retain authored text. Excel recalculation-on-open/save is disabled. Values become native numbers only after exact decimal round-trip validation; otherwise terminating exact decimals remain full text and non-terminating rational values remain fraction text. Styles, geometry, merges and frozen columns are retained; comments, suggestions and contributor metadata are omitted.

The host supplies authorized immutable stored results, route/access policy, attachment headers, source provenance and download orchestration. The browser callback is an integration seam, not a trusted storage or authorization boundary. See [HOST-PARITY.md](HOST-PARITY.md).

Verification: `src/excel/excel.test.ts`, `scripts/check-package.mjs` through both packaged child-process paths, and `e2e/flight-parity.spec.ts` through preview/cancel/replace/download.
