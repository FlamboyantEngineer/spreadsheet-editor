# Source provenance

The following modules were extracted from upstream `43f4f64773c9eca230aabc390ea7b03be3c290da`:

| Standalone path                         | Original path                                                                        |
| --------------------------------------- | ------------------------------------------------------------------------------------ |
| `src/engine/jspreadsheet-host.tsx`      | `apps/spreadsheets/client/spreadsheet-editor/components/JSpreadsheetHost.tsx`        |
| `src/engine/spreadsheet-cell-spill.ts`  | `apps/spreadsheets/client/spreadsheet-editor/components/spreadsheet-cell-spill.ts`   |
| `src/model/formula-parser.ts`           | `packages/platform-sdk/src/formula-parser.ts`                                        |
| `src/model/formula-reference-parser.ts` | `apps/spreadsheets/client/spreadsheet-editor/components/formula-reference-parser.ts` |
| `src/model/spreadsheet-addressing.ts`   | `packages/platform-sdk/src/spreadsheet-addressing.ts`                                |

The React grid host now loads no application stylesheets or fixed URLs. The standalone calculation orchestration and exact rational engine are new: they replace the application's fixed-decimal calculation boundary for this package. The original application remains intact pending consumer migration.

Jspreadsheet CE, its React adapter, jSuites and decimal.js retain their respective upstream licenses distributed with those packages. This repository is private and UNLICENSED; extraction does not grant additional rights to its source.
