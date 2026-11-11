# Source provenance

The following modules were extracted at upstream commit `43f4f64773c9eca230aabc390ea7b03be3c290da`:

| Standalone path                         | Original path                                                                        |
| --------------------------------------- | ------------------------------------------------------------------------------------ |
| `src/engine/jspreadsheet-host.tsx`      | `apps/spreadsheets/client/spreadsheet-editor/components/JSpreadsheetHost.tsx`        |
| `src/engine/spreadsheet-cell-spill.ts`  | `apps/spreadsheets/client/spreadsheet-editor/components/spreadsheet-cell-spill.ts`   |
| `src/model/formula-parser.ts`           | `packages/platform-sdk/src/formula-parser.ts`                                        |
| `src/model/formula-reference-parser.ts` | `apps/spreadsheets/client/spreadsheet-editor/components/formula-reference-parser.ts` |
| `src/model/spreadsheet-addressing.ts`   | `packages/platform-sdk/src/spreadsheet-addressing.ts`                                |

The React grid host now loads no application stylesheets or fixed URLs. The standalone calculation orchestration and exact rational engine are new: they replace the application's fixed-decimal calculation boundary for this package.

Jspreadsheet CE, its React adapter, jSuites and decimal.js retain their respective upstream licenses distributed with those packages. This repository is private and UNLICENSED; extraction does not grant additional rights to its source.

The requirements comparison, enabled-function fixture and compatible standalone behavior were reviewed against the same source commit above. New rational catalog, units, calendar, clipboard, review, presence and Excel adapters implement those reusable contracts without importing application services. The reviewed function registry is `apps/shared/src/workbook-formula-functions.ts`; the extracted requirement source is `docs/product-specs/spreadsheet-editor/`. `specs/HOST-COMPATIBILITY.md` records application boundaries and numerical differences.

ExcelJS and fflate retain the upstream licenses distributed with those packages.

The core-port unit catalog and portable numeric behavior fixtures were adapted from commit `43f4f64773c9eca230aabc390ea7b03be3c290da`: `packages/platform-sdk/src/unit-constants.ts`, `numerical-core.ts`, `numerical-baseline.ts`, and `test/features/numerical/{quantity-engine,spreadsheet-numerical}.test.ts`. The rational implementation remains native to this repository. Numeric fixture provenance and individual intentional differences are recorded in `specs/NUMERIC-CORPUS.json` and `specs/NUMERIC-INTENTIONAL-DIFFERENCES.json`.

The generic captured Excel formula contract was reviewed against the same commit's `apps/spreadsheets/server/xlsx/xlsx-stored-formulas.ts`; host names are explicit capability descriptors rather than application-specific branches. HTML conversion uses parse5, whose upstream license is distributed with that package.

The external-provider lifecycle contract was reviewed against `FlamboyantEngineer/collaboration-tools` commit `8e3b096b0b8bab90858c4d4bdc56a2e7118989b0`, `src/client/room-provider.ts`. No transport implementation or workbook-specific protocol was added to that repository.
