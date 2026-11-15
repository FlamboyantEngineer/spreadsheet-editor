---
'@flamboyantengineer/spreadsheet-editor': minor
---

Remove the `./engine` entry point; the grid engine is internal to the editor. Remove the unused `setWorkbookCell`, `normalizeSpreadsheetMergeCells`, `isSpreadsheetFormulaParseable` and `buildFormulaTokenOverlayStyle` exports (use `writeCells` instead of `setWorkbookCell`). `exceljs`, `fflate` and `parse5` become optional peer dependencies, needed only by `./excel`.
