---
'@flamboyantengineer/spreadsheet-editor': minor
---

Preserve workbook edits across host acceptance, saves, version changes and undo. A controlled editor paired with a workbook session now uses `acceptChange` to return the exact workbook the session should save. Custom sessions expose a stable `getSnapshot` for subscribers.

Scope grid actions to the active editor, verify the origin and target of clipboard moves, and bound grid rendering and navigation to safe coordinates. Keep formulas, exact values and cell types consistent across calculation, XLSX conversion and document insertion.
