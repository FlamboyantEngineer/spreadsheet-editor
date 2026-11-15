---
'@flamboyantengineer/spreadsheet-editor': minor
---

Speed up whole-workbook calculation: forward reads resume without thrown errors, sheet extents are measured once, and calculation dependencies are recorded for formula cells only (`scope: 'formulas'`). A 50,000-cell edit now recalculates about three times faster.
