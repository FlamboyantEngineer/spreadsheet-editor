---
'@flamboyantengineer/spreadsheet-editor': patch
---

Fix an inline `onPanelChange` that stored its state re-rendering forever, row and column insertion on tall, narrow sheets, empty border style entries, sheet names outside ASCII in formulas, quadratic reference rewriting, and calculation transport limits that refused workbooks the cell limit accepts. A refused calculation request now says which limit it hit.
