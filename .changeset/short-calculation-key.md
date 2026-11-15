---
'@flamboyantengineer/spreadsheet-editor': patch
---

Make `workbookCalculationKey` a short `wck1-` digest of the calculation inputs, computed once per workbook, instead of the full canonical JSON. Hosts compare and echo the key as before; its value changes.
