---
'@flamboyantengineer/spreadsheet-editor': patch
---

Replace all writes every match in one pass, so replacing 1,000 matches in a 50,000-cell workbook takes one change and one undo step instead of seconds of repeated copies.
