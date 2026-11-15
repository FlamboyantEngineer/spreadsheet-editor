---
'@flamboyantengineer/spreadsheet-editor': patch
---

Local undo and redo keep the operations of each step instead of a copy of the whole workbook, so history on large workbooks uses far less memory. Only steps that add or remove lines or worksheets, or replace the workbook, still keep a snapshot.
