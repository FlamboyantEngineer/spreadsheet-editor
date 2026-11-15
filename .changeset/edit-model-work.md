---
'@flamboyantengineer/spreadsheet-editor': patch
---

Editing a large workbook no longer does whole-workbook work per edit and per keystroke: cell patches compare only what a command wrote, the editor's own calculation results are not re-validated, the change context's baseline is copied only when read, and a draft walks only the cells it reaches. At 50,000 cells a commit takes about a quarter of the time and a formula-bar keystroke about 4 ms.
