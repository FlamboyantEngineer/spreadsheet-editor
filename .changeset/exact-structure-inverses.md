---
'@flamboyantengineer/spreadsheet-editor': patch
---

Structure changes and their inverses no longer add empty `style`, `formats`, `inputValues`, `cellProvenance`, `rows`, `mergeCells` or `suggestions` entries a workbook did not have, so applying an operation's inverse restores the workbook exactly.
