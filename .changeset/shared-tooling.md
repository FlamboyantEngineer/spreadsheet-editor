---
'@flamboyantengineer/spreadsheet-editor': patch
---

Checked with the shared lint rules and the reference, structure and export checks of the base component library's tooling. A formula that reads another formula whose value is empty no longer shows `#ERROR!`. The package root no longer carries a `'use client'` directive; each module that uses React hooks carries its own.
