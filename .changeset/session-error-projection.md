---
'@flamboyantengineer/spreadsheet-editor': patch
---

Retain terminal workbook-session failures through newer edits until an explicit save or flush retries them. Report background session failures through the editor handle and subscriptions consistently with the save status, while preserving causal receipt coverage and active drafts.
