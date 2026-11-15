---
'@flamboyantengineer/spreadsheet-editor': minor
---

Add the `./host` entry: a `SpreadsheetHost` contract (identity, permissions, a workbook store with revisions, versions, presence, drafts and lifecycle) shaped like the document editor's host adapters, `resolveSpreadsheetMode`, and `createMemoryHost` and `createMemoryPresenceRoom` for labs and tests. The editor does not use it yet.
