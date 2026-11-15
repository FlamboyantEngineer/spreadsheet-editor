---
'@flamboyantengineer/spreadsheet-editor': minor
---

Saving goes through a workbook session over the host's workbook store. `createWorkbookSession` (in `./host`) and `useWorkbookSession` load a workbook, record each accepted change by its transaction id, save against the revision it was edited from, autosave, retry an unavailable store while autosaving, stop on a conflict for the reader to keep their changes or take the stored workbook, and keep drafts across a reload through a `drafts` adapter. A save covers only the changes accepted before it started, and `acknowledge(transactionId)` marks changes saved by other means. The editor takes the session as its `session` prop; `onSave`, `onFlush`, `autosaveMs` and the `SpreadsheetSaveContext` type are removed.
