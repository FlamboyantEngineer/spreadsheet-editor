---
'@flamboyantengineer/spreadsheet-editor': minor
---

Saved versions come from a `versions` prop, the host's `VersionStoreAdapter`, with `workbookId`. The History pane lists the workbook's versions and follows the store; the editor opens a version read-only in place of the workbook and returns to the current one itself, and saves a named version, renames one and restores one where the store and the reader's `saveVersion` and `edit` capabilities allow. Removes the `history` and `historical` props and the `SpreadsheetHistory`, `SpreadsheetHistoryEntry` and `SpreadsheetHistoricalView` types.
