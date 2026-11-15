---
'@flamboyantengineer/spreadsheet-editor': minor
---

The editor shares presence through the host's `presence` adapter, addressed by a new `workbookId` prop, in place of the `collaborators` and `onPresenceChange` props. It publishes the reader's worksheet and selected range, the range they are writing in as an advisory lock and whether they are active or away, and withdraws when it unmounts. Peers carry document-editor-shaped records (`at`, `editLock`, `status`) and are placed on the workbook before they are shown. The `SpreadsheetCollaborator` and `SpreadsheetPresence` types leave the main entry, and the toolbar shows each peer's status in place of their activity. A new `workbookId` starts a new editor session.
