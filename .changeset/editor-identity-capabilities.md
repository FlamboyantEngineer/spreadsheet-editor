---
'@flamboyantengineer/spreadsheet-editor': minor
---

The editor takes the host's `identity` adapter and `EditorCapabilities` (`capabilities`) from the `./host` entry, in place of `currentUser`, `permissions` and `readOnly`, which are removed. `permissions.decideSuggestions` is now `capabilities.decide`; `readOnly` is `capabilities={READ_ONLY_CAPABILITIES}`. `SpreadsheetPermissions` is removed and `SpreadsheetMode` comes from `./host`.
