---
'@flamboyantengineer/spreadsheet-editor': patch
---

Split the editor hook's internals by concern (cell display, find, comments, versions, saving, notices, history steps, worksheet commands and the imperative handle). The API and behaviour are unchanged.
