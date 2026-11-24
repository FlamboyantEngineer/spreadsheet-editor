---
'@flamboyantengineer/spreadsheet-editor': minor
---

Isolate replacement workbook sessions, serialize saves with taking the stored copy, and keep canonical inputs and drafts intact during no-op replacements. Find runs only while open. Frozen-column resize and scroll repairs visit the mounted rows.

Preserve clipped formula and annotation ranges through structural undo and reject worksheet deletion that would miss concurrent references. Line operations can include comparison frames for exact inversion. Publish the same cycle errors regardless of traversal order, require current canonical values for capture, and report Excel export fallbacks accurately.
