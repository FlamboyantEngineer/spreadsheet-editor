# Document insertion

## SE-DOC-1 — Portable insertion payloads

The selected range can become a normal table or a live spreadsheet. Payloads carry `schemaVersion: 1`, `sheetId` and normalized A1 range.

- A table payload carries calculated text cells, merge spans, ordinary ProseMirror table JSON and escaped HTML. Formulas are replaced by their current displayed values. Empty cells remain valid paragraphs. It requires matching result dimensions; callers of the headless helper supply results for the current workbook. The React editor enforces current-result identity before enabling insertion.
- A live spreadsheet payload carries a deep copy of the entire source workbook, including raw formulas and referenced worksheets. It is independently editable. The selected range describes the intended document view; dependency data outside that selection is retained. The consumer owns its embedded node, React view, persistence and range presentation.

The ordinary JSON uses `table`, `table_row`, `table_cell`, `paragraph` and `text` nodes with colspan, rowspan and colwidth cell attributes. This matches the document editor's table schema. Hosts can assign their own document element identifiers during insertion.

Implementation: `src/document/insertion.ts`, `src/editor/spreadsheet-editor.tsx`.

Verification: `src/document/insertion.test.ts`.

## SE-DOC-2 — Usable embedding demonstration

Labs render the normal table as document content and the live payload as a second editable SpreadsheetEditor instance. Editing the embedded workbook must not mutate the original editor. Both instances can use client or server calculation independently.

Implementation: `labs/src/main.tsx`.

Verification: `e2e/editor.spec.ts` (SE-DOC-2).
