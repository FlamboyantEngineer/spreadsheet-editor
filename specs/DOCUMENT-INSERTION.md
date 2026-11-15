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

## SE-DOC-3 — Linked authoritative capture

The optional `/document` `createLinkedSpreadsheetInsertion` helper captures a version-one `linked-spreadsheet`: opaque `source: { id, revision }`, selected `sheetId`/normalized `range`, detached full `workbook`, validated ready `publication`, and ordinary captured `table`. A linked capture never runs a calculator. A stale, missing or non-ready authoritative publication rejects. Captured table JSON/HTML is the deterministic read-only, historical, export and print representation.

`updateLinkedSpreadsheetInsertion` explicitly refreshes the same source/selection using newly resolved host data. It requires `expectedRevision` to match the current captured revision and retains the source identity. Reusing the same source revision with different captured workbook/publication data rejects. Opaque source revisions have no package-defined ordering; the host resolver supplies authoritative current data and authorization. Historical consumers retain the captured table and do not automatically refresh. The existing independent `spreadsheet` insertion remains unchanged. Consumers can adapt the neutral capture to their own embedded document node without a circular package dependency.

Implementation: `src/document/linked.ts`.

Verification: `src/document/linked.test.ts` and packed `/document` imports.
