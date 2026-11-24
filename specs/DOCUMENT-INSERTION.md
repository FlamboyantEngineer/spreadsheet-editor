# Document insertion

## SE-DOC-1 - Portable insertion payloads

### SE-DOC-1.1 - Portable insertion payloads

The selected range can become a normal table or a live spreadsheet. Payloads carry `schemaVersion: 1`, `sheetId` and normalized A1 range.

- A table payload carries rows of calculated text cells with their merge spans (`rows`) and escaped HTML (`html`). Formulas are replaced by their current displayed values. Empty cells remain valid paragraphs. It requires matching result dimensions. Callers of the headless helper supply results for the current workbook. The React editor checks that the canonical publication has a current result for every value the table uses, including merged anchors outside the selected range.
- A live spreadsheet payload carries a deep copy of the entire source workbook, including raw formulas and referenced worksheets. It is independently editable. The selected range describes the intended document view; dependency data outside that selection is retained. The consumer owns its embedded node, React view, persistence and range presentation.

Each output row lists the cells that start in it as `{ text, colspan, rowspan }`. Cells covered by a merge are omitted. A merge clipped by the selected range uses the intersection as its span and displays the value and format of the original merge anchor. An explicit range may address any cell within the grid and is limited to 50,000 cells. The default range covers the worksheet's authored data rows and columns.

The document editor turns the rows into its stored table content with `tableFromRows` from its `./document` entry and gives the table its element identifiers as it inserts it. Other consumers can read the rows or the HTML. A live spreadsheet payload maps onto a document embed. The host keeps the workbook, names it and the range in the embed's `source` and its revision in `snapshot`, and its embed source answers with the range's calculated rows through `tableFromRows`.

Implementation: `src/document/insertion.ts`, `src/document/table-capture.ts`, `src/editor/editor-host-commands.ts`.

Verification: `src/document/insertion.test.ts`, `src/editor/document-insertion.test.ts`.

## SE-DOC-2 - Usable embedding demonstration

### SE-DOC-2.1 - Usable embedding demonstration

Labs render the normal table as document content and the live payload as a second editable SpreadsheetEditor instance. Editing the embedded workbook must not mutate the original editor. Both instances can use client or server calculation independently.

Implementation: `labs/src/main.tsx`.

Verification: `e2e/editor.spec.ts` (SE-DOC-2.1).

## SE-DOC-3 - Linked authoritative capture

### SE-DOC-3.1 - Linked authoritative capture

The optional `/document` `createLinkedSpreadsheetInsertion` helper captures a version-one `linked-spreadsheet`: opaque `source: { id, revision }`, selected `sheetId`/normalized `range`, detached full `workbook`, validated ready `publication`, and ordinary captured `table`. A linked capture never runs a calculator. A stale, missing or non-ready authoritative publication rejects. Captured table JSON/HTML is the deterministic read-only, historical, export and print representation.

`updateLinkedSpreadsheetInsertion(current, update)` explicitly refreshes the same source/selection using newly resolved host data: `update` carries the workbook, publication, source, `expectedRevision` and optional display options. It requires `expectedRevision` to match the current captured revision and retains the source identity. Reusing the same source revision with different captured workbook/publication data rejects. Opaque source revisions have no package-defined ordering; the host resolver supplies authoritative current data and authorization. Historical consumers retain the captured table and do not automatically refresh. The existing independent `spreadsheet` insertion remains unchanged. Consumers can adapt the neutral capture to their own embedded document node without a circular package dependency.

Implementation: `src/document/linked.ts`.

Verification: `src/document/linked.test.ts` and the packed document entry imports checked by the package check.
