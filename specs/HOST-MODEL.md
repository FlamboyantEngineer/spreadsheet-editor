# Host model

## SE-HOST-MODEL-1 - Canonical authored values and logical identities

### SE-HOST-MODEL-1.1 - Canonical authored values and logical identities

Version-1 worksheets optionally carry `inputValues`, `rowIds` and `columnIds`. Canonical inputs are validated rational/quantity values attached to existing nonformula authored cells. Their primitive `data` remains the edit representation, including text in a different input unit. Replacement and clearing discard the canonical override; formatting preserves it. JSON, internal clipboard, structure, review baselines, snapshots and live document insertion retain canonical values. Metadata addresses count once toward the cell budget.

Line identity lists, when supplied, contain unique nonempty bounded strings and cover the effective logical extent. `worksheetExtent(sheet, workbook?)` includes data, explicit minimum dimensions, geometry, styles, formats, canonical inputs, provenance, merges and available review anchors. Structural changes update the saved extent even when the affected final line is blank. Insertions receive explicit generated IDs; deletion removes only affected IDs and retains at least one line. `materializeWorksheetExtent` gives editor commands their initial 26-by-100 viewport without forcing every headless model to allocate a blank matrix; when that padding would take the sheet past the cell limit (a tall, narrow sheet such as 10,000 by 20), the command works on the logical extent instead. Hosts supply `WorkbookIdFactory` for deterministic entity generation.

Formatting, borders, merges, resizing, and comment anchors also extend line identities when they reach beyond the saved extent. An editor worksheet created from an identified worksheet begins with identities for those same axes.

Implementation: `src/model/workbook.ts`, `src/model/extent.ts`, `src/model/structure.ts`, `src/model/editing.ts`.

Verification: `src/model/host-model.test.ts` (SE-HOST-MODEL-1.1), `src/model/structure.test.ts`.

## SE-HOST-MODEL-2 - Semantic command records

### SE-HOST-MODEL-2.1 - Semantic command records

`WorkbookChangeContext` records a transaction ID, source, detached baseline and ordered `WorkbookOperation` list. Operations explicitly identify cells, lines, sheet addition/deletion/rename/order, geometry, merges, annotations and deliberate import/new/reset replacement. A line operation comes from its structural command, including when two blank matrices are identical; it is never inferred from cell differences. Insert/delete records carry line IDs and surviving neighbors. Cell records carry authored value, canonical input, style, format and contributor provenance, with stable line IDs when available. Rewritten cross-sheet references and annotation anchors belong to the same command result. Removed authored cells and sparse geometry remain available to `invertWorkbookOperations` for standalone undo.

`collectCellOperations` visits authored or metadata coordinates, or a caller-supplied bounded target set. It does not construct an implicit grid of blank cells. `changeWorkbookStructureWithOperations` and `writeCellsWithOperations` return the workbook and its command records together. The editor owns publication and rollback, and the host owns collaboration preflight and durable storage.

Growth from an ordinary edit records extent-only identity changes without moving formulas or setting minimum dimensions. Explicit structure commands carry any viewport and identity preparation needed to replay them. Sparse cell changes preserve blank slots, row lengths, and optional map presence in the same coordinate frame. Legacy suggestions with recorded structure commands keep their line operations on acceptance. Transaction preflight recognizes leading whitespace before formulas and compares merge coverage on deleted lines. Contiguous line deletions replay as one span.

Implementation: `src/model/operations.ts`.

Verification: `src/model/host-model.test.ts` and `src/model/operation-roundtrip.test.ts` (SE-HOST-MODEL-2.1).

## SE-HOST-MODEL-3 - Authored clipboard blocks

### SE-HOST-MODEL-3.1 - Authored clipboard blocks

Internal copy snapshots primitive expressions and supported style, format, canonical input and contributor provenance. A copied block tiles a destination only when both destination dimensions are exact multiples of the source. Smaller and nonmultiple destinations receive one complete block from their top-left. Every copied formula translates against its original source coordinates, preserving absolute axes. A cut moves its block once, clears the source atomically and rejects a stale source before mutation. Source row and column identities must match when both are complete. Otherwise, the whole source worksheet must still match the snapshot. Foreign delimited text is bounded by transport, cell count, column count and field length limits before workbook mutation.

Implementation: `src/model/clipboard.ts`, `src/model/editing.ts`.

Verification: `src/model/host-model.test.ts` (SE-HOST-MODEL-3.1, SE-HOST-MODEL-2.1), `src/model/clipboard.test.ts`.
