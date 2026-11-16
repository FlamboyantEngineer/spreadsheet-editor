---
'@flamboyantengineer/spreadsheet-editor': minor
---

Every area now meets the shared size limits (complexity, depth, parameters, function and file length), so functions that took more than three parameters take an options object instead:

- `./calculation`: `formulaOperandKey(cell: { sheetId, x, y }, span)`, `makeQuantity(value, unit, { quantityKind })` and `formatExactDecimal(text, locale, options)`. New types `ResolvedUnit` and `QuantityOptions`.
- `./document`: `updateLinkedSpreadsheetInsertion(current, update)`, where `update` is a `LinkedSpreadsheetUpdate`.
- `./model` editing: `writeCells` and `writeCellsWithOperations(workbook, values, { sheetId, x, y, idFactory? })`; `formatCells(workbook, { sheetId, range }, { properties, format? })`; `clearSelection`, `borderCells` and `mergeSelection(workbook, { sheetId, range }, flag)`; `createSpreadsheetCopy(workbook, { sheetId, range }, { text, cut? })`; `pasteSpreadsheetCopy(workbook, copy, { sheetId, x, y, destination?, idFactory? })`; `selectionLock(workbook, { sheetId, selection }, peers)`.
- `./model` structure: `changeWorkbookStructureWithOperations(workbook, { sheetId, idFactory? }, change)`, `rewriteStructureReferences(formula, change, { currentSheet, targetSheet })` and `sheetOperation({ kind, sheetId? }, before, after)`.
- `./model` review: `addComment(workbook, { sheetId, range, text }, stamp)`, `setCommentState(workbook, { id, state }, stamp)`, `replyToComment(workbook, { id, text }, stamp)`, `editCommentEntry(workbook, { threadId, postId }, { text, actorId, time })`, `deleteCommentEntry(workbook, { threadId, postId }, actorId)`, `toggleCommentReaction(workbook, { threadId, postId }, { emoji, actorId })`, `proposeWorkbookChange(workbook, { proposed, summary, operation? }, stamp)` and `decideSuggestion(workbook, { id, decision }, stamp)`. New types `NewComment`, `CommentEntryRef`, `WorkbookProposal` and `SuggestionDecision`.

`parseDelimited` and `serializeDelimited` are still exported from `./model`. The editor, the grid engine and every behaviour are unchanged.
