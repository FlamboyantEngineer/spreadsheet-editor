---
'@flamboyantengineer/spreadsheet-editor': minor
---

Comment threads in the workbook take the document editor's record shape. A thread has `state` (`'open'` or `'resolved'`, with `resolvedBy` and `resolvedAt`) and `entries`, each written by a person id with `createdAt`, `editedAt` and reactions; a suggestion's thread names it in `suggestion` and records its `decision`. The editor names people through the `identity` adapter. The `CommentPost` type becomes `CommentEntry` and `CommentDecision` is new; in `./model`, `editCommentPost` and `deleteCommentPost` become `editCommentEntry` and `deleteCommentEntry`, and `setCommentState` resolves and reopens a thread. Workbooks with threads in the earlier `posts`/`resolved` shape are rejected.
