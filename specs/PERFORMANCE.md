# Performance

## SE-PERF-1 - Grid work scales with what changed

### SE-PERF-1.1 - Grid work scales with what changed

A worksheet of 10,000 cells stays smooth to type into and edit, and one of 50,000 cells stays usable. The grid's per-edit work is limited to the cells that changed:

- Typing in a cell rewrites only the draft cell and the cells whose value or readiness changed; unchanged cells keep their text, presentation and DOM nodes.
- A row whose values, inputs, formats and presentation are the same objects as at the last paint, and that holds no draft or open cell editor, is not visited, unless the engine has rewritten one of its cells. The editor gives a values row a new identity whenever a cell in it changes value or readiness. Calculated presentation and rich content are applied as rows enter the viewport. Review markers remain valid while their original engine row is detached.
- Committing an edit updates changed values in place. The grid rebuilds its rows only when its extent changes or more than 2,000 cells change at once.
- Zoom and authored styles are applied to every cell only when zoom changes; otherwise only to cells whose authored style changed and cells the engine has newly created. At 100% the stylesheet's own text metrics apply without inline copies.
- Shared-border reconciliation visits only cells with an inline border or a clipboard edge and their right and bottom neighbors. Unchanged column, row and frozen-column measures are not rewritten.
- The editor repairs frozen offsets only on attached rows during scrolling and column resize previews. Newly attached rows receive current offsets when they enter the viewport. A resize commit copies only the changed worksheet geometry and shares authored data.
- Native change capture compares cell data directly instead of serializing the worksheet.

- Only the rows near the viewport are attached to the grid body, with about 20 rows of overscan on each side. Spacer rows stand for the detached rows at their own heights, so the scrollbar spans the whole worksheet and attached rows sit where they would otherwise. Row offsets are reused while row geometry is unchanged. Rows attached later take the current zoom metrics, shared borders and text spill. Keyboard moves, find and scrolling a cell into view attach the target row before it is measured, a merge that crosses the window's edge is attached whole, and selection outlines are measured on the attached rows only. The engine builds every row; windowing is by rows, not columns. Implicit viewport padding respects the workbook cell limit.
- A change in dimensions can arrive while a calculation or protected edit is pending. Paint stays within the mounted grid, uses the current raw worksheet to determine which authored cells exist, and retains eligible values only where the current and retained outputs overlap. A protected edit keeps its mounted worksheet until replacement can be applied. Whole-row and whole-column reference outlines use the mounted grid extent.

Implementation: `src/engine/cell-renderer.ts`, `src/engine/row-window.ts`, `src/engine/worksheet-presentation.ts`, `src/engine/shared-grid-borders.ts`, `src/engine/worksheet-data.ts`, `src/engine/grid-geometry.ts`, `src/engine/header-resize.ts`, `src/engine/workbook-grid.tsx`, `labs/src/large-lab.tsx`.

Verification: `e2e/grid-budget.spec.ts` (mutation and time budgets on the large worksheet), `e2e/row-window.spec.ts` (scrolling, keyboard moves, zoom, editing, find and selection on 50,000 cells), `src/engine/row-window.test.ts`, `src/engine/grid-budget.test.ts`, `src/engine/cell-renderer.test.ts`, `src/engine/shared-grid-borders.test.ts`, `src/engine/worksheet-data.test.ts`, `src/engine/grid-geometry.test.ts`, `src/engine/header-resize.test.ts`.

## SE-PERF-2 - Work per edit and per keystroke

### SE-PERF-2.1 - Work per edit and per keystroke

A workbook of 10,000 cells stays smooth to edit and one of 50,000 cells stays usable, with client calculation on the main thread. The editor's own work for one committed edit, and for one formula-bar keystroke, does not grow with every cell of the workbook:

- A command names the cells it writes; its cell patches compare only those cells. A command that names none is compared by value in place, without copying or serializing unchanged cells.
- What the editor produced itself is not validated or copied again: a client calculation result is trusted, the host's frozen baseline is copied only when the host reads it, and a draft is attributed to its contributor once.
- What depends only on a calculation result (the reverse dependencies, the ready values) is computed once per result. A keystroke walks only the cells its draft cell reaches, and output rows without a pending cell keep their identity, while a row with one is always new.

While the client or server recalculates after an edit, the editor keeps showing the last ready value of every cell that does not read an edited cell, directly or through other formulas; only those readers show as pending. The calculation itself stays pending with no current result, so exports and document insertion still wait for it. A structural or naming change shows every cell as pending.

Replace all writes every match into one copy of the workbook, as one change and one undo step. Local undo keeps each step's operations, which carry the before and after state of every cell, geometry, merge, annotation and name they change, together with the row lengths and optional maps those cells need, and rebuilds the workbook from them on undo and redo. Only a step that adds or removes lines or worksheets, or replaces the workbook, keeps a snapshot of the workbook it returns to.

Calculation itself still recalculates the whole workbook; a calculation worker and incremental recalculation are not part of this requirement.

Implementation: `src/editor/use-editor.ts`, `src/editor/history.ts`, `src/editor/command-operations.ts`, `src/editor/use-publication.ts`, `src/editor/use-calculation.ts`, `src/model/operations.ts`, `src/calculation/publication.ts`.

Verification: `src/editor/editor-performance.test.ts` (SE-PERF-2.1, SE-8) keeps a Node budget for a 10,000-cell and a 50,000-cell workbook and for replace all; `src/editor/history.test.ts` (SE-9) undoes and redoes every kind of step back to the same workbook.
