# Editor chrome

## SE-CHROME-1 — Full reusable shell

The editor shall provide File, Edit, Insert, Format, View and Review menus; a scrolling formatting toolbar; an address/formula bar with function completions; worksheet tabs; a find/replace bar; docked Properties, History, Comments and Suggestions panes; accessible selection count/sum/average; and worksheet-strip zoom controls. Generic controls use `@flamboyantengineer/ui-components`. The host styles documented hooks; `--se-zoom` is a functional layout property rather than a package stylesheet.

Formatting shall support bold, italic, strikethrough, font family/size, text/fill colors, numeric formats and decimal places, horizontal/vertical alignment, borders, wrapping and merge/unmerge. Numeric formats round only display text, preserving exact rational calculations. Undo/redo includes formatting and review state. A frozen-column setting and display toggles expose grid integration without adding application services.

Authoring writes canonical data; Reviewing stages cell/format proposals and privately previews the local reviewer’s pending values; Reading refuses mutation. `readOnly` forces Reading, and host permissions independently gate editing, commenting, proposing and deciding suggestions in controls, commands and native engine callbacks. Optional host callbacks expose Save and Refresh; extra host panes extend the built-in panel set. Save errors are visible and edits made during a save remain dirty.

Implementation: `src/editor/contracts.ts`, `src/editor/editor-toolbar.tsx`, `src/editor/editor-menus.ts`, `src/editor/spreadsheet-editor.tsx`, `src/editor/use-editor.ts`, `src/calculation/cell-format.ts`.

Verification: `e2e/full-editor.spec.ts` (SE-CHROME-1, SE-REVIEW-5), `src/calculation/cell-format.test.ts`, existing browser editor tests.

## SE-CHROME-2 — Worksheet and structural commands

Worksheet actions shall add, duplicate, rename, reorder and delete sheets, retaining at least one. Row/column commands shall insert before/after and delete selections, move authored formatting and merge spans, rewrite qualified/unqualified and absolute formula references across worksheets, and retain quoted literals and scientific numeric text unchanged. References to deleted cells/sheets become `#REF!`; partially deleted ranges shrink. Comment anchors move with cells; wholly deleted anchors become orphaned. Affected suggestions become outdated instead of applying to another cell. Imports accept CSV or the versioned workbook JSON; exports expose the complete workbook JSON.

Implementation: `src/model/structure.ts`, `src/model/editing.ts`, `src/editor/editor-menus.ts`.

Verification: `src/model/structure.test.ts`, `e2e/full-editor.spec.ts` (SE-CONTEXT-1, SE-CHROME-2).

## SE-CHROME-3 — Search and keyboard ownership

Find/replace shall search raw values/formulas across worksheets, navigate to matches and replace one/all through the normal author/review mutation path. Case-insensitive literal queries must not execute as regular expressions. Typed search/comment/dialog content must never reach the vendor grid's document keyboard handler. Each instance owns its selection, menus and panels. Clipboard text preserves numeric strings and quoting; a host may supply clipboard adapters.

Implementation: `src/editor/use-editor.ts`, `src/engine/workbook-grid.tsx`, `src/model/editing.ts`.

Verification: `e2e/full-editor.spec.ts` (SE-CHROME-3, SE-REVIEW-6), `src/model/structure.test.ts`.

## SE-CONTEXT-1 — Shared React context menu

Secondary click, Shift+F10 and the Context Menu key shall open the shared React context menu for the selected cells. Clicking outside the selection first selects the clicked cell; clicking within a range keeps that range. The menu offers clipboard, clear, row/column structure, formatting, comments and suggestion navigation, with disabled states matching the current mode and host permissions. Escape dismisses and restores focus. The vendor context menu shall stay disabled. Worksheet tabs shall expose a shared context menu for sheet actions.

Implementation: `src/editor/editor-menus.ts`, `src/editor/spreadsheet-editor.tsx`, `src/engine/workbook-grid.tsx`.

Verification: `e2e/full-editor.spec.ts` (SE-CONTEXT-1, SE-CHROME-2, SE-REVIEW-5).

## SE-CHROME-4 — Host interaction parity

`chrome` selects minimal, embedded or full controls. The full preset remains the default. Feature flags and host panels support templates and application embedding without platform services. The formula row has a fixed-width accessible address field, flexible entry area, Cancel/Apply, and a constrained calculated result. Formula entry supports Escape cancellation, Enter commit/navigation, partial references and genuine pointer reference picking. Native entry uses the same completion vocabulary. Menus and panels release vendor keyboard ownership; dialogs show inline validation and block grid input.

The engine root remains mounted through cell/style/geometry changes. Protected edits defer data/presentation replacement until commit/cancel. Fractional header resize previews only engine measure channels and commits one immutable workbook change on release; Escape, blur, pointer cancellation, permission loss and teardown restore original measures. Optional host blur subscriptions avoid requiring a global application event bus. Row/column header menus and header Delete/Backspace use canonical structural commands; changes intersecting merges require shared-modal confirmation. Worksheets support insert-before and wraparound left/right moves.

`onSave` receives a cancellation signal, stable request id and manual/autosave kind. `autosaveMs` opts in. Saving never starts merely from viewing or selecting. Manual saves coalesce and queue after autosave; later edits remain dirty. Only autosave 503 errors retry quietly with bounded backoff. Historical views force Reading and keep the host-controlled banner until the host restores the current version. The full preset keeps an additional dedicated Suggestions pane; hosts may replace built-in panels.

Verification: `src/editor/use-editor.test.ts`, `e2e/flight-parity.spec.ts`, `e2e/full-editor.spec.ts` and the pinned [comparison](HOST-PARITY.md).
