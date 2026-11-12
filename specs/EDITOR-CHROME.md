# Editor chrome

## SE-CHROME-1 — Full reusable shell

The editor shall provide File, Edit, Insert, Format, View and Review menus; a scrolling formatting toolbar; an address/formula bar with function completions; worksheet tabs; a find/replace bar; docked Properties, History, Comments and Suggestions panes; accessible selection count/sum/average; and worksheet-strip zoom controls. Generic controls use `@flamboyantengineer/ui-components`. The host styles documented hooks; grid zoom is engine presentation geometry, independent of chrome layout.

Formatting shall support bold, italic, strikethrough, font family/size, text/fill colors, numeric formats and decimal places, horizontal/vertical alignment, borders, wrapping and merge/unmerge. Numeric formats round only display text, preserving exact rational calculations. Undo/redo includes formatting and review state. A frozen-column setting and display toggles expose grid integration without adding application services.

Authoring writes canonical data; Reviewing stages cell/format proposals and privately previews the local reviewer’s pending values; Reading refuses mutation. `readOnly` forces Reading, and host permissions independently gate editing, commenting, proposing and deciding suggestions in controls, commands and native engine callbacks. Optional host callbacks expose Save and Refresh; extra host panes extend the built-in panel set. Save errors are visible and edits made during a save remain dirty.

Implementation: `src/editor/contracts.ts`, `src/editor/editor-toolbar.tsx`, `src/editor/editor-menus.ts`, `src/editor/spreadsheet-editor.tsx`, `src/editor/use-editor.ts`, `src/calculation/cell-format.ts`.

Verification: `e2e/full-editor.spec.ts` (SE-CHROME-1, SE-REVIEW-5), `src/calculation/cell-format.test.ts`, existing browser editor tests.

## SE-CHROME-2 — Worksheet and structural commands

Worksheet actions shall add, duplicate, rename, reorder and delete sheets, retaining at least one. Row/column commands shall insert before/after and delete selections, move authored formatting and merge spans, rewrite qualified/unqualified and absolute formula references across worksheets, and retain quoted literals and scientific numeric text unchanged. References to deleted cells/sheets become `#REF!`; partially deleted ranges shrink. Comment anchors move with cells; wholly deleted anchors become orphaned. Affected suggestions become outdated instead of applying to another cell. Imports accept CSV or the versioned workbook JSON; exports expose the complete workbook JSON.

Implementation: `src/model/structure.ts`, `src/model/editing.ts`, `src/editor/editor-menus.ts`.

Verification: `src/model/structure.test.ts`, `e2e/full-editor.spec.ts` (SE-CONTEXT-1, SE-CHROME-2).

## SE-CHROME-3 — Search and keyboard ownership

Find/replace shall search raw values/formulas and current formatted result text across worksheets, deduplicate cells, navigate to matches and replace authored matches one/all through the normal author/review mutation path. Result-only matches are navigable and never replace formulas. Pending or stale result text is excluded. Case-insensitive literal queries must not execute as regular expressions. Typed search/comment/dialog content must never reach the vendor grid's document keyboard handler. Each instance owns its selection, menus and panels. Clipboard text preserves numeric strings and quoting; a host may supply clipboard adapters.

Implementation: `src/editor/use-editor.ts`, `src/engine/workbook-grid.tsx`, `src/model/editing.ts`.

Verification: `e2e/full-editor.spec.ts` (SE-CHROME-3, SE-REVIEW-6), `src/model/structure.test.ts`.

## SE-CONTEXT-1 — Shared React context menu

Secondary click, Shift+F10 and the Context Menu key shall open the shared React context menu for the selected cells. Clicking outside the selection first selects the clicked cell; clicking within a range keeps that range. The menu offers clipboard, clear, row/column structure, formatting, comments and suggestion navigation, with disabled states matching the current mode and host permissions. Escape dismisses and restores focus. The vendor context menu shall stay disabled. Worksheet tabs shall expose a shared context menu for sheet actions. Context commands show public library glyphs and only supported shortcut hints. The pinned library supplies dedicated font, numeric-format, decimal, alignment, border, wrap, merge/unmerge, cut, add-row, comment and file glyphs, documented in `docs/UI-COMPONENTS-ICON-REQUEST.md`. Secondary-clicking an inactive worksheet does not activate it.

Implementation: `src/editor/editor-menus.ts`, `src/editor/spreadsheet-editor.tsx`, `src/engine/workbook-grid.tsx`.

Verification: `e2e/full-editor.spec.ts` (SE-CONTEXT-1, SE-CHROME-2, SE-REVIEW-5).

## SE-CHROME-4 — interaction parity

`chrome` selects minimal, embedded or full controls. The full preset remains the default. Feature flags and host panels support templates and application embedding without platform services. The formula row composes the public FormulaBar with a fixed-width accessible address field, a flexible CodeEditor surface, and a constrained calculated result. It has no Cancel, Apply or Restore buttons; Enter commits, Escape cancels, and leaving the editor commits. Completion acceptance and IME composition take priority over commit/navigation. Pending edits are preserved while using editor controls and applied atomically with formatting or worksheet changes. Formula entry supports Escape cancellation, Enter commit/navigation, partial references and genuine pointer reference picking. Native entry uses the same completion vocabulary. Menus and panels release vendor keyboard ownership; dialogs show inline validation and block grid input.

The engine root remains mounted through cell/style/geometry changes. Protected edits defer data/presentation replacement until commit/cancel. Fractional header resize previews only engine measure channels and commits one immutable workbook change on release; Escape, blur, pointer cancellation, permission loss and teardown restore original measures. Optional host blur subscriptions avoid requiring a global application event bus. Row/column header menus and header Delete/Backspace use canonical structural commands; changes intersecting merges require shared-modal confirmation. Worksheets support insert-before and wraparound left/right moves.

Native text entry shall fit inside the existing cell or complete merged span without changing row heights, column widths or neighboring cell positions. Its input stays out of table sizing, inherits the cell's font, padding and text color, and restores temporary positioning on commit/cancel. Frozen anchors retain sticky positioning. Text and formula-reference highlighting remain visible in both themes; the formula mirror follows single-line input alignment instead of moving merged-cell text to the top of its span.

`onSave` receives a cancellation signal, stable request id and manual/autosave kind. `autosaveMs` opts in. Saving never starts merely from viewing or selecting. Manual saves coalesce and queue after autosave; later edits remain dirty. Only autosave 503 errors retry quietly with bounded backoff. Historical views force Reading and keep the host-controlled banner until the host restores the current version. The full preset keeps an additional dedicated Suggestions pane; hosts may replace built-in panels.

Verification: `src/editor/use-editor.test.ts`, `e2e/native-cell-editing.spec.ts`, `e2e/interaction-parity.spec.ts`, `e2e/full-editor.spec.ts` and the pinned [comparison](HOST-COMPATIBILITY.md).

## SE-CHROME-5 — Grid zoom and visual integration

Zoom scales grid cell text, row heights, column widths and headers inside the fixed grid viewport. Editor width/height, toolbar, formula bar, worksheet tabs, sidebar and zoom controls do not scale. Default and authored grid borders and selection/reference/presence strokes remain one CSS pixel at every supported zoom level. Resizing headers converts pointer deltas back to canonical workbook dimensions; changing zoom never authors workbook geometry. Visible selection edges enclose all four sides in light and dark themes, including ranges clipped by frozen columns or scrolling.

During a pointer drag, the local perimeter shall follow the engine's current highlight before the same paint. Geometry work coalesces into one queued animation frame without postponing it for later pointer events or waiting for parent selection/statistics updates. Ordinary rectangular ranges measure their corners instead of scanning every selected cell; merged or hidden endpoints retain their complete bounds. Controlled selection echoes do not interrupt the live drag or scroll its anchor back into view.

Selection, formula-reference and presence overlay containers shall be clipped to the visible cell body, excluding both sticky header bands and scrollbar gutters. Header bounds use the live row and column headers as well as the corner, including at fractional zoom and across frozen columns. The vendor fill handle shall also be clipped there, so an offscreen selection cannot paint or intercept clicks over headers; scrolling back restores its visible handle.

The embedded toolbar uses icon-only action triggers and the public ToolbarRow default appearance. All triggers have accessible names and tooltips. Full mode retains labelled application menus. Worksheet controls use public Tabs with a scrollable tab list with Add immediately after the final tab and a separate fixed zoom control, without an ellipsis menu, close marks or a zoom-reset control. Double-click, F2 and Rename open editing directly inside the worksheet tab. Enter/blur commits a valid unique normalized name, Escape cancels, and invalid Enter submissions remain editable with feedback (leaving an invalid field cancels). Rename and structural commands preserve pending formula edits and remap references.

Implementation: `src/engine/workbook-grid.tsx`, `src/engine/worksheet-presentation.ts`, `src/engine/header-resize.ts`, `src/editor/editor-formula-bar.tsx`, `src/editor/worksheet-strip.tsx`, `src/editor/editor-toolbar.tsx`, `labs/src/reference.css`.

Verification: `src/engine/grid-geometry.test.ts`, `src/engine/selection-publication.test.ts`, `e2e/selection-rendering.spec.ts`, and existing editor suites cover geometry, single-frame drag updates, native fill-handle clipping, scrolling, frozen columns, inline rename and formula ownership across the three lab presets.

## SE-CHROME-6 — Chrome surfaces and grid boundaries

Cell secondary clicks cancel the browser context menu before opening the shared application menu. Native text-editing fields retain their own text context menu. Header context actions and selection behavior remain available.

Unstyled cells use the library canvas token (white in light mode); row and column headers use the slightly shaded surface token. Authored fills and selected-cell/header states retain their meaning. Toolbar rows use the canvas background and popup menus keep the library's default menu surface. The full menubar spans the editor width without a border. Full-mode action buttons share one scrolling icon toolbar; font, size, color, number-format and unit fields remain readable controls. The embedded action trigger uses the public three-line glyph and a standard bordered button variant, inset from the toolbar edge. FormulaBar has no extra surrounding box; the address and formula fields retain their normal borders.

Sticky header backgrounds stay opaque, including selected headers: the selection token is tinted over the surface instead of exposing cells beneath it. Column headers own their bottom border, and row headers and the corner own their right border, so these one-pixel edges stay visible while cells scroll underneath. The adjacent resting cell omits its duplicate default seam; authored and clipboard borders retain their own paint. Header content pixels remain unchanged during vertical/horizontal scrolling in both themes and at fractional zoom.

Add Worksheet sits immediately after the last worksheet inside the scrolling tab list, including when tabs overflow. Zoom remains separate at the trailing end or on a second mobile row. The grid reserves its own vertical scrollbar gutter and confines overlays to its viewport; a collapsed or expanded dock panel cannot cover the scrollbar.

Verification: `e2e/sticky-headers.spec.ts`, `e2e/chrome-layout.spec.ts`, `src/editor/editor-toolbar.test.tsx`, existing editor and lab browser suites.

## SE-CHROME-7 — Text size and color controls

The full toolbar shall use the public NumberField for font size, with IncreaseFontSizeGlyph and DecreaseFontSizeGlyph step buttons. Buttons and plain arrow keys move to the nearest larger or smaller preset; an arbitrary size, including a fractional size, remains valid between presets. Sizes are CSS pixels, with a visible px suffix, bounded from 1 to 400. Embedded mode opens the same field from its named font-size icon, preserving the icon-only toolbar. Field edits apply on Enter, blur or a step, without restore/apply/cancel controls.

Both full and embedded toolbars shall expose public ColorSwatchButton controls for Text color and Cell fill. FontColorGlyph and BackgroundColorGlyph sit above a bar showing the authored color, or an empty bar when unset. Each button opens a public Popover with ColorSwatchPicker presets, a ColorField for custom entry and its full picker, and a Clear command that removes only that color property. Picking a preset or clearing closes the palette and restores focus to its trigger. The controls retain provider glyph overrides and the library's default appearance. The reference stylesheet owns only editor layout and sizing.

These controls use the existing formatting transaction path: they format the selected range, compose pending formula input atomically, support undo/redo and stage review suggestions. Disabled editing prevents changes and closes open formatting popovers. No calculation or serialized workbook contract changes.

Implementation: `src/editor/text-format-controls.tsx`, `src/editor/editor-toolbar.tsx`, `src/editor/editor-commands.ts`.

Verification: `src/editor/editor-toolbar.test.tsx`, `e2e/text-format-controls.spec.ts`, `e2e/labs.spec.ts`, `e2e/chrome-layout.spec.ts`.
