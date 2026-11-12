# Styling hooks

Package source loads no CSS. The host provides vendor structural grid CSS and styles React chrome through these public hooks. Base library controls retain their `ui-*` hooks.

| Hook                                                                                     | Ownership                                                           |
| ---------------------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| `.se-editor`                                                                             | Editor root, accepts `className` and an accessible label            |
| `[data-calculation-mode="client\|server"]`                                               | Selected calculation location on the root                           |
| `[data-pending="true\|false"]`                                                           | Current calculation state on the root                               |
| `.se-editor__toolbar`                                                                    | Editing commands                                                    |
| `.se-editor__header`, `.se-editor__title`                                                | Workbook title, mode and host save controls                         |
| `.se-editor__menubar`                                                                    | Full-width main menu bar                                            |
| `.se-editor__formula-bar`                                                                | Raw formula/value field and calculated output                       |
| `.se-editor__address`                                                                    | Active A1 address                                                   |
| `.se-editor__engine`                                                                     | Vendor-owned grid root                                              |
| `.se-editor__sheets`                                                                     | Worksheet tabs with inline Add and separate Zoom                    |
| `.se-editor__insertion`                                                                  | Both insertion commands                                             |
| `.se-editor__status`                                                                     | Visually hidden calculation status                                  |
| `.se-editor__workspace`, `.se-editor__main`, `.se-editor__sidebar`                       | Grid and docked panels                                              |
| `.se-editor__grid`                                                                       | Fixed grid viewport; `data-gridlines` toggles gridline presentation |
| `.se-editor__find`                                                                       | Find/replace controls                                               |
| `[data-mode="author\|review\|read"]`                                                     | Current interaction mode on the editor root                         |
| `.se-comments`, `.se-comments__actions`                                                  | Comment pane and filter/actions                                     |
| `.se-suggestions`, `.se-suggestion`, `.se-suggestion__change`, `.se-suggestion__actions` | Suggestion pane and review decisions                                |
| `.se-properties`, `.se-history`                                                          | Built-in metadata and session history panels                        |
| `.se-editor__notice`                                                                     | Action failure message and dismissal                                |
| `td[data-approximate="true\|false"]`                                                     | Calculated engine cell approximation status                         |
| `td[data-se-comments]`, `td[data-se-suggestions]`                                        | Open comment/pending suggestion count at an engine cell             |

Vendor cell markup remains an engine boundary. Hosts should limit vendor selectors to structural integration and prefer their own authored cell styles for content. `labs/src/reference.css` is a standalone presentation example.

Additional public hooks: root `[data-chrome="minimal|embedded|full"]`; `.se-editor__tool-group`; `.se-editor__formula-input`, `.se-editor__formula-widget`; `.se-editor__grid-surface`, `.se-editor__native-formula`; `.se-editor__worksheet-tabs`, `.se-editor__worksheet-panel`, `.se-editor__sheet-tabs`; `.se-editor__reference-frame`, `.se-editor__reference-outline`, `.se-editor__selection-outline`; `.se-editor__presence-outline[data-kind="selection|lock"][data-color]`; `.se-editor__historical`; `.se-comments__conversation`; `.se-editor__notice[data-tone="warning|danger"]`; `td[data-se-contributor="unknown|0…7"]` and `td[data-se-ai="true"]`. Formula CodeEditor tokens use semantic `reference-0` through `reference-7` types. Contributor/presence colors are presentation attributes and never authored cell styles. Hosts control the actual palette. Selection outlines and comment markers must not intercept pointer events or block text spill.

The public library owns ToolbarRow, FormulaBar/CodeEditor and Tabs control appearance. Hosts should preserve default control metrics and focus treatment. Grid zoom is applied through engine presentation sizes and text metrics, with strokes kept at one CSS pixel; hosts must not apply CSS `zoom` or a transform to the grid ancestor. The engine positions `.se-editor__reference-frame` over the cell body, clipping all outlines below column headers, beside row headers and before scrollbar gutters. Child outline coordinates are relative to that clipped frame. Selection fragments can omit edges at viewport/frozen clipping boundaries; the engine provides functional positions and the host supplies their stroke color. Native fill-handle clipping is engine-owned too.

Host presentation adds `.se-editor__cell-content` for React portal targets and cell attributes `data-se-decoration` and `data-se-presentation-status`. Decorations are bounded semantic tokens. Titles, accessible descriptions, attributes and portal content clear when publications, sheets or sessions change. `[data-calculation-mode="external"]` identifies authoritative calculation mode. The host owns palette, rich-content sanitation and links.

`.se-editor__font-size` sizes the public NumberField, including its font-size step glyphs and px suffix. `.se-editor__color-palette` lays out the public swatch picker, custom ColorField and Clear command within a public Popover. Color buttons retain `ui-color-swatch-button` and its swatch hooks; hosts should preserve their glyph/bar composition and accessible color descriptions. Embedded font-size entry uses the same NumberField inside a Popover.
