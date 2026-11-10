# Styling hooks

Package source loads no CSS. The host provides vendor structural grid CSS and styles React chrome through these public hooks. Base library controls retain their `ui-*` hooks.

| Hook                                                                                     | Ownership                                                                                                    |
| ---------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `.se-editor`                                                                             | Editor root, accepts `className` and an accessible label                                                     |
| `[data-calculation-mode="client\|server"]`                                               | Selected calculation location on the root                                                                    |
| `[data-pending="true\|false"]`                                                           | Current calculation state on the root                                                                        |
| `.se-editor__toolbar`                                                                    | Editing commands                                                                                             |
| `.se-editor__header`, `.se-editor__title`                                                | Workbook title, mode and host save controls                                                                  |
| `.se-editor__menubar`, `.se-editor__actions`                                             | Main menus and secondary commands                                                                            |
| `.se-editor__formula-bar`                                                                | Raw formula/value field and calculated output                                                                |
| `.se-editor__address`                                                                    | Active A1 address                                                                                            |
| `.se-editor__engine`                                                                     | Vendor-owned grid root                                                                                       |
| `.se-editor__sheets`                                                                     | Worksheet buttons                                                                                            |
| `.se-editor__insertion`                                                                  | Both insertion commands                                                                                      |
| `.se-editor__status`                                                                     | Status or provider error                                                                                     |
| `.se-editor__workspace`, `.se-editor__main`, `.se-editor__sidebar`                       | Grid and docked panels                                                                                       |
| `.se-editor__grid`                                                                       | Grid viewport; `--se-zoom` is the requested scale, 1 at 100%; `data-gridlines` toggles gridline presentation |
| `.se-editor__find`, `.se-editor__statusbar`, `.se-editor__summary`                       | Find/replace and bottom status controls                                                                      |
| `[data-mode="author\|review\|read"]`                                                     | Current interaction mode on the editor root                                                                  |
| `.se-comments`, `.se-comments__actions`                                                  | Comment pane and filter/actions                                                                              |
| `.se-suggestions`, `.se-suggestion`, `.se-suggestion__change`, `.se-suggestion__actions` | Suggestion pane and review decisions                                                                         |
| `.se-properties`, `.se-history`                                                          | Built-in metadata and session history panels                                                                 |
| `.se-editor__notice`                                                                     | Action failure message and dismissal                                                                         |
| `td[data-approximate="true\|false"]`                                                     | Calculated engine cell approximation status                                                                  |
| `td[data-se-comments]`, `td[data-se-suggestions]`                                        | Open comment/pending suggestion count at an engine cell                                                      |

Vendor cell markup remains an engine boundary. Hosts should limit vendor selectors to structural integration and prefer their own authored cell styles for content. `labs/src/reference.css` is a standalone presentation example.
