# Styling hooks

Package source loads no CSS. The host provides vendor structural grid CSS and styles React chrome through these public hooks. Base library controls retain their `ui-*` hooks.

| Hook                                       | Ownership                                                |
| ------------------------------------------ | -------------------------------------------------------- |
| `.se-editor`                               | Editor root, accepts `className` and an accessible label |
| `[data-calculation-mode="client\|server"]` | Selected calculation location on the root                |
| `[data-pending="true\|false"]`             | Current calculation state on the root                    |
| `.se-editor__toolbar`                      | Editing commands                                         |
| `.se-editor__formula-bar`                  | Raw formula/value field and calculated output            |
| `.se-editor__address`                      | Active A1 address                                        |
| `.se-editor__engine`                       | Vendor-owned grid root                                   |
| `.se-editor__sheets`                       | Worksheet buttons                                        |
| `.se-editor__insertion`                    | Both insertion commands                                  |
| `.se-editor__status`                       | Status or provider error                                 |
| `td[data-approximate="true\|false"]`       | Calculated engine cell approximation status              |

Vendor cell markup remains an engine boundary. Hosts should limit vendor selectors to structural integration and prefer their own authored cell styles for content. `labs/src/reference.css` is a standalone presentation example.
