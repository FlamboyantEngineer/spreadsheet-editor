# Request to ui-components: spreadsheet formatting glyphs

Please add the following 31 reusable glyphs to the public [FlamboyantEngineer/ui-components](https://github.com/FlamboyantEngineer/ui-components) library. They support icon-only spreadsheet controls and labelled context menus in [spreadsheet-editor](https://github.com/FlamboyantEngineer/spreadsheet-editor). Keep the semantics generic so document tables and other tabular editors can use them too.

The editor currently uses existing public library glyphs through a single adapter, `src/editor/editor-commands.ts`, at the pinned public commit [`3f7adc110178af35a15f0ed82fc1cd3d9aa3189d`](https://github.com/FlamboyantEngineer/ui-components/tree/3f7adc110178af35a15f0ed82fc1cd3d9aa3189d). These interim glyphs express the command category; menu labels, accessible names and tooltips retain each command's precise meaning. Several choices share a glyph until dedicated marks are available. The editor contains no local SVG substitutes and needs no unpublished library commit to run its labs.

Export each component below from `@flamboyantengineer/ui-components/core` and add its exact provider key to `GlyphName` and `GlyphMap`:

| Export                    | Provider key           | Intended mark                               | Interim public glyph |
| ------------------------- | ---------------------- | ------------------------------------------- | -------------------- |
| `FontFamilyGlyph`         | `font-family`          | Font family                                 | `ParagraphGlyph`     |
| `FontSizeGlyph`           | `font-size`            | Font size                                   | `Heading1Glyph`      |
| `NumberFormatGlyph`       | `number-format`        | Numeric format                              | `ListNumberedGlyph`  |
| `PercentGlyph`            | `percent`              | Percentage                                  | `ListNumberedGlyph`  |
| `CurrencyGlyph`           | `currency`             | Currency                                    | `ListNumberedGlyph`  |
| `ScientificNotationGlyph` | `scientific-notation`  | Scientific notation                         | `FormulaGlyph`       |
| `FractionGlyph`           | `fraction`             | Fraction                                    | `FormulaGlyph`       |
| `UnitGlyph`               | `unit`                 | Measurement unit                            | `ResizeGlyph`        |
| `IncreaseDecimalsGlyph`   | `increase-decimals`    | Increase displayed decimal places           | `PlusGlyph`          |
| `DecreaseDecimalsGlyph`   | `decrease-decimals`    | Decrease displayed decimal places           | `MinusGlyph`         |
| `CellAlignmentGlyph`      | `cell-alignment`       | Cell alignment chooser                      | `ParagraphGlyph`     |
| `AlignTopLeftGlyph`       | `align-top-left`       | Text at top left of a cell                  | `ArrowUpGlyph`       |
| `AlignTopCenterGlyph`     | `align-top-center`     | Text at top center of a cell                | `ArrowUpGlyph`       |
| `AlignTopRightGlyph`      | `align-top-right`      | Text at top right of a cell                 | `ArrowUpGlyph`       |
| `AlignMiddleLeftGlyph`    | `align-middle-left`    | Text at middle left of a cell               | `ArrowLeftGlyph`     |
| `AlignMiddleCenterGlyph`  | `align-middle-center`  | Text at middle center of a cell             | `DragHandleGlyph`    |
| `AlignMiddleRightGlyph`   | `align-middle-right`   | Text at middle right of a cell              | `ArrowRightGlyph`    |
| `AlignBottomLeftGlyph`    | `align-bottom-left`    | Text at bottom left of a cell               | `ArrowDownGlyph`     |
| `AlignBottomCenterGlyph`  | `align-bottom-center`  | Text at bottom center of a cell             | `ArrowDownGlyph`     |
| `AlignBottomRightGlyph`   | `align-bottom-right`   | Text at bottom right of a cell              | `ArrowDownGlyph`     |
| `TableBorderTopGlyph`     | `table-border-top`     | Top cell border                             | `ArrowUpGlyph`       |
| `TableBorderRightGlyph`   | `table-border-right`   | Right cell border                           | `ArrowRightGlyph`    |
| `TableBorderBottomGlyph`  | `table-border-bottom`  | Bottom cell border                          | `ArrowDownGlyph`     |
| `TableBorderLeftGlyph`    | `table-border-left`    | Left cell border                            | `ArrowLeftGlyph`     |
| `TableBorderAllGlyph`     | `table-border-all`     | All table borders, including internal lines | `TableGlyph`         |
| `TableBorderOutsideGlyph` | `table-border-outside` | Outside table border only                   | `StopGlyph`          |
| `TableBorderNoneGlyph`    | `table-border-none`    | Remove table borders                        | `NoSymbolGlyph`      |
| `WrapTextGlyph`           | `wrap-text`            | Wrap text within a cell                     | `ParagraphGlyph`     |
| `MergeCellsGlyph`         | `merge-cells`          | Merge selected cells                        | `LinkGlyph`          |
| `UnmergeCellsGlyph`       | `unmerge-cells`        | Unmerge cells                               | `UnlinkGlyph`        |
| `CutGlyph`                | `cut`                  | Cut to clipboard, conventionally scissors   | `ClipboardGlyph`     |

Match the library's public glyph contract:

- Use the existing normalized outline system: 24×24 view box, current color, shared stroke treatment and provider-aware `createGlyph` construction.
- Support the existing `Glyph`, `Icon`, `Button`, `DropdownMenu` and `MenuEntry.icon` APIs. Each glyph must expose its provider key and allow `UiProvider.icons` overrides without breaking those consumers.
- Keep SVG output decorative by default; the enclosing control supplies its accessible name. Preserve caller SVG props and the established SSR, hydration and client-boundary behavior.
- Keep each export tree-shakeable, with accurate public TypeScript declarations and no new editor, styling or browser-global dependencies.
- Alignments must clearly distinguish all nine combinations; borders must distinguish the four sides, all lines, outside only and none. Decimal changes should remain distinguishable from zoom or row/column insertion.
- Check legibility at 16px and 24px in light and dark themes. Add catalogue/lab examples, provider-override and public-export tests, then run the library's required checks.

An implementation was prepared and checked locally as commit `e847833e0983ee2734c0af5155eb5ef7d5dcd6dc` on branch `feat/tabular-formatting-glyphs`, based on the public commit above. That commit is not a published dependency: branch publication returned HTTP 403 for the available `FlamboyantEngineer` account. The [ready-to-apply patch](ui-components-tabular-formatting-glyphs.patch) includes the implementation, tests, lab examples and minor changeset. A maintainer can review and publish the prepared changes or implement the same public contract. After the exports are available at a public Git commit, pin that commit in spreadsheet-editor and replace only the aliases in `editor-commands.ts` with the requested exports.
