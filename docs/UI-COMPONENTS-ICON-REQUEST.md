# Public spreadsheet glyphs

The editor uses the public [FlamboyantEngineer/ui-components](https://github.com/FlamboyantEngineer/ui-components) library, pinned to [`d85a044a7c2466902cc0b88acbd9dd3b8260f03a`](https://github.com/FlamboyantEngineer/ui-components/tree/d85a044a7c2466902cc0b88acbd9dd3b8260f03a). The original formatting-glyph request is fulfilled; the earlier prototype patch and category-icon substitutions are retired.

All formatting glyphs come directly from the public `/core` exports through `src/editor/editor-commands.ts`. Comment and suggestion controls also use the public comment glyphs.

| Editor control                                      | Public glyphs                                                                                                                                                              |
| --------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Font family and size                                | `FontFamilyGlyph`, `FontSizeGlyph`                                                                                                                                         |
| Font size stepping                                  | `IncreaseFontSizeGlyph`, `DecreaseFontSizeGlyph`                                                                                                                           |
| Text and cell fill colors                           | `FontColorGlyph`, `BackgroundColorGlyph` above public `ColorSwatchButton` bars                                                                                             |
| Number formats                                      | `NumberFormatGlyph`, `PercentGlyph`, `CurrencyGlyph`, `ScientificNotationGlyph`, `FractionGlyph`, `UnitGlyph`                                                              |
| Decimal places                                      | `IncreaseDecimalsGlyph`, `DecreaseDecimalsGlyph`                                                                                                                           |
| Alignment chooser                                   | `CellAlignmentGlyph`                                                                                                                                                       |
| Cell alignment choices                              | All nine `AlignTopLeftGlyph` through `AlignBottomRightGlyph` positions                                                                                                     |
| Border choices                                      | `TableBorderTopGlyph`, `TableBorderRightGlyph`, `TableBorderBottomGlyph`, `TableBorderLeftGlyph`, `TableBorderAllGlyph`, `TableBorderOutsideGlyph`, `TableBorderNoneGlyph` |
| Wrap / merge / unmerge                              | `WrapTextGlyph`, `MergeCellsGlyph`, `UnmergeCellsGlyph`                                                                                                                    |
| Cut                                                 | `CutGlyph`                                                                                                                                                                 |
| Add row below                                       | `LayoutAddBelowGlyph`                                                                                                                                                      |
| Add comment / accept or reject suggestion           | `ChatBubblePlusGlyph`, `ChatBubbleCheckGlyph`, `ChatBubbleXMarkGlyph`                                                                                                      |
| Workbook and Excel commands / CSV and JSON commands | `FileSpreadsheetGlyph`, `FileDataGlyph`                                                                                                                                    |

The update also supplies revised Save, warning and other existing glyph drawings. The embedded action menu retains its public three-line `ParagraphGlyph`.

The public NumberField supplies preset stepping and arbitrary fractional entry, with its size glyphs on the steppers. ColorSwatchButton, ColorSwatchPicker, ColorField and Popover supply color controls in both toolbar presets. The editor adds no substitute glyphs or control skins; its reference CSS only arranges these controls.

Glyphs retain the library's `UiProvider.icons` overrides. Border choices preserve the library's multiple strokes and dotted inactive edges. The editor adds no local SVG substitutes. Keep the manifest and lockfile pinned to the same reviewed public commit when updating the library.
