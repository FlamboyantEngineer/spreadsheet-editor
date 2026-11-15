---
'@flamboyantengineer/spreadsheet-editor': minor
---

The grid now limits its work to the cells that changed. Typing and committing no longer repaint, restyle or rebuild every cell: committed values are written in place, zoom metrics are written for every cell only when zoom changes (and not at all at 100%), and shared borders are reconciled only around cells with an inline border or a clipboard edge. In the `./engine` entry, `renderCalculatedCells` takes an optional `repaint` argument that paints every cell again, and `renderReviewMarkers` returns whether it changed any cell.
