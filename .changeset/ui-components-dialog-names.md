---
'@flamboyantengineer/spreadsheet-editor': patch
---

Build on ui-components' shadcn-aligned names: the editor's dialog uses `Dialog` from `ui-components/dialog` and the sheet bar `TabsList`, `TabsTrigger` and `TabsContent`. The dialog now closes on an outside press, as ui-components dialogs do by default.
