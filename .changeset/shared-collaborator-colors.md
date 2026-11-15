---
'@flamboyantengineer/spreadsheet-editor': patch
---

A collaborator's colour slot now comes from the base library's `getColorIndex`, so a person has the same slot in the spreadsheet, in the base library's avatars and in the other editors. A person whose id holds characters outside the Basic Multilingual Plane, or leading or trailing spaces, may move to another slot.
