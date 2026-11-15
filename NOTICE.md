# Notices

## Provenance

Parts of this library began as code written for an earlier, private application and were rewritten here as standalone modules: the grid host and cell spill in `src/engine/`, and the formula parser, formula reference parser and spreadsheet addressing in `src/model/`. The exact rational calculation engine, the calculation orchestration and the units, calendar, clipboard, review, presence and Excel adapters are native to this repository. The numeric behaviour fixtures in `specs/NUMERIC-CORPUS.json` were adapted from the same earlier code; deliberate differences are listed in `specs/NUMERIC-INTENTIONAL-DIFFERENCES.json`, and `specs/HOST-COMPATIBILITY.md` records the integration boundaries and numerical choices.

The library imports no application services and loads no application stylesheets, fixed URLs or transport code.

## Third-party software

Jspreadsheet CE, its React adapter, jSuites, decimal.js, ExcelJS, fflate and parse5 retain their respective upstream licenses, distributed with those packages.

This repository is private and UNLICENSED.
