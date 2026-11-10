# spreadsheet-editor

A standalone React spreadsheet editor built on [`@flamboyantengineer/ui-components`](https://github.com/FlamboyantEngineer/ui-components).

- Exact rational arithmetic, including scientific notation, large integers and cancellation. `1E-25` stays nonzero; `1-1E-25` stays distinct from `1`; `1-(1-1E-25)` returns exactly `1E-25`.
- Client or server calculations through the same headless evaluator. Server mode accepts a host provider or the included HTTP adapter and Fetch-standard route handler.
- Cross-sheet formulas, ranges, a formula bar, worksheet selection, undo/redo, row/column append, resizing, merged cells and multiple independent instances.
- Document insertion as an ordinary table snapshot or an editable spreadsheet retaining live formulas.
- Unstyled React chrome using the base library's Button and TextField. The grid uses the extracted Jspreadsheet React host. Hosts own presentation; labs provide a reference stylesheet.

This private `0.x` package is the standalone extraction stage. Consumer migrations are separate: neither the existing application nor `document-editor` is changed here. Application persistence, permissions, collaboration, reviews, history services and workflows remain host responsibilities.

## Run the labs

Use Node 22 or newer and Git access to the private base component library:

```sh
npm ci
npm run labs
```

Open **http://127.0.0.1:4173/**. The command builds the browser assets and starts a static server plus a real `/api/calculate` endpoint. It needs no database or application backend. Switch calculation mode, edit the workbook, change display precision, and try both document insertion modes.

```sh
npm run labs -- --port 4200 --host 0.0.0.0
```

The included server is a local development harness. Mount `handleCalculationRequest` behind your application's routing and authorization for a deployed server.

## Embed the editor

Install this repository at a reviewed Git commit together with your host's React, React DOM and `@flamboyantengineer/ui-components` versions. Git installation prepares the package; a source checkout uses `npm run build`.

```tsx
import { SpreadsheetEditor, createWorkbook } from '@flamboyantengineer/spreadsheet-editor';

<SpreadsheetEditor
  defaultValue={createWorkbook([['1E-25', '=1-A1', '=1-B1']])}
  onChange={(workbook) => saveWorkbook(workbook)}
  onInsert={(insertion) => insertIntoDocument(insertion)}
/>;
```

`value` / `onChange` gives the host control of the workbook. `defaultValue` provides instance-local state. Optional props include `readOnly`, `initialSheetId`, `display`, `calculation`, `className` and `aria-label`. A read-only instance hides editing commands and locks grid/formula editing.

The host loads the vendor grid's structural CSS:

```ts
import 'jsuites/dist/jsuites.css';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
```

The editor loads no CSS or application URLs itself. Style its documented [`se-*` hooks](specs/HOOK-CONTRACT.md), and the base components' public hooks, in your own stylesheet. The lab's `labs/src/reference.css` is an example, not a package import. The grid loads lazily; headless entry points and the root import are safe in Node without a DOM.

## Choose the calculation location

Client calculation is the default:

```tsx
<SpreadsheetEditor calculation={{ mode: 'client', options: { precision: 80 } }} />
```

Server calculation uses a provider:

```tsx
import { createHttpCalculator } from '@flamboyantengineer/spreadsheet-editor/calculation';

const provider = createHttpCalculator('/api/calculate');
<SpreadsheetEditor calculation={{ mode: 'server', provider }} />;
```

Create providers outside render or memoize them. Optional HTTP configuration accepts `fetch`, `headers`, `credentials` and a response byte limit. Implement `CalculationProvider.calculate(workbook, options, signal)` for another transport. The editor cancels superseded requests and discards stale responses even if a transport ignores cancellation. Server failures are shown explicitly; they never trigger local calculation or leave old formula results visible as current.

On a server that supports Web `Request` / `Response`:

```ts
import { handleCalculationRequest } from '@flamboyantengineer/spreadsheet-editor/calculation';

export async function POST(request: Request) {
  return handleCalculationRequest(request, {
    limits: { maxCells: 50_000, maxOperations: 200_000 },
  });
}
```

Host-defined names and synchronous functions can be supplied to `createClientCalculator(host)` or the handler's `host` option. They run where calculation runs. `calculateWorkbook(workbook, options, host)` is the synchronous headless API.

## Exact values and display

Store authored numbers as **strings**. A JavaScript number or JSON number may already have lost digits before this package receives it.

```ts
import {
  calculateWorkbook,
  formatCellValue,
} from '@flamboyantengineer/spreadsheet-editor/calculation';
import { createWorkbook } from '@flamboyantengineer/spreadsheet-editor/model';

const result = calculateWorkbook(createWorkbook([['=1/3', '=A1*3']]));
// A1: { kind: 'number', numerator: '1', denominator: '3', approximate: false }
// B1: { kind: 'number', numerator: '1', denominator: '1', approximate: false }
formatCellValue(result.worksheets[0].values[0][0], { significantDigits: 6 }); // '0.333333'
```

Addition, subtraction, multiplication, division, integer powers, comparisons, aggregates and perfect-square roots preserve exact rationals. Formatting rounds a copy; it does not change the stored number. Explicit `ROUND` functions do change the calculated value as requested.

Non-rational operations such as `SQRT(2)` and `SIN(1)` use an isolated decimal context: 80 significant digits by default, configurable from 16 to 1000. Their results and dependent arithmetic carry `approximate: true`, their working precision, and an `≈` display marker. Numerators and denominators are strings over HTTP, so transport never converts them through binary floating point.

Arithmetic has no fixed rounding precision for exact values. Resource limits still apply: default maximum numerator/denominator size is 10,000 digits, workbook size is 50,000 cells, and evaluation has operation and dependency-depth budgets. Exceeding an evaluation budget produces `#LIMIT!`; invalid or oversized request bodies are rejected. See [CALCULATION.md](specs/CALCULATION.md) for the supported function catalog and limits.

## Insert into a document

`onInsert` offers a choice for the selected range:

| Payload kind  | Content                                                            | Consumer integration                                                                                                            |
| ------------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------- |
| `table`       | Calculated text, merge spans, ordinary table JSON and escaped HTML | Insert `schema.nodeFromJSON(insertion.node)` with a document transaction, or use the HTML import path.                          |
| `spreadsheet` | A deep copy of the full source workbook, selected sheet and range  | Store the payload in an embedded node and mount `SpreadsheetEditor` in its React node view; persist changes through `onChange`. |

Live insertion retains the complete workbook so formulas can read cells and sheets outside the visible selection. It is a snapshot independent of the source, not a link to an external workbook. Normal tables contain the current calculated values with no live formulas. A normal-table insertion is disabled while results are pending or unavailable.

The server-safe `createDocumentInsertion(workbook, result, options)` helper exposes the same contract without mounting the editor. The existing `document-editor` table schema accepts the ordinary table JSON; registration and persistence of an embedded live node belong to its later consumer migration. See [DOCUMENT-INSERTION.md](specs/DOCUMENT-INSERTION.md).

## Entry points and checks

| Entry point                              | Contents                                                           |
| ---------------------------------------- | ------------------------------------------------------------------ |
| `@flamboyantengineer/spreadsheet-editor` | React editor and primary public types                              |
| `…/calculation`                          | Exact arithmetic, formulas, formatting, providers and HTTP handler |
| `…/model`                                | Versioned workbook, addressing and extracted formula parsers       |
| `…/document`                             | Both document insertion payloads and table projection              |
| `…/engine`                               | Advanced browser-only Jspreadsheet host and engine helpers         |

```sh
npm run check
```

Checks include lint, formatting, types, unit tests, a clean package build, built-entry-point imports in plain Node, and Chromium browser tests. Browser tests use `CHROMIUM_PATH`, system Chromium when available, or Playwright's browser (`npx playwright install chromium`). CI needs `UI_COMPONENTS_READ_TOKEN` to install the private base library, as in the companion libraries.

See [CONTRIBUTING.md](CONTRIBUTING.md), the [behavior specifications](specs/EDITOR.md) and [source provenance](NOTICE.md). This first package does not expose application-specific external formulas, unit/quantity evaluation, XLSX import/export, or middle-row/column reference rewriting; unknown functions return `#NAME?`. Hosts can use the formula adapter for application functions.
