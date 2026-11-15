import { useState } from 'react';
import { createMemoryHost } from '../../src/host/index.js';
import { type CellInput, createWorkbook, SpreadsheetEditor } from '../../src/index.js';

const alex = createMemoryHost({ reader: { id: 'alex', name: 'Alex' } }).identity;

const COLUMNS = 25;

/** Rows from `?rows=` (default 400, so 10,000 cells); the grid budget tests use the default. */
function requestedRows() {
  const rows = Number(new URLSearchParams(location.search).get('rows'));
  return Number.isInteger(rows) && rows > 0 && rows <= 10_000 ? rows : 400;
}

/** Numbers in the first three columns, a row total in the last, text in between. */
function largeSheet(rows: number) {
  const data: CellInput[][] = [];
  for (let y = 0; y < rows; y++) {
    const row: CellInput[] = [];
    for (let x = 0; x < COLUMNS; x++) {
      if (x === COLUMNS - 1 && y > 0) row.push(`=SUM(A${y + 1}:C${y + 1})`);
      else if (x < 3) row.push(String((x + 1) * (y + 1)));
      else row.push(`Item ${y + 1}.${x + 1}`);
    }
    data.push(row);
  }
  return createWorkbook(data);
}

export function LargeLab() {
  const [rows] = useState(requestedRows);
  const [workbook, setWorkbook] = useState(() => largeSheet(rows));
  return (
    <section className="lab-demo" aria-label="Large worksheet demo">
      <p>
        {rows.toLocaleString('en')} rows by {COLUMNS} columns (
        {(rows * COLUMNS).toLocaleString('en')} cells). Add <code>?rows=2000</code> to the address
        for 50,000 cells.
      </p>
      <SpreadsheetEditor
        chrome="embedded"
        aria-label="Large spreadsheet"
        value={workbook}
        onChange={setWorkbook}
        identity={alex}
      />
    </section>
  );
}
