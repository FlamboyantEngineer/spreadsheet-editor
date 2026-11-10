import { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { UiProvider } from '@flamboyantengineer/ui-components/core';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { SpreadsheetEditor } from '../../src/index.js';
import {
  createHttpCalculator,
  createClientCalculator,
  calculateWorkbook,
  formatCellValue,
} from '../../src/calculation/index.js';
import { createDocumentInsertion, type DocumentInsertion } from '../../src/document/index.js';
import type { Workbook } from '../../src/model/index.js';
import 'jsuites/dist/jsuites.css';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import './reference.css';

const sample: Workbook = {
  schemaVersion: 1,
  worksheets: [
    {
      id: 'calculations',
      name: 'Calculations',
      data: [
        ['Example', 'Value', 'Check'],
        ['Small number', '1E-25', '=B2=0'],
        ['Near one', '=1-B2', '=B3=1'],
        ['Recover difference', '=1-B3', '=B4=B2'],
        ['Exact third', '=1/3', '=B5*3'],
        ['Approximate root', '=SQRT(2)', '=B6*B6'],
        ['Cross-sheet', '=Inputs!A1*B2', '=SUM(B2:B4)'],
      ],
      columns: [{ width: 200 }, { width: 300 }, { width: 260 }],
    },
    {
      id: 'inputs',
      name: 'Inputs',
      data: [['3'], ['=Calculations!B3']],
      columns: [{ width: 300 }],
    },
  ],
};

function TablePreview({ insertion }: { insertion: Extract<DocumentInsertion, { kind: 'table' }> }) {
  return (
    <table className="lab-document-table">
      <caption>Inserted document table</caption>
      <tbody>
        {insertion.rows.map((row, y) => (
          <tr key={y}>
            {row.map((cell, x) => (
              <td key={x} colSpan={cell.colspan} rowSpan={cell.rowspan}>
                {cell.text}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function App() {
  const [workbook, setWorkbook] = useState(sample);
  const [mode, setMode] = useState<'client' | 'server'>('client');
  const [digits, setDigits] = useState('32');
  const [insertion, setInsertion] = useState<DocumentInsertion | null>(null);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [previewPending, setPreviewPending] = useState(false);
  const server = useMemo(() => createHttpCalculator('/api/calculate'), []);
  const client = useMemo(() => createClientCalculator(), []);
  const display = useMemo(
    () => ({
      significantDigits: /^\d+$/.test(digits) ? Math.min(1000, Math.max(1, Number(digits))) : 32,
    }),
    [digits],
  );
  const examples = useMemo(() => calculateWorkbook(sample), []);
  const selectedMode = useMemo(
    () => (mode === 'server' ? ({ mode, provider: server } as const) : ({ mode } as const)),
    [mode, server],
  );
  async function previewTable() {
    setPreviewPending(true);
    setPreviewError(null);
    try {
      const result = await (mode === 'server' ? server : client).calculate(workbook);
      setInsertion(
        createDocumentInsertion(workbook, result, { mode: 'table', range: 'A1:C7', display }),
      );
    } catch (error) {
      setPreviewError(error instanceof Error ? error.message : 'Preview failed.');
    } finally {
      setPreviewPending(false);
    }
  }
  return (
    <UiProvider>
      <main>
        <header>
          <p className="lab-eyebrow">Spreadsheet editor / labs</p>
          <h1>Small numbers. Exact answers.</h1>
          <p>
            A reusable spreadsheet with exact rational arithmetic, client or server calculation, and
            document table insertion.
          </p>
        </header>
        <div className="lab-controls" role="group" aria-label="Lab settings">
          <Button isSelected={mode === 'client'} onPress={() => setMode('client')}>
            Calculate on client
          </Button>
          <Button isSelected={mode === 'server'} onPress={() => setMode('server')}>
            Calculate on server
          </Button>
          <TextField label="Display significant digits" value={digits} onChange={setDigits} />
          <Button
            onPress={() => {
              setWorkbook(structuredClone(sample));
              setInsertion(null);
            }}
          >
            Reset workbook
          </Button>
        </div>
        <section aria-label="Precision examples" className="lab-examples">
          {[
            ['1E-25', 1],
            ['1 − 1E-25', 2],
            ['1 − (1 − 1E-25)', 3],
          ].map(([label, row]) => (
            <div key={label}>
              <span>{label}</span>
              <output>
                {formatCellValue(examples.worksheets[0].values[Number(row)][1], display)}
              </output>
            </div>
          ))}
        </section>
        <SpreadsheetEditor
          value={workbook}
          onChange={setWorkbook}
          calculation={selectedMode}
          display={display}
          onInsert={setInsertion}
        />
        <p className="lab-hint">
          Select cells in the grid, edit a value in the formula bar or by double-clicking a cell,
          then choose an insertion mode. Approximate results carry ≈.
        </p>
        <section className="lab-document" aria-label="Document insertion preview">
          <h2>Document insertion</h2>
          <p>
            The live mode keeps the source workbook and its formulas. The normal table is a snapshot
            of calculated text.
          </p>
          <Button onPress={previewTable} isDisabled={previewPending}>
            Preview full table
          </Button>
          <Button
            onPress={() =>
              setInsertion(
                createDocumentInsertion(workbook, null, { mode: 'spreadsheet', range: 'A1:C7' }),
              )
            }
          >
            Preview live spreadsheet
          </Button>
          {previewError && <p role="alert">{previewError}</p>}
          {insertion?.kind === 'table' && <TablePreview insertion={insertion} />}
          {insertion?.kind === 'spreadsheet' && (
            <SpreadsheetEditor
              key={JSON.stringify(insertion.workbook)}
              aria-label="Inserted live spreadsheet"
              defaultValue={insertion.workbook}
              initialSheetId={insertion.sheetId}
              calculation={selectedMode}
              display={display}
            />
          )}
          {insertion && (
            <details>
              <summary>Insertion payload</summary>
              <pre>{JSON.stringify(insertion, null, 2)}</pre>
            </details>
          )}
        </section>
      </main>
    </UiProvider>
  );
}

createRoot(document.getElementById('root')!).render(<App />);
