import { useMemo, useState } from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { SpreadsheetEditor } from '../../src/index.js';
import {
  createHttpCalculator,
  createClientCalculator,
  calculateWorkbook,
  formatCellValue,
} from '../../src/calculation/index.js';
import { createDocumentInsertion, type DocumentInsertion } from '../../src/document/index.js';
import type { SpreadsheetHistoryEntry } from '../../src/editor/contracts.js';
import type { Workbook } from '../../src/model/index.js';

import { useExcelLab } from './excel-lab.js';
import { sample } from './sample.js';

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

export function FullLab() {
  const excel = useExcelLab();
  const [workbook, setWorkbook] = useState(sample);
  const [saved, setSaved] = useState<Workbook | null>(null);
  const [snapshots, setSnapshots] = useState<(SpreadsheetHistoryEntry & { workbook: Workbook })[]>(
    [],
  );
  const [historical, setHistorical] = useState<
    (SpreadsheetHistoryEntry & { workbook: Workbook }) | null
  >(null);
  const [reviewerOnly, setReviewerOnly] = useState(false);
  const [presence, setPresence] = useState(false);
  const [autosave, setAutosave] = useState(false);
  const [actorId, setActorId] = useState('alex');
  const [readOnly, setReadOnly] = useState(false);
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
    <>
      {excel.download}
      {excel.dialog}
      <section className="lab-demo" aria-label="Full editor demo">
        <SpreadsheetEditor
          onImportXlsx={excel.requestImport}
          onExportXlsx={excel.exportWorkbook}
          sessionKey={historical?.id ?? 'current'}
          value={historical?.workbook ?? workbook}
          historical={
            historical
              ? { timestamp: historical.timestamp, onReturnCurrent: () => setHistorical(null) }
              : undefined
          }
          history={{
            entries: snapshots,
            onOpen: (entry) =>
              setHistorical(snapshots.find((snapshot) => snapshot.id === entry.id) ?? null),
          }}
          autosaveMs={autosave ? 1000 : undefined}
          permissions={reviewerOnly ? { edit: false } : undefined}
          collaborators={
            presence
              ? [
                  {
                    id: 'morgan',
                    name: 'Morgan',
                    sheetId: 'calculations',
                    selection: { x1: 1, x2: 1, y1: 1, y2: 1 },
                    editLock: { x1: 1, x2: 1, y1: 1, y2: 1 },
                    activity: 'Editing',
                  },
                ]
              : []
          }
          getCommentLink={(id) => new URL(`?comment=${encodeURIComponent(id)}`, location.href).href}
          initialCommentId={new URLSearchParams(location.search).get('comment') ?? undefined}
          onChange={setWorkbook}
          calculation={selectedMode}
          display={display}
          onInsert={setInsertion}
          currentUser={{ id: actorId, name: actorId === 'alex' ? 'Alex' : 'Morgan' }}
          onSave={async (snapshot, context) => {
            if (context.signal.aborted) return;
            setSaved(snapshot);
            if (context.kind === 'manual')
              setSnapshots((entries) => [
                ...entries,
                {
                  id: crypto.randomUUID(),
                  timestamp: new Date().toISOString(),
                  author: { id: actorId, name: actorId === 'alex' ? 'Alex' : 'Morgan' },
                  message: 'Saved workbook version',
                  state: 'Draft',
                  revision: String(entries.length + 1),
                  workbook: snapshot,
                },
              ]);
          }}
          readOnly={readOnly}
        />
      </section>
      <section className="lab-section" aria-labelledby="settings-title">
        <h2 id="settings-title">Lab settings</h2>
        <p>
          Switch the calculator, act as another reviewer, or reload a saved snapshot. Changes stay
          in this lab session.
        </p>
        <div className="lab-controls" role="group" aria-label="Lab settings">
          <Button isSelected={mode === 'client'} onPress={() => setMode('client')}>
            Calculate on client
          </Button>
          <Button isSelected={mode === 'server'} onPress={() => setMode('server')}>
            Calculate on server
          </Button>
          <TextField label="Display significant digits" value={digits} onChange={setDigits} />
          <SelectField
            label="Acting as"
            value={actorId}
            onChange={(value) => setActorId(String(value))}
            options={[
              { id: 'alex', label: 'Alex' },
              { id: 'morgan', label: 'Morgan' },
            ]}
          />
          <Button isSelected={reviewerOnly} onPress={() => setReviewerOnly(!reviewerOnly)}>
            Reviewer permissions
          </Button>
          <Button isSelected={presence} onPress={() => setPresence(!presence)}>
            Show collaborator lock
          </Button>
          <Button isSelected={autosave} onPress={() => setAutosave(!autosave)}>
            Autosave after one second
          </Button>
          <Button isSelected={readOnly} onPress={() => setReadOnly(!readOnly)}>
            Read-only host
          </Button>
          <Button isDisabled={!saved} onPress={() => saved && setWorkbook(structuredClone(saved))}>
            Reload saved workbook
          </Button>
          <Button
            onPress={() => {
              setWorkbook(structuredClone(sample));
              setInsertion(null);
            }}
          >
            Reset workbook
          </Button>
        </div>
      </section>
      <section className="lab-section" aria-labelledby="precision-title">
        <h2 id="precision-title">Exact values</h2>
        <p>Display rounding never changes the values used in calculations.</p>
        <div aria-label="Precision examples" className="lab-examples">
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
        </div>
        <p className="lab-hint">
          Select cells in the grid, edit a value in the formula bar or by double-clicking a cell,
          then choose an insertion mode. Approximate results carry ≈.
        </p>
      </section>
      <section className="lab-section lab-document" aria-label="Document insertion preview">
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
            chrome="embedded"
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
    </>
  );
}
