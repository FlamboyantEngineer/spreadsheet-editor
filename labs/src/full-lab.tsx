import { useEffect, useMemo, useState } from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { SpreadsheetEditor, useWorkbookSession } from '../../src/index.js';
import {
  createMemoryHost,
  createMemoryPresenceRoom,
  FULL_CAPABILITIES,
  READ_ONLY_CAPABILITIES,
} from '../../src/host/index.js';
import {
  createHttpCalculator,
  createClientCalculator,
  calculateWorkbook,
  formatCellValue,
} from '../../src/calculation/index.js';
import { createDocumentInsertion, type DocumentInsertion } from '../../src/document/index.js';
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

const alex = { id: 'alex', name: 'Alex' };
const morgan = { id: 'morgan', name: 'Morgan' };
const LAB_WORKBOOK = 'lab-workbook';

export function FullLab() {
  const excel = useExcelLab();
  const [workbook, setWorkbook] = useState(sample);
  const [reviewerOnly, setReviewerOnly] = useState(false);
  const [presence, setPresence] = useState(false);
  const [autosave, setAutosave] = useState(false);
  const [actorId, setActorId] = useState('alex');
  const [readOnly, setReadOnly] = useState(false);
  const hosts = useMemo(() => {
    const room = createMemoryPresenceRoom();
    return {
      alex: createMemoryHost({
        reader: alex,
        people: [morgan],
        presence: room,
        workbooks: { [LAB_WORKBOOK]: sample },
      }),
      morgan: createMemoryHost({ reader: morgan, people: [alex], presence: room }),
    };
  }, []);
  useEffect(() => {
    if (!presence) return;
    const place = { sheetId: 'calculations', range: 'B2' };
    hosts.morgan.presence?.publish(LAB_WORKBOOK, { at: place, editLock: place, status: 'active' });
    return () => hosts.morgan.presence?.publish(LAB_WORKBOOK, null);
  }, [presence, hosts]);
  // Both readers save to the first host's store, as two people would to one server.
  const store = hosts.alex;
  const session = useWorkbookSession({
    workbooks: store.workbooks,
    drafts: store.drafts,
    workbookId: LAB_WORKBOOK,
    autosaveDelay: autosave ? 1000 : Infinity,
  });
  /** Replaces the editor's workbook with `next` as an unsaved change, as an import would. */
  function replaceWorkbook(next: Workbook) {
    if (!session?.workbook()) return;
    session.change(structuredClone(next), {
      transactionId: crypto.randomUUID(),
      source: 'host-command',
      baseline: session.workbook()!,
      operations: [],
    });
    setWorkbook(next);
  }
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
          session={session}
          versions={store.versions}
          capabilities={
            readOnly
              ? READ_ONLY_CAPABILITIES
              : reviewerOnly
                ? { ...FULL_CAPABILITIES, edit: false }
                : FULL_CAPABILITIES
          }
          workbookId={LAB_WORKBOOK}
          presence={hosts[actorId === 'alex' ? 'alex' : 'morgan'].presence}
          getCommentLink={(id) => new URL(`?comment=${encodeURIComponent(id)}`, location.href).href}
          initialCommentId={new URLSearchParams(location.search).get('comment') ?? undefined}
          onChange={setWorkbook}
          calculation={selectedMode}
          display={display}
          onInsert={setInsertion}
          identity={hosts[actorId === 'alex' ? 'alex' : 'morgan'].identity}
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
          <Button
            isDisabled={!session}
            onPress={() =>
              void session?.takeStored().then(() => setWorkbook(session.workbook() ?? sample))
            }
          >
            Reload saved workbook
          </Button>
          <Button
            onPress={() => {
              replaceWorkbook(sample);
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
