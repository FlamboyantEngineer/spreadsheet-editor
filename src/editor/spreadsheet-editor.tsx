'use client';

import { lazy, Suspense, useMemo, useState } from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { createClientCalculator, type CalculationProvider } from '../calculation/provider.js';
import { formatCellValue, type DisplayOptions } from '../calculation/format.js';
import type { CalculationOptions, FormulaHost } from '../calculation/evaluator.js';
import {
  createWorkbook,
  copyWorkbook,
  setWorkbookCell,
  appendWorksheet,
  type Workbook,
  type Worksheet,
} from '../model/workbook.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { createDocumentInsertion, type DocumentInsertion } from '../document/insertion.js';
import type { JSpreadsheetSelectionRange } from '../engine/jspreadsheet-host.js';
import { useCalculation } from './use-calculation.js';

const WorkbookGrid = lazy(() => import('../engine/workbook-grid.js'));

export type CalculationMode =
  | { mode: 'client'; host?: FormulaHost; options?: CalculationOptions }
  | { mode: 'server'; provider: CalculationProvider; options?: CalculationOptions };

export interface SpreadsheetEditorProps {
  value?: Workbook;
  defaultValue?: Workbook;
  onChange?(workbook: Workbook): void;
  calculation?: CalculationMode;
  display?: DisplayOptions;
  readOnly?: boolean;
  className?: string;
  'aria-label'?: string;
  /** Hide insertion commands when the embedding host does not supply this callback. */
  onInsert?(insertion: DocumentInsertion): void;
  initialSheetId?: string;
}

export function SpreadsheetEditor(props: SpreadsheetEditorProps) {
  const [local, setLocal] = useState(() => copyWorkbook(props.defaultValue ?? createWorkbook()));
  const workbook = props.value ?? local;
  const [history, setHistory] = useState<{
    head: Workbook | null;
    past: Workbook[];
    future: Workbook[];
  }>({ head: null, past: [], future: [] });
  const currentHistory =
    history.head === workbook ? history : { head: workbook, past: [], future: [] };
  const [sheetId, setSheetId] = useState(props.initialSheetId ?? workbook.worksheets[0].id);
  const sheet = workbook.worksheets.find((sheet) => sheet.id === sheetId) ?? workbook.worksheets[0];
  const [selection, setSelection] = useState<JSpreadsheetSelectionRange>({
    x1: 0,
    y1: 0,
    x2: 0,
    y2: 0,
  });
  const authored = sheet.data[selection.y1]?.[selection.x1] ?? '';
  const [draft, setDraft] = useState<{ key: string; value: string } | null>(null);
  const draftKey = `${sheet.id}:${selection.x1}:${selection.y1}:${String(authored)}`;
  const client = useMemo(
    () =>
      createClientCalculator(
        props.calculation?.mode === 'client' ? props.calculation.host : undefined,
      ),
    [props.calculation?.mode === 'client' ? props.calculation.host : undefined],
  );
  const provider = props.calculation?.mode === 'server' ? props.calculation.provider : client;
  const { result, error, pending } = useCalculation(workbook, provider, props.calculation?.options);
  const output = result?.worksheets.find((result) => result.id === sheet.id)?.values;
  function publish(next: Workbook) {
    if (!props.value) setLocal(next);
    props.onChange?.(next);
  }
  function commit(next: Workbook) {
    setHistory({
      head: next,
      past: [...currentHistory.past.slice(-49), copyWorkbook(workbook)],
      future: [],
    });
    publish(next);
  }
  function undo() {
    const previous = currentHistory.past.at(-1);
    if (!previous || props.readOnly) return;
    setHistory({
      head: previous,
      past: currentHistory.past.slice(0, -1),
      future: [copyWorkbook(workbook), ...currentHistory.future],
    });
    publish(previous);
  }
  function redo() {
    const next = currentHistory.future[0];
    if (!next || props.readOnly) return;
    setHistory({
      head: next,
      past: [...currentHistory.past, copyWorkbook(workbook)],
      future: currentHistory.future.slice(1),
    });
    publish(next);
  }
  function updateSheet(next: Worksheet) {
    if (!props.readOnly)
      commit({
        ...workbook,
        worksheets: workbook.worksheets.map((current) =>
          current.id === sheet.id ? next : current,
        ),
      });
  }
  const cellAddress = formatSpreadsheetAddress(selection.x1, selection.y1);
  const selectedRange = `${cellAddress}:${formatSpreadsheetAddress(selection.x2, selection.y2)}`;
  function mergeSelection() {
    const merged = { ...sheet.mergeCells };
    merged[cellAddress] = [selection.x2 - selection.x1 + 1, selection.y2 - selection.y1 + 1];
    updateSheet({ ...sheet, mergeCells: merged });
  }
  function unmergeSelection() {
    const merged = { ...sheet.mergeCells };
    delete merged[cellAddress];
    updateSheet({ ...sheet, mergeCells: merged });
  }
  const surfaceKey = JSON.stringify([
    sheet.id,
    props.readOnly,
    sheet.columns,
    sheet.rows,
    sheet.mergeCells,
    sheet.style,
  ]);
  function insert(mode: 'table' | 'spreadsheet') {
    props.onInsert?.(
      createDocumentInsertion(workbook, result, {
        mode,
        sheetId: sheet.id,
        range: selectedRange,
        display: props.display,
      }),
    );
  }
  function applyFormula() {
    if (!props.readOnly && draft?.key === draftKey) {
      commit(setWorkbookCell(workbook, sheet.id, selection.x1, selection.y1, draft.value));
      setDraft(null);
    }
  }
  return (
    <section
      className={['se-editor', props.className].filter(Boolean).join(' ')}
      aria-label={props['aria-label'] ?? 'Spreadsheet editor'}
      data-calculation-mode={props.calculation?.mode ?? 'client'}
      data-pending={pending}
    >
      {!props.readOnly && (
        <div className="se-editor__toolbar" role="group" aria-label="Spreadsheet commands">
          <Button onPress={undo} isDisabled={!currentHistory.past.length}>
            Undo
          </Button>
          <Button onPress={redo} isDisabled={!currentHistory.future.length}>
            Redo
          </Button>
          <Button
            onPress={() => {
              const next = appendWorksheet(workbook);
              commit(next);
              setSheetId(next.worksheets.at(-1)!.id);
            }}
          >
            Add worksheet
          </Button>
          <Button
            onPress={() =>
              updateSheet({
                ...sheet,
                data: [
                  ...sheet.data,
                  Array.from(
                    { length: Math.max(1, ...sheet.data.map((row) => row.length)) },
                    () => null,
                  ),
                ],
              })
            }
          >
            Add row
          </Button>
          <Button
            onPress={() => {
              const width = Math.max(0, ...sheet.data.map((row) => row.length));
              updateSheet({
                ...sheet,
                data: (sheet.data.length ? sheet.data : [[]]).map((row) => [
                  ...row,
                  ...Array.from({ length: width - row.length + 1 }, () => null),
                ]),
              });
            }}
          >
            Add column
          </Button>
          <Button
            onPress={mergeSelection}
            isDisabled={selection.x1 === selection.x2 && selection.y1 === selection.y2}
          >
            Merge cells
          </Button>
          <Button onPress={unmergeSelection} isDisabled={!sheet.mergeCells?.[cellAddress]}>
            Unmerge cells
          </Button>
        </div>
      )}
      <div className="se-editor__formula-bar">
        <span className="se-editor__address">{cellAddress}</span>
        <TextField
          label="Cell value or formula"
          value={draft?.key === draftKey ? draft.value : String(authored)}
          onChange={(value) => setDraft({ key: draftKey, value })}
          isReadOnly={props.readOnly}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              applyFormula();
            }
          }}
        />
        <Button onPress={applyFormula} isDisabled={props.readOnly || draft?.key !== draftKey}>
          Apply
        </Button>
        <output aria-label="Calculated cell value">
          {pending
            ? 'Calculating…'
            : error
              ? 'Unavailable'
              : formatCellValue(output?.[selection.y1]?.[selection.x1] ?? null, props.display)}
        </output>
      </div>
      <Suspense fallback={<p role="status">Loading spreadsheet…</p>}>
        <WorkbookGrid
          key={surfaceKey}
          worksheet={sheet}
          values={output}
          display={props.display}
          readOnly={props.readOnly}
          unavailableText={error ? 'Unavailable' : '…'}
          onChange={updateSheet}
          onSelectionChange={(range) => {
            setSelection(range);
            setDraft(null);
          }}
        />
      </Suspense>
      <div className="se-editor__sheets" role="group" aria-label="Worksheets">
        {workbook.worksheets.map((worksheet) => (
          <Button
            key={worksheet.id}
            isSelected={worksheet.id === sheet.id}
            onPress={() => {
              setSheetId(worksheet.id);
              setSelection({ x1: 0, y1: 0, x2: 0, y2: 0 });
              setDraft(null);
            }}
          >
            {worksheet.name}
          </Button>
        ))}
      </div>
      {props.onInsert && (
        <div className="se-editor__insertion" role="group" aria-label="Insert selected cells">
          <Button onPress={() => insert('table')} isDisabled={pending || !!error}>
            Insert as document table
          </Button>
          <Button onPress={() => insert('spreadsheet')}>Insert as live spreadsheet</Button>
        </div>
      )}
      <p className="se-editor__status" role={error ? 'alert' : 'status'}>
        {error ??
          (pending
            ? 'Calculating…'
            : `Calculated ${props.calculation?.mode === 'server' ? 'on server' : 'on client'}`)}
      </p>
    </section>
  );
}
