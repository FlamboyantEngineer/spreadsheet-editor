import { useEffect, useMemo, useRef, useState } from 'react';
import { Modal } from '@flamboyantengineer/ui-components/modal';
import { FileField } from '@flamboyantengineer/ui-components/file-field';
import { SpreadsheetEditor } from '../../src/index.js';
import type { Workbook } from '../../src/model/index.js';
import type { CalculationResult } from '../../src/calculation/index.js';
import type { XlsxImport } from '../../src/excel/contracts.js';

type Prepared = XlsxImport & { result: CalculationResult };
export function useExcelLab() {
  const [open, setOpen] = useState(false),
    [pending, setPending] = useState(false),
    [prepared, setPrepared] = useState<Prepared | null>(null),
    [error, setError] = useState<string | null>(null),
    [download, setDownload] = useState<string | null>(null);
  const request = useRef<((workbook: Workbook | null) => void) | null>(null),
    controller = useRef<AbortController | null>(null);
  useEffect(
    () => () => {
      controller.current?.abort();
      request.current?.(null);
    },
    [],
  );
  useEffect(
    () => () => {
      if (download) URL.revokeObjectURL(download);
    },
    [download],
  );
  const calculator = useMemo(() => ({ calculate: async () => prepared!.result }), [prepared]);
  function close(workbook: Workbook | null = null) {
    controller.current?.abort();
    request.current?.(workbook);
    request.current = null;
    setOpen(false);
    setPending(false);
  }
  async function select(files: File[]) {
    const file = files[0];
    if (!file) return;
    controller.current?.abort();
    const current = new AbortController();
    controller.current = current;
    setPending(true);
    setError(null);
    setPrepared(null);
    try {
      const response = await fetch('/api/xlsx/import', {
        method: 'POST',
        body: file,
        signal: current.signal,
      });
      const result = (await response.json()) as Prepared & { error?: string };
      if (!response.ok) throw new Error(result.error ?? 'Excel import failed.');
      if (!current.signal.aborted) setPrepared(result);
    } catch (error) {
      if (!current.signal.aborted)
        setError(error instanceof Error ? error.message : 'Excel import failed.');
    } finally {
      if (!current.signal.aborted) setPending(false);
    }
  }
  return {
    requestImport: () =>
      new Promise<Workbook | null>((resolve) => {
        request.current?.(null);
        request.current = resolve;
        setOpen(true);
        setPrepared(null);
        setError(null);
      }),
    exportWorkbook: async (workbook: Workbook, result: CalculationResult) => {
      const response = await fetch('/api/xlsx/export', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ workbook, result }),
      });
      if (!response.ok) {
        const body = (await response.json()) as { error: string };
        throw new Error(body.error);
      }
      setDownload(URL.createObjectURL(await response.blob()));
    },
    download: download && (
      <p className="lab-download" role="status">
        Excel export is ready.{' '}
        <a href={download} download="workbook.xlsx">
          Download workbook.xlsx
        </a>
      </p>
    ),
    dialog: (
      <Modal
        title="Import Excel workbook"
        size="lg"
        isOpen={open}
        onOpenChange={(next) => {
          if (!next) close();
        }}
        primaryAction={{
          label: 'Replace workbook',
          isDisabled: !prepared || pending,
          onPress: () => close(prepared?.workbook ?? null),
        }}
        cancelAction={{}}
      >
        <p>
          Choose one .xlsx file, review the calculated preview and conversion report, then replace
          the lab workbook.
        </p>
        <FileField
          label="Excel workbook"
          acceptedFileTypes={['.xlsx']}
          maxSize={20 * 1024 * 1024}
          isDisabled={pending}
          onSelect={(files) => void select(files)}
        />
        {pending && <p role="status">Preparing workbook…</p>}
        {error && <p role="alert">{error}</p>}
        {prepared && (
          <>
            <p>
              {prepared.report.worksheets} worksheets · {prepared.report.cells} cells ·{' '}
              {prepared.report.formulas} formulas · {prepared.report.mergedRanges} merged ranges
            </p>
            <SpreadsheetEditor
              aria-label="Excel import preview"
              chrome="embedded"
              value={prepared.workbook}
              calculation={{ mode: 'server', provider: calculator }}
              readOnly
              features={{
                comments: false,
                suggestions: false,
                search: false,
                panels: false,
                modeControl: false,
              }}
            />
            <h3>Conversion report</h3>
            {prepared.report.warnings.length ? (
              <ul>
                {prepared.report.warnings.map((warning) => (
                  <li key={warning.code}>
                    <strong>
                      {warning.disposition === 'skipped' ? 'Skipped' : 'Not imported'}:
                    </strong>{' '}
                    {warning.message} ({warning.count}) {warning.locations.join(', ')}
                    {warning.count > warning.locations.length ? ', …' : ''}
                  </li>
                ))}
              </ul>
            ) : (
              <p>No conversion warnings.</p>
            )}
          </>
        )}
      </Modal>
    ),
  };
}
