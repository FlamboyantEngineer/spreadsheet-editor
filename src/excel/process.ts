import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import type { CalculationResult } from '../calculation/evaluator.js';
import { requireDefined } from '../model/require-defined.js';
import type { Workbook } from '../model/workbook.js';
import { validateFormulaCapabilities } from './capabilities.js';
import type { ExcelFormulaCapabilities, XlsxExport, XlsxImport } from './contracts.js';
import {
  decodeExportResponse,
  encodeImportRequest,
  excelInputLimit,
  type ExcelOperation,
  excelOutputLimit,
} from './protocol.js';

export interface ExcelProcessOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
  formulaCapabilities?: ExcelFormulaCapabilities;
}
let exportsInFlight = 0;
/** Export the captured workbook and results in a bounded child, without evaluation or live data reads. */
export async function exportXlsx(
  workbook: Workbook,
  result: CalculationResult,
  options: ExcelProcessOptions = {},
): Promise<Uint8Array> {
  return (await exportXlsxWithReport(workbook, result, options)).bytes;
}
/** Node-only conversion boundary: 640 MiB heap, 768 MiB RSS, two minutes and 100 MiB output. */
export async function importXlsx(
  bytes: Uint8Array,
  options: ExcelProcessOptions = {},
): Promise<XlsxImport> {
  return JSON.parse(
    (
      await runChild('import', encodeImportRequest(bytes, options.formulaCapabilities), options)
    ).toString(),
  ) as XlsxImport;
}
/** Same immutable export with a grouped conversion-loss report. */
export async function exportXlsxWithReport(
  workbook: Workbook,
  result: CalculationResult,
  options: ExcelProcessOptions = {},
): Promise<XlsxExport> {
  if (exportsInFlight >= 2)
    throw Object.assign(new Error('Two Excel exports are already running. Try again.'), {
      status: 503,
    });
  exportsInFlight++;
  try {
    const request = Buffer.from(
      JSON.stringify({
        workbook,
        result,
        formulaCapabilities: validateFormulaCapabilities(options.formulaCapabilities),
      }),
    );
    return decodeExportResponse(await runChild('export', request, options));
  } finally {
    exportsInFlight--;
  }
}
function runChild(
  operation: ExcelOperation,
  bytes: Uint8Array,
  options: ExcelProcessOptions,
): Promise<Buffer> {
  if (bytes.byteLength > excelInputLimit(operation))
    return Promise.reject(new RangeError('Excel process input exceeds its byte limit.'));
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [
        '--max-old-space-size=640',
        fileURLToPath(new URL('./worker.js', import.meta.url)),
        operation,
      ],
      { stdio: ['pipe', 'pipe', 'pipe'], signal: options.signal },
    );
    const output: Buffer[] = [];
    let size = 0,
      error = '',
      settled = false;
    const finish = (failure?: Error, result?: Buffer) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (failure) reject(failure);
      else resolve(requireDefined(result));
    };
    const timer = setTimeout(
      () => {
        child.kill();
        finish(new Error('Excel processing exceeded two minutes.'));
      },
      Math.min(120000, Math.max(1, options.timeoutMs ?? 120000)),
    );
    child.stdout.on('data', (chunk: Buffer) => {
      if (settled) return;
      size += chunk.length;
      if (size > excelOutputLimit) {
        child.kill();
        finish(new Error('Excel process output exceeds 100 MiB.'));
      } else output.push(chunk);
    });
    child.stderr.on('data', (chunk: Buffer) => {
      if (error.length < 8192) error += chunk.toString().slice(0, 8192 - error.length);
    });
    child.on('error', (failure) => finish(failure));
    child.stdin.on('error', (failure) => {
      if ((failure as NodeJS.ErrnoException).code !== 'EPIPE') finish(failure);
    });
    child.on('close', (code) => {
      if (settled) return;
      if (code !== 0)
        finish(new Error(error || 'Excel processing failed within its resource limits.'));
      else finish(undefined, Buffer.concat(output, size));
    });
    child.stdin.end(bytes);
  });
}
