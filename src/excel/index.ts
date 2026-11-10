import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import type { Workbook } from '../model/workbook.js';
import type { CalculationResult } from '../calculation/evaluator.js';
import type { XlsxImport } from './contracts.js';
export type { XlsxImport, XlsxWarning, XlsxReport } from './contracts.js';

export interface ExcelProcessOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
}
let exportsInFlight = 0;
/** Export the captured workbook and results in a bounded child, without evaluation or live data reads. */
export async function exportXlsx(
  workbook: Workbook,
  result: CalculationResult,
  options: ExcelProcessOptions = {},
): Promise<Uint8Array> {
  if (exportsInFlight >= 2)
    throw Object.assign(new Error('Two Excel exports are already running. Try again.'), {
      status: 503,
    });
  exportsInFlight++;
  try {
    return new Uint8Array(
      await runChild('export', Buffer.from(JSON.stringify({ workbook, result })), options),
    );
  } finally {
    exportsInFlight--;
  }
}
/** Node-only conversion boundary: 640 MiB heap, 768 MiB RSS, two minutes and 100 MiB output. */
export async function importXlsx(
  bytes: Uint8Array,
  options: ExcelProcessOptions = {},
): Promise<XlsxImport> {
  return JSON.parse((await runChild('import', bytes, options)).toString()) as XlsxImport;
}
function runChild(
  operation: 'import' | 'export',
  bytes: Uint8Array,
  options: ExcelProcessOptions,
): Promise<Buffer> {
  if (bytes.byteLength > (operation === 'import' ? 20 : 100) * 1024 * 1024)
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
      else resolve(result!);
    };
    const timer = setTimeout(
      () => {
        child.kill();
        finish(new Error('Excel processing exceeded two minutes.'));
      },
      Math.min(120000, Math.max(1, options.timeoutMs ?? 120000)),
    );
    child.stdout.on('data', (chunk: Buffer) => {
      size += chunk.length;
      if (size > 100 * 1024 * 1024) {
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
    child.on('close', (code) =>
      finish(
        code !== 0
          ? new Error(error || 'Excel processing failed within its resource limits.')
          : undefined,
        Buffer.concat(output),
      ),
    );
    child.stdin.end(bytes);
  });
}
