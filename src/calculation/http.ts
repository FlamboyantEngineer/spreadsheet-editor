import { assertWorkbook, type Workbook } from '../model/workbook.js';
import {
  calculateWorkbook,
  type CalculationOptions,
  type CalculationResult,
  type FormulaHost,
} from './evaluator.js';
import type { CalculationProvider } from './provider.js';

interface CalculationRequest {
  schemaVersion: 1;
  requestId: string;
  workbook: Workbook;
  options?: CalculationOptions;
}

async function readJson(
  body: ReadableStream<Uint8Array> | null,
  maxBytes: number,
): Promise<unknown> {
  if (!body) throw new TypeError('Missing JSON body.');
  const reader = body.getReader(),
    decoder = new TextDecoder();
  let size = 0,
    text = '';
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > maxBytes) {
        await reader.cancel();
        throw new RangeError('JSON body exceeds the byte limit.');
      }
      text += decoder.decode(value, { stream: true });
    }
    return JSON.parse(text + decoder.decode()) as unknown;
  } finally {
    reader.releaseLock();
  }
}

export function assertCalculationResult(
  value: unknown,
  workbook: Workbook,
): asserts value is CalculationResult {
  if (!value || typeof value !== 'object') throw new TypeError('Invalid calculation response.');
  const result = value as CalculationResult;
  if (
    result.schemaVersion !== 1 ||
    !Array.isArray(result.worksheets) ||
    result.worksheets.length !== workbook.worksheets.length
  )
    throw new TypeError('Invalid result worksheets.');
  for (let index = 0; index < workbook.worksheets.length; index++) {
    const sheet = workbook.worksheets[index],
      output = result.worksheets[index];
    if (
      output?.id !== sheet.id ||
      !Array.isArray(output.values) ||
      output.values.length !== sheet.data.length
    )
      throw new TypeError('Result does not match the workbook.');
    output.values.forEach((row, y) => {
      if (!Array.isArray(row) || row.length !== sheet.data[y].length)
        throw new TypeError('Result does not match the row.');
      for (const cell of row) {
        if (cell === null || typeof cell === 'string' || typeof cell === 'boolean') continue;
        if (!cell || typeof cell !== 'object') throw new TypeError('Invalid calculated cell.');
        if (
          cell.kind === 'error' &&
          [
            '#DIV/0!',
            '#VALUE!',
            '#REF!',
            '#NAME?',
            '#NUM!',
            '#N/A',
            '#ERROR!',
            '#LIMIT!',
            '#CYCLE!',
          ].includes(cell.code)
        )
          continue;
        if (
          cell.kind !== 'number' ||
          typeof cell.numerator !== 'string' ||
          typeof cell.denominator !== 'string' ||
          cell.numerator.length > 100_001 ||
          cell.denominator.length > 100_000 ||
          !/^-?(?:0|[1-9]\d*)$/.test(cell.numerator) ||
          !/^[1-9]\d*$/.test(cell.denominator) ||
          typeof cell.approximate !== 'boolean' ||
          (cell.approximate &&
            (!Number.isInteger(cell.precision) || cell.precision! < 16 || cell.precision! > 1000))
        )
          throw new TypeError('Invalid numeric result.');
      }
    });
  }
}

export interface HttpCalculatorOptions {
  fetch?: typeof globalThis.fetch;
  headers?: HeadersInit;
  credentials?: RequestCredentials;
  maxResponseBytes?: number;
}

export function createHttpCalculator(
  endpoint: string | URL,
  options: HttpCalculatorOptions = {},
): CalculationProvider {
  let sequence = 0;
  return {
    async calculate(workbook, calculationOptions, signal) {
      assertWorkbook(workbook, calculationOptions);
      const requestId = String(++sequence);
      const headers = new Headers(options.headers);
      headers.set('content-type', 'application/json');
      const response = await (options.fetch ?? globalThis.fetch)(endpoint, {
        method: 'POST',
        headers,
        credentials: options.credentials ?? 'same-origin',
        signal,
        body: JSON.stringify({
          schemaVersion: 1,
          requestId,
          workbook,
          options: calculationOptions,
        } satisfies CalculationRequest),
      });
      if (!response.ok) throw new Error(`Calculation server returned HTTP ${response.status}.`);
      const body = await readJson(response.body, options.maxResponseBytes ?? 16 * 1024 * 1024);
      if (
        !body ||
        typeof body !== 'object' ||
        !('requestId' in body) ||
        body.requestId !== requestId ||
        !('result' in body)
      )
        throw new TypeError('Calculation response has an invalid request id.');
      assertCalculationResult(body.result, workbook);
      signal?.throwIfAborted();
      return body.result;
    },
  };
}

export interface CalculationHandlerOptions {
  host?: FormulaHost;
  /** Server ceilings take precedence over caller-requested limits. */
  limits?: CalculationOptions;
  maxRequestBytes?: number;
}

function boundedOptions(requested: unknown, ceilings: CalculationOptions = {}): CalculationOptions {
  if (
    requested !== undefined &&
    (!requested || typeof requested !== 'object' || Array.isArray(requested))
  )
    throw new TypeError('Invalid calculation options.');
  const options = requested as Record<string, unknown> | undefined;
  const result: Record<string, number> = {};
  const defaults: Required<CalculationOptions> = {
    maxCells: 50_000,
    maxCellLength: 10_000,
    maxWorksheets: 100,
    maxOperations: 200_000,
    maxDepth: 128,
    maxIntegerDigits: 10_000,
    precision: 80,
  };
  for (const key of Object.keys(defaults) as (keyof CalculationOptions)[]) {
    const value = options?.[key];
    if (
      value !== undefined &&
      (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1)
    )
      throw new TypeError('Calculation options must be positive integers.');
    // Approximation precision is a requested value, bounded by the server's maximum.
    const ceiling = ceilings[key] ?? (key === 'precision' ? 1000 : defaults[key]);
    result[key] = Math.min(typeof value === 'number' ? value : defaults[key], ceiling);
  }
  return result;
}

/** Fetch-standard route handler: mount behind the host's own routing and authorization. */
export async function handleCalculationRequest(
  request: Request,
  configuration: CalculationHandlerOptions = {},
): Promise<Response> {
  if (request.method !== 'POST')
    return Response.json({ error: 'Use POST.' }, { status: 405, headers: { allow: 'POST' } });
  if (!request.headers.get('content-type')?.toLowerCase().startsWith('application/json'))
    return Response.json({ error: 'Expected application/json.' }, { status: 415 });
  try {
    const value = await readJson(request.body, configuration.maxRequestBytes ?? 2 * 1024 * 1024);
    if (!value || typeof value !== 'object') throw new TypeError('Invalid calculation request.');
    const body = value as Partial<CalculationRequest>;
    if (
      body.schemaVersion !== 1 ||
      typeof body.requestId !== 'string' ||
      body.requestId.length > 128
    )
      throw new TypeError('Invalid calculation request version or id.');
    const options = boundedOptions(body.options, configuration.limits);
    assertWorkbook(body.workbook, options);
    const result = calculateWorkbook(body.workbook, options, configuration.host);
    return Response.json(
      { requestId: body.requestId, result },
      { headers: { 'cache-control': 'no-store' } },
    );
  } catch (error) {
    return Response.json(
      {
        error:
          error instanceof RangeError
            ? 'Calculation request exceeds a configured limit.'
            : 'Invalid calculation request.',
      },
      { status: error instanceof RangeError ? 413 : 400 },
    );
  }
}
