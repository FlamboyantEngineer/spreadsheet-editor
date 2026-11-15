import { assertWorkbook, DEFAULT_WORKBOOK_LIMITS, type Workbook } from '../model/workbook.js';
import {
  calculateWorkbook,
  type CalculationOptions,
  type FormulaHost,
} from './evaluator.js';
import type { CalculationProvider } from './provider.js';
import { assertCalculationResult } from './result-validation.js';

export { assertCalculationResult } from './result-validation.js';

interface CalculationRequest {
  schemaVersion: 1;
  requestId: string;
  workbook: Workbook;
  options?: CalculationOptions;
}

/**
 * Default transport limits follow the cell limit: a request may use 64 bytes per cell and a response
 * 256 bytes per cell, so a workbook the cell limit accepts is not refused for its size.
 */
const REQUEST_BYTES_PER_CELL = 64,
  RESPONSE_BYTES_PER_CELL = 256;

class BodyTooLargeError extends RangeError {}

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
        throw new BodyTooLargeError('JSON body exceeds the byte limit.');
      }
      text += decoder.decode(value, { stream: true });
    }
    return JSON.parse(text + decoder.decode()) as unknown;
  } finally {
    reader.releaseLock();
  }
}

export interface HttpCalculatorOptions {
  fetch?: typeof globalThis.fetch;
  headers?: HeadersInit;
  credentials?: RequestCredentials;
  /** Defaults to 256 bytes per cell of the requested cell limit (64 MB for 250,000 cells). */
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
      const body = await readJson(
        response.body,
        options.maxResponseBytes ??
          (calculationOptions?.maxCells ?? DEFAULT_WORKBOOK_LIMITS.maxCells) *
            RESPONSE_BYTES_PER_CELL,
      );
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
  /** Defaults to 64 bytes per cell of the server's cell limit (16 MB for 250,000 cells). */
  maxRequestBytes?: number;
}

type NumericLimit = keyof Omit<CalculationOptions, 'timeZone' | 'wordSizeBits'>;
const LIMIT_DEFAULTS: Required<Pick<CalculationOptions, NumericLimit>> = {
  ...DEFAULT_WORKBOOK_LIMITS,
  maxOperations: 2_000_000,
  maxDepth: 128,
  maxIntegerDigits: 10_000,
  precision: 80,
};

function boundedLimit(key: NumericLimit, value: unknown, ceilings: CalculationOptions): number {
  if (
    value !== undefined &&
    (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1)
  )
    throw new TypeError('Calculation options must be positive integers.');
  // Approximation precision is a requested value, bounded by the server's maximum.
  const ceiling = ceilings[key] ?? (key === 'precision' ? 1000 : LIMIT_DEFAULTS[key]);
  return Math.min(typeof value === 'number' ? value : LIMIT_DEFAULTS[key], ceiling);
}

function calendarContext(
  options: Record<string, unknown> | undefined,
  ceilings: CalculationOptions,
): Pick<CalculationOptions, 'timeZone' | 'wordSizeBits'> {
  const timeZone = ceilings.timeZone ?? options?.timeZone ?? 'UTC';
  if (typeof timeZone !== 'string' || timeZone.length > 100)
    throw new TypeError('Invalid calendar time zone.');
  new Intl.DateTimeFormat('en-US', { timeZone });
  const wordSizeBits = ceilings.wordSizeBits ?? options?.wordSizeBits;
  if (
    wordSizeBits !== undefined &&
    (typeof wordSizeBits !== 'string' ||
      !/^[1-9]\d{0,5}$/.test(wordSizeBits) ||
      BigInt(wordSizeBits) > 100000n)
  )
    throw new TypeError('Invalid wordSizeBits context.');
  return { timeZone, ...(wordSizeBits === undefined ? {} : { wordSizeBits }) };
}

function boundedOptions(requested: unknown, ceilings: CalculationOptions = {}): CalculationOptions {
  if (
    requested !== undefined &&
    (!requested || typeof requested !== 'object' || Array.isArray(requested))
  )
    throw new TypeError('Invalid calculation options.');
  const options = requested as Record<string, unknown> | undefined;
  const result: Record<string, number> = {};
  for (const key of Object.keys(LIMIT_DEFAULTS) as NumericLimit[])
    result[key] = boundedLimit(key, options?.[key], ceilings);
  return { ...result, ...calendarContext(options, ceilings) };
}

/** The request's workbook and bounded options, or a TypeError or RangeError. */
async function readCalculationRequest(request: Request, configuration: CalculationHandlerOptions) {
  const value = await readJson(
    request.body,
    configuration.maxRequestBytes ??
      (configuration.limits?.maxCells ?? DEFAULT_WORKBOOK_LIMITS.maxCells) * REQUEST_BYTES_PER_CELL,
  );
  if (!value || typeof value !== 'object') throw new TypeError('Invalid calculation request.');
  const body = value as Partial<CalculationRequest>;
  if (body.schemaVersion !== 1 || typeof body.requestId !== 'string' || body.requestId.length > 128)
    throw new TypeError('Invalid calculation request version or id.');
  const options = boundedOptions(body.options, configuration.limits);
  assertWorkbook(body.workbook, options);
  return { requestId: body.requestId, workbook: body.workbook, options };
}

function refusal(error: unknown): Response {
  if (!(error instanceof RangeError))
    return Response.json({ error: 'Invalid calculation request.' }, { status: 400 });
  // Say which limit was hit: the request's byte size, or a workbook or calculation limit.
  return Response.json(
    {
      error: 'Calculation request exceeds a configured limit.',
      limit: error instanceof BodyTooLargeError ? 'requestBytes' : 'workbook',
      detail: error.message,
    },
    { status: 413 },
  );
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
    const { requestId, workbook, options } = await readCalculationRequest(request, configuration);
    const result = calculateWorkbook(workbook, options, configuration.host);
    return Response.json({ requestId, result }, { headers: { 'cache-control': 'no-store' } });
  } catch (error) {
    return refusal(error);
  }
}
