import { assertWorkbook, DEFAULT_WORKBOOK_LIMITS, type Workbook } from '../model/workbook.js';
import {
  calculateWorkbook,
  type CalculationOptions,
  type CalculationResult,
  type FormulaHost,
} from './evaluator.js';
import type { CalculationProvider } from './provider.js';
import {
  assertJsonValue,
  assertNumericValue,
  type CellValue,
  FORMULA_ERROR_CODES,
} from './values.js';

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
      for (const cell of row) validateCell(cell);
    });
  }
  if (result.context !== undefined) {
    const context = result.context;
    if (
      !context ||
      typeof context !== 'object' ||
      !context.operands ||
      typeof context.operands !== 'object' ||
      Array.isArray(context.operands) ||
      Object.keys(context.operands).length > 100000
    )
      throw new TypeError('Invalid calculation context.');
    if (context.hostContext !== undefined) assertJsonValue(context.hostContext);
    for (const operand of Object.values(context.operands)) {
      if (
        !operand ||
        typeof operand !== 'object' ||
        typeof operand.formula !== 'string' ||
        operand.formula.length > 200000 ||
        typeof operand.name !== 'string' ||
        operand.name.length > 256
      )
        throw new TypeError('Invalid captured operand.');
      validateCell(operand.value);
      if (operand.hostContext !== undefined) assertJsonValue(operand.hostContext);
    }
  }
  if (result.dependencies !== undefined) {
    const dependencies = result.dependencies;
    if (
      !dependencies ||
      typeof dependencies !== 'object' ||
      !dependencies.cells ||
      typeof dependencies.cells !== 'object' ||
      Array.isArray(dependencies.cells) ||
      (dependencies.scope !== undefined && dependencies.scope !== 'formulas')
    )
      throw new TypeError('Invalid calculation dependencies.');
    const sheets = new Set(workbook.worksheets.map((sheet) => sheet.id));
    const cellKey = (key: string) => {
      let tuple: unknown;
      try {
        tuple = JSON.parse(key);
      } catch {
        throw new TypeError('Invalid dependency cell key.');
      }
      if (
        !Array.isArray(tuple) ||
        tuple.length !== 3 ||
        typeof tuple[0] !== 'string' ||
        !sheets.has(tuple[0]) ||
        !Number.isSafeInteger(tuple[1]) ||
        !Number.isSafeInteger(tuple[2]) ||
        tuple[1] < 0 ||
        tuple[1] > 16383 ||
        tuple[2] < 0 ||
        tuple[2] > 1048575
      )
        throw new TypeError('Invalid dependency cell key.');
    };
    if (Object.keys(dependencies.cells).length > 250000)
      throw new RangeError('Too many calculation dependencies.');
    for (const [key, reads] of Object.entries(dependencies.cells)) {
      cellKey(key);
      if (!Array.isArray(reads) || reads.length > 10000)
        throw new TypeError('Invalid dependency reads.');
      reads.forEach((read) => {
        if (typeof read !== 'string') throw new TypeError('Invalid dependency cell key.');
        cellKey(read);
      });
    }
    if (dependencies.extents !== undefined) {
      if (
        !dependencies.extents ||
        typeof dependencies.extents !== 'object' ||
        Array.isArray(dependencies.extents)
      )
        throw new TypeError('Invalid extent dependencies.');
      for (const [key, extents] of Object.entries(dependencies.extents)) {
        cellKey(key);
        if (
          !Array.isArray(extents) ||
          extents.length > 200 ||
          extents.some((sheet) => typeof sheet !== 'string' || !sheets.has(sheet))
        )
          throw new TypeError('Invalid extent dependency.');
      }
    }
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

function boundedOptions(requested: unknown, ceilings: CalculationOptions = {}): CalculationOptions {
  if (
    requested !== undefined &&
    (!requested || typeof requested !== 'object' || Array.isArray(requested))
  )
    throw new TypeError('Invalid calculation options.');
  const options = requested as Record<string, unknown> | undefined;
  const result: Record<string, number> = {};
  const defaults: Required<Omit<CalculationOptions, 'timeZone' | 'wordSizeBits'>> = {
    ...DEFAULT_WORKBOOK_LIMITS,
    maxOperations: 2_000_000,
    maxDepth: 128,
    maxIntegerDigits: 10_000,
    precision: 80,
  };
  for (const key of Object.keys(defaults) as (keyof Omit<
    CalculationOptions,
    'timeZone' | 'wordSizeBits'
  >)[]) {
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
  return { ...result, timeZone, ...(wordSizeBits === undefined ? {} : { wordSizeBits }) };
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
    const value = await readJson(
      request.body,
      configuration.maxRequestBytes ??
        (configuration.limits?.maxCells ?? DEFAULT_WORKBOOK_LIMITS.maxCells) *
          REQUEST_BYTES_PER_CELL,
    );
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
}

function validateCell(value: CellValue): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (!value || typeof value !== 'object') throw new TypeError('Invalid calculated cell.');
  if (value.kind === 'error' && (FORMULA_ERROR_CODES as readonly string[]).includes(value.code))
    return;
  assertNumericValue(value);
}
