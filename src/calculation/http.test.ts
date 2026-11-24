import { describe, expect, it } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook } from './evaluator.js';
import { assertCalculationResult, createHttpCalculator, handleCalculationRequest } from './http.js';

const request = (body: unknown) =>
  new Request('http://localhost/calculate', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(body),
  });

describe('SE-CALC-3.1: interchangeable server calculation', () => {
  it('round-trips the exact browser result through JSON and the server handler', async () => {
    const workbook = createWorkbook([['1E-25', '=1-A1', '=1-B1', '=1/3', '=SQRT(2)']]);
    const server = createHttpCalculator('http://localhost/calculate', {
      fetch: (async (url, init) =>
        handleCalculationRequest(new Request(url, init))) as typeof fetch,
    });
    expect(await server.calculate(workbook)).toEqual(calculateWorkbook(workbook));
  });
  it('keeps typed errors, error-like text and blank results local to their cells', async () => {
    const workbook = createWorkbook([
      ['=XFE1', '=SUM(A1048576:A1048577)', '="#N/A label"', '=IF(FALSE,1,)'],
    ]);
    const provider = createHttpCalculator('http://localhost/calculate', {
      fetch: (async (url, init) =>
        handleCalculationRequest(new Request(url, init))) as typeof fetch,
    });
    const result = await provider.calculate(workbook);
    const invalidReference = { kind: 'error', code: '#REF!' };
    expect(result).toEqual(calculateWorkbook(workbook));
    expect(result.worksheets[0].values[0]).toEqual([
      invalidReference,
      invalidReference,
      '#N/A label',
      null,
    ]);
  });
  it('rejects a differently serialized key that would skip dependency invalidation', () => {
    const workbook = createWorkbook([['1', '=A1']]);
    const result = calculateWorkbook(workbook);
    requireDefined(result.dependencies).cells['["sheet-1",1,0]'] = ['["sheet-1", 0, 0]'];
    expect(() => assertCalculationResult(result, workbook)).toThrow('Invalid dependency cell key');
  });
  it('forwards authentication headers and cancellation without local fallback', async () => {
    const controller = new AbortController();
    const provider = createHttpCalculator('http://localhost/calculate', {
      headers: { authorization: 'Bearer example' },
      fetch: (async (_url, init) => {
        expect(new Headers(init?.headers).get('authorization')).toBe('Bearer example');
        expect(init?.signal).toBe(controller.signal);
        return new Response('', { status: 503 });
      }) as typeof fetch,
    });
    await expect(provider.calculate(createWorkbook(), {}, controller.signal)).rejects.toThrow(
      '503',
    );
  });
  it('rejects malformed and mismatched numeric responses', async () => {
    const provider = createHttpCalculator('http://localhost/calculate', {
      fetch: (async () => Response.json({ requestId: 'bad', result: {} })) as typeof fetch,
    });
    await expect(provider.calculate(createWorkbook())).rejects.toThrow('request id');
    const invalid = createHttpCalculator('http://localhost/calculate', {
      fetch: (async () =>
        Response.json({
          requestId: '1',
          result: {
            schemaVersion: 1,
            worksheets: [
              {
                id: 'sheet-1',
                values: [
                  [{ kind: 'number', numerator: '1', denominator: '0', approximate: false }],
                ],
              },
            ],
          },
        })) as typeof fetch,
    });
    await expect(invalid.calculate(createWorkbook())).rejects.toThrow('numeric');
  });
  it('checks method, content type, body limits, and workbook shape', async () => {
    expect((await handleCalculationRequest(new Request('http://localhost'))).status).toBe(405);
    expect(
      (
        await handleCalculationRequest(
          new Request('http://localhost', { method: 'POST', body: '{}' }),
        )
      ).status,
    ).toBe(415);
    expect(
      (await handleCalculationRequest(request({ schemaVersion: 1, requestId: '1', workbook: {} })))
        .status,
    ).toBe(400);
    expect(
      (await handleCalculationRequest(request({ many: 'x'.repeat(100) }), { maxRequestBytes: 32 }))
        .status,
    ).toBe(413);
  });
  it('enforces server limits even when the client requests more', async () => {
    const response = await handleCalculationRequest(
      request({
        schemaVersion: 1,
        requestId: '1',
        workbook: createWorkbook([['1', '2']]),
        options: { maxCells: 1000 },
      }),
      { limits: { maxCells: 1 } },
    );
    expect(response.status).toBe(413);
  });
  it('has space for JSON framing even with a one-cell limit', async () => {
    const limits = { maxCells: 1 };
    const provider = createHttpCalculator('http://localhost/calculate', {
      fetch: (async (url, init) =>
        handleCalculationRequest(new Request(url, init), { limits })) as typeof fetch,
    });
    const workbook = createWorkbook([['1']]);
    expect(await provider.calculate(workbook, limits)).toEqual(calculateWorkbook(workbook, limits));
  });
  it('fits a large workbook of short values in the default transport limits', async () => {
    // 200,000 short values: 2.6 MiB of request JSON and about 10 MiB of response JSON.
    const workbook = createWorkbook(
      Array.from({ length: 2_000 }, (_, y) =>
        Array.from({ length: 100 }, (_, x) => String(y * 100 + x)),
      ),
    );
    const provider = createHttpCalculator('http://localhost/calculate', {
      fetch: (async (url, init) =>
        handleCalculationRequest(new Request(url, init))) as typeof fetch,
    });
    const result = await provider.calculate(workbook);
    expect(result.worksheets[0].values[1_999][99]).toMatchObject({ numerator: '199999' });
  }, 60_000);
  it('says which limit a refused request hit', async () => {
    const tooLarge = await handleCalculationRequest(request({ many: 'x'.repeat(100) }), {
      maxRequestBytes: 32,
    });
    expect(tooLarge.status).toBe(413);
    expect(await tooLarge.json()).toMatchObject({ limit: 'requestBytes' });
    const tooManyCells = await handleCalculationRequest(
      request({ schemaVersion: 1, requestId: '1', workbook: createWorkbook([['1', '2']]) }),
      { limits: { maxCells: 1 }, maxRequestBytes: 1024 },
    );
    expect(tooManyCells.status).toBe(413);
    expect(await tooManyCells.json()).toMatchObject({
      limit: 'workbook',
      detail: 'Worksheet extent exceeds the cell limit.',
    });
  });
});
