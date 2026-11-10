import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook } from './evaluator.js';
import { createHttpCalculator, handleCalculationRequest } from './http.js';

const request = (body: unknown) =>
  new Request('http://localhost/calculate', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(body),
  });

describe('SE-CALC-3: interchangeable server calculation', () => {
  it('round-trips the exact browser result through JSON and the server handler', async () => {
    const workbook = createWorkbook([['1E-25', '=1-A1', '=1-B1', '=1/3', '=SQRT(2)']]);
    const server = createHttpCalculator('http://localhost/calculate', {
      fetch: (async (url, init) =>
        handleCalculationRequest(new Request(url, init))) as typeof fetch,
    });
    expect(await server.calculate(workbook)).toEqual(calculateWorkbook(workbook));
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
});
