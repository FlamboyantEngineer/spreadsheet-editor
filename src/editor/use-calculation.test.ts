// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it, vi } from 'vitest';
import { calculateWorkbook, type CalculationResult } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import { createWorkbook } from '../model/workbook.js';
import { useCalculation } from './use-calculation.js';

describe('SE-UI-2.1: calculation lifecycle', () => {
  it('retains pending and ready calculations through changes that do not affect authored calculation inputs', async () => {
    const workbook = createWorkbook([['=1/3']]);
    let finish!: (result: CalculationResult) => void;
    const calculate = vi.fn<CalculationProvider['calculate']>(
      () =>
        new Promise((resolve) => {
          finish = resolve;
        }),
    );
    const provider = { calculate };
    const { result, rerender } = renderHook(({ value }) => useCalculation(value, provider), {
      initialProps: { value: workbook },
    });
    await waitFor(() => expect(calculate).toHaveBeenCalledTimes(1));
    const decorated = structuredClone(workbook);
    decorated.name = 'Renamed workbook';
    decorated.worksheets[0].style = { A1: 'font-weight:bold' };
    rerender({ value: decorated });
    expect(calculate.mock.calls[0][2]?.aborted).toBe(false);
    await act(async () => {
      finish(calculateWorkbook(workbook));
    });
    expect(result.current.result?.worksheets[0].values[0][0]).toMatchObject({
      numerator: '1',
      denominator: '3',
    });
    rerender({ value: structuredClone(decorated) });
    expect(result.current.pending).toBe(false);
    expect(calculate).toHaveBeenCalledTimes(1);
  });

  it('ignores an older response even if the transport ignores cancellation', async () => {
    const first = createWorkbook([['1']]),
      second = createWorkbook([['2']]);
    const pending: { resolve(value: CalculationResult): void; signal?: AbortSignal }[] = [];
    const provider: CalculationProvider = {
      calculate: (_workbook, _options, signal) =>
        new Promise((resolve) => pending.push({ resolve, signal })),
    };
    const { result, rerender, unmount } = renderHook(
      ({ workbook }) => useCalculation(workbook, provider),
      { initialProps: { workbook: first } },
    );
    await waitFor(() => expect(pending).toHaveLength(1));
    rerender({ workbook: second });
    expect(result.current.pending).toBe(true);
    await waitFor(() => expect(pending).toHaveLength(2));
    expect(pending[0].signal?.aborted).toBe(true);
    await act(async () => pending[1].resolve(calculateWorkbook(second)));
    expect(result.current.result).toEqual(calculateWorkbook(second));
    await act(async () => pending[0].resolve(calculateWorkbook(first)));
    expect(result.current.result).toEqual(calculateWorkbook(second));
    unmount();
    expect(pending[1].signal?.aborted).toBe(true);
  });

  it('starts a fresh pass instead of exposing old results when calculation is reenabled', async () => {
    const workbook = createWorkbook([['=NOW()']]);
    const replies: ((value: CalculationResult) => void)[] = [];
    const provider: CalculationProvider = {
      calculate: () => new Promise((resolve) => replies.push(resolve)),
    };
    const { result, rerender } = renderHook(
      ({ enabled }) => useCalculation(workbook, provider, { enabled }),
      { initialProps: { enabled: true } },
    );
    await waitFor(() => expect(replies).toHaveLength(1));
    await act(async () => {
      replies[0](calculateWorkbook(workbook));
    });
    expect(result.current.result).not.toBeNull();
    rerender({ enabled: false });
    expect(result.current.result).toBeNull();
    rerender({ enabled: true });
    expect(result.current.pending).toBe(true);
    expect(result.current.result).toBeNull();
    await waitFor(() => expect(replies).toHaveLength(2));
    await act(async () => {
      replies[1](calculateWorkbook(workbook));
    });
    expect(result.current.pending).toBe(false);
  });
});
