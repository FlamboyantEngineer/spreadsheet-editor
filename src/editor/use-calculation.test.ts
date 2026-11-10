// @vitest-environment jsdom
import { act, renderHook, waitFor } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../model/workbook.js';
import { calculateWorkbook, type CalculationResult } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import { useCalculation } from './use-calculation.js';

describe('SE-UI-6: calculation lifecycle', () => {
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
});
