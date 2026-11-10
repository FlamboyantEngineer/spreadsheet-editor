'use client';

import { useEffect, useState } from 'react';
import type { Workbook } from '../model/workbook.js';
import { assertCalculationResult } from '../calculation/http.js';
import type { CalculationOptions, CalculationResult } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';

export function useCalculation(
  workbook: Workbook,
  provider: CalculationProvider,
  options?: CalculationOptions,
) {
  const [settled, setSettled] = useState<{
    workbook: Workbook;
    provider: CalculationProvider;
    options: CalculationOptions | undefined;
    result: CalculationResult | null;
    error: string | null;
  } | null>(null);
  useEffect(() => {
    const controller = new AbortController();
    Promise.resolve()
      .then(() => provider.calculate(workbook, options, controller.signal))
      .then((result) => {
        if (controller.signal.aborted) return;
        assertCalculationResult(result, workbook);
        setSettled({ workbook, provider, options, result, error: null });
      })
      .catch((error: unknown) => {
        if (!controller.signal.aborted)
          setSettled({
            workbook,
            provider,
            options,
            result: null,
            error: error instanceof Error ? error.message : 'Calculation failed.',
          });
      });
    return () => controller.abort();
  }, [workbook, provider, options]);
  const current =
    settled?.workbook === workbook && settled.provider === provider && settled.options === options;
  return {
    result: current ? settled.result : null,
    error: current ? settled.error : null,
    pending: !current,
  };
}
