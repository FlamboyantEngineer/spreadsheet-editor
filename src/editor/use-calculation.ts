'use client';

import { useEffect, useMemo, useState } from 'react';
import type { Workbook } from '../model/workbook.js';
import { assertCalculationResult } from '../calculation/http.js';
import type { CalculationOptions, CalculationResult } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import { workbookCalculationKey } from '../calculation/publication.js';

export function useCalculation(
  workbook: Workbook,
  provider: CalculationProvider,
  options?: CalculationOptions,
  enabled = true,
) {
  const inputKey = useMemo(() => workbookCalculationKey(workbook), [workbook]);
  const request = useMemo(
    () => ({ workbook, provider, options, enabled }),
    [inputKey, provider, options, enabled],
  );
  const [settled, setSettled] = useState<{
    request: typeof request;
    result: CalculationResult | null;
    error: string | null;
  } | null>(null);
  useEffect(() => {
    if (!enabled) return;
    const controller = new AbortController();
    Promise.resolve()
      .then(() => request.provider.calculate(request.workbook, request.options, controller.signal))
      .then((result) => {
        if (controller.signal.aborted) return;
        assertCalculationResult(result, request.workbook);
        setSettled({ request, result, error: null });
      })
      .catch((error: unknown) => {
        if (!controller.signal.aborted)
          setSettled({
            request,
            result: null,
            error: error instanceof Error ? error.message : 'Calculation failed.',
          });
      });
    return () => controller.abort();
  }, [request]);
  const current = settled?.request === request;
  return {
    result: enabled && current ? settled.result : null,
    error: enabled && current ? settled.error : null,
    pending: enabled && !current,
  };
}
