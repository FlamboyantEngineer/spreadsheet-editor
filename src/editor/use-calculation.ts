'use client';

import { useEffect, useMemo, useState } from 'react';
import type { Workbook } from '../model/workbook.js';
import { assertCalculationResult } from '../calculation/http.js';
import type { CalculationOptions, CalculationResult } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import { workbookCalculationKey } from '../calculation/publication.js';

const acceptedKeys = new WeakMap<Workbook, string>();
/**
 * The key of a workbook the editor accepted. The editor never changes an accepted workbook, so
 * its key is computed once and shared by every reader (SE-3).
 */
export function acceptedCalculationKey(workbook: Workbook): string {
  let key = acceptedKeys.get(workbook);
  if (key === undefined) acceptedKeys.set(workbook, (key = workbookCalculationKey(workbook)));
  return key;
}

export function useCalculation(
  workbook: Workbook,
  provider: CalculationProvider,
  options?: CalculationOptions,
  enabled = true,
  /** The built-in client calculator's own results are not re-validated (SE-PERF-1). */
  trusted = false,
) {
  const inputKey = useMemo(() => acceptedCalculationKey(workbook), [workbook]);
  const request = useMemo(
    () => ({ workbook, provider, options, enabled, trusted }),
    [inputKey, provider, options, enabled, trusted],
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
        if (!request.trusted) assertCalculationResult(result, request.workbook);
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
