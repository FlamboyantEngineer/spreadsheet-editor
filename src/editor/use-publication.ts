'use client';
import { useMemo, useRef } from 'react';
import type { Workbook } from '../model/workbook.js';
import type { CalculationOptions } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import {
  assertExternalCalculationSnapshot,
  calculationCellKey,
  orderedJson,
  workbookCalculationKey,
} from '../calculation/publication.js';
import type { CalculationMode } from './contracts.js';
import type { CellPresentationSnapshot } from './public-host.js';
import { useCalculation } from './use-calculation.js';

export function usePublication(
  workbook: Workbook,
  mode: CalculationMode | undefined,
  provider: CalculationProvider,
  options: CalculationOptions,
  enabled = true,
) {
  const native = useCalculation(workbook, provider, options, enabled && mode?.mode !== 'external');
  const inputKey = useMemo(() => workbookCalculationKey(workbook), [workbook]);
  const accepted = useRef<{ revision: number; payload: string } | null>(null);
  const local = useRef({ result: native.result, revision: 0 });
  if (native.result && local.current.result !== native.result)
    local.current = { result: native.result, revision: local.current.revision + 1 };
  return useMemo(() => {
    let result = native.result,
      error = native.error,
      pending = native.pending,
      revision = local.current.revision,
      valid = true;
    const snapshot = mode?.mode === 'external' ? mode.snapshot : null;
    if (mode?.mode === 'external') {
      result = null;
      error = null;
      pending = true;
      if (snapshot)
        try {
          assertExternalCalculationSnapshot(snapshot, workbook);
          const payload = orderedJson(snapshot);
          if (
            accepted.current &&
            (snapshot.revision < accepted.current.revision ||
              (snapshot.revision === accepted.current.revision &&
                payload !== accepted.current.payload))
          )
            throw new Error('Conflicting or obsolete calculation revision.');
          accepted.current = { revision: snapshot.revision, payload };
          revision = snapshot.revision;
          result = snapshot.result ?? null;
          pending = snapshot.status === 'pending';
          error =
            snapshot.error ??
            (snapshot.status === 'unavailable' ? 'Calculation unavailable.' : null);
        } catch (cause) {
          error = cause instanceof Error ? cause.message : 'Invalid calculation publication.';
          pending = false;
          valid = false;
        }
    }
    const readyCells = new Set<string>();
    if (result)
      for (const sheet of workbook.worksheets)
        sheet.data.forEach((row, y) =>
          row.forEach((_input, x) => {
            const key = calculationCellKey(sheet.id, x, y);
            if (
              valid &&
              (mode?.mode === 'external'
                ? (snapshot?.cellStates?.[key] ?? snapshot?.status) === 'ready'
                : !error && !pending)
            )
              readyCells.add(key);
          }),
        );
    return {
      result,
      error,
      pending,
      inputKey,
      revision,
      readyCells,
      status: error
        ? ('unavailable' as const)
        : pending
          ? ('pending' as const)
          : ('ready' as const),
    };
  }, [workbook, mode, native.result, native.error, native.pending, inputKey]);
}

export function acceptedPresentation(
  value: CellPresentationSnapshot | undefined,
  inputKey: string,
  revision: number,
): CellPresentationSnapshot | undefined {
  if (!value || value.inputKey !== inputKey || value.revision !== revision) return undefined;
  if (!value.cells || typeof value.cells !== 'object' || Object.keys(value.cells).length > 250_000)
    throw new TypeError('Invalid cell presentation map.');
  for (const entry of Object.values(value.cells)) {
    for (const field of ['text', 'tooltip', 'ariaDescription', 'defaultUnit'] as const)
      if (
        entry[field] !== undefined &&
        (typeof entry[field] !== 'string' || entry[field]!.length > 32_768)
      )
        throw new TypeError('Invalid cell presentation text.');
    if (
      entry.status !== undefined &&
      !['resolved', 'stale', 'missing', 'error'].includes(entry.status)
    )
      throw new TypeError('Invalid cell presentation status.');
    if (entry.decoration !== undefined && !/^[a-z][a-z0-9-]{0,63}$/.test(entry.decoration))
      throw new TypeError('Invalid cell presentation decoration.');
  }
  return value;
}
