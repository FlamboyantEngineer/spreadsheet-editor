'use client';
import { useMemo, useRef } from 'react';
import type { CalculationResult } from '../calculation/evaluator.js';
import { calculationFormats, worksheetCalculationExtent } from '../model/extent.js';
import type { Workbook } from '../model/workbook.js';
import type { CalculationOptions } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import {
  assertExternalCalculationSnapshot,
  calculationCellKey,
  pendingCalculationCells,
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
  const native = useCalculation(
    workbook,
    provider,
    options,
    enabled && mode?.mode !== 'external',
    mode?.mode !== 'server',
  );
  const inputKey = useMemo(() => workbookCalculationKey(workbook), [workbook]);
  const accepted = useRef<{ revision: number; payload: string } | null>(null);
  const local = useRef({ result: native.result, revision: 0 });
  if (native.result && local.current.result !== native.result)
    local.current = { result: native.result, revision: local.current.revision + 1 };
  const kind = mode?.mode,
    snapshot = mode?.mode === 'external' ? mode.snapshot : null;
  // Keyed on the mode's kind and snapshot, so a host's inline calculation prop costs nothing.
  return useMemo(() => {
    let result = native.result,
      error = native.error,
      pending = native.pending,
      revision = local.current.revision,
      valid = true;
    if (kind === 'external') {
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
              (kind === 'external'
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
  }, [workbook, kind, snapshot, native.result, native.error, native.pending, inputKey]);
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

/** Cells whose calculation inputs differ between two workbooks of the same structure, or null. */
function changedCalculationCells(
  before: Workbook,
  after: Workbook,
): { cells: string[]; extents: string[] } | null {
  if (before.worksheets.length !== after.worksheets.length) return null;
  const cells: string[] = [],
    extents: string[] = [];
  for (const [index, sheet] of before.worksheets.entries()) {
    const next = after.worksheets[index];
    const inputs = (value: typeof sheet) =>
      orderedJson({
        id: value.id,
        name: value.name,
        inputValues: value.inputValues,
        rowIds: value.rowIds,
        columnIds: value.columnIds,
        formats: calculationFormats(value),
      });
    if (inputs(sheet) !== inputs(next)) return null;
    const a = worksheetCalculationExtent(sheet),
      b = worksheetCalculationExtent(next);
    if (a.rows !== b.rows || a.columns !== b.columns) extents.push(sheet.id);
    for (let y = 0; y < Math.max(sheet.data.length, next.data.length); y++) {
      const previous = sheet.data[y],
        current = next.data[y];
      if (previous === current) continue;
      for (let x = 0; x < Math.max(previous?.length ?? 0, current?.length ?? 0); x++)
        if ((previous?.[x] ?? null) !== (current?.[x] ?? null))
          cells.push(calculationCellKey(sheet.id, x, y));
    }
  }
  return { cells, extents };
}

/**
 * What the editor shows while the client or server recalculates: the last ready result, with
 * the cells that read an edited cell pending, like a draft. The publication itself stays pending,
 * so exports and insertions still wait for the current result.
 */
export function useRecalculatingDisplay(
  workbook: Workbook,
  publication: ReturnType<typeof usePublication>,
  enabled: boolean,
): { result: CalculationResult | null; readyCells: ReadonlySet<string> } {
  const settled = useRef<{
    workbook: Workbook;
    result: CalculationResult;
    readyCells: ReadonlySet<string>;
  } | null>(null);
  if (publication.status === 'ready' && publication.result)
    settled.current = { workbook, result: publication.result, readyCells: publication.readyCells };
  const last = settled.current;
  const recalculating = enabled && publication.status === 'pending' && last !== null;
  return useMemo(() => {
    if (!recalculating) return publication;
    const changed = changedCalculationCells(last.workbook, workbook);
    if (!changed) return publication;
    const pending = pendingCalculationCells(last.result, changed.cells, changed.extents);
    const readyCells = new Set<string>();
    for (const key of last.readyCells) if (!pending.has(key)) readyCells.add(key);
    return { result: last.result, readyCells };
  }, [recalculating, last, workbook, publication]);
}
