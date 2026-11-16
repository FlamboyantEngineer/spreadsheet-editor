'use client';

import { useMemo, useRef } from 'react';
import type { CalculationOptions, CalculationResult } from '../calculation/evaluator.js';
import type { CalculationProvider } from '../calculation/provider.js';
import {
  assertExternalCalculationSnapshot,
  calculationCellKey,
  orderedJson,
  pendingCalculationCells,
} from '../calculation/publication.js';
import { calculationFormats, worksheetCalculationExtent } from '../model/extent.js';
import type { Workbook, Worksheet } from '../model/workbook.js';
import type { CalculationMode } from './contracts.js';
import type { CellPresentationSnapshot } from './public-host.js';
import { acceptedCalculationKey, useCalculation } from './use-calculation.js';

/** Where a publication's calculation comes from. */
interface PublicationSources {
  mode: CalculationMode | undefined;
  provider: CalculationProvider;
  options: CalculationOptions;
  enabled?: boolean | undefined;
}

type ExternalSnapshot = Extract<CalculationMode, { mode: 'external' }>['snapshot'];
type AcceptedRevision = { revision: number; payload: string };

/** A host's snapshot as a publication, refusing one older than or conflicting with `accepted`. */
function externalPublication(
  snapshot: ExternalSnapshot,
  workbook: Workbook,
  accepted: AcceptedRevision | null,
) {
  if (!snapshot)
    return { result: null, error: null, pending: true, revision: undefined, valid: true };
  try {
    assertExternalCalculationSnapshot(snapshot, workbook);
    const payload = orderedJson(snapshot);
    if (
      accepted &&
      (snapshot.revision < accepted.revision ||
        (snapshot.revision === accepted.revision && payload !== accepted.payload))
    )
      throw new Error('Conflicting or obsolete calculation revision.');
    return {
      result: snapshot.result ?? null,
      error:
        snapshot.error ?? (snapshot.status === 'unavailable' ? 'Calculation unavailable.' : null),
      pending: snapshot.status === 'pending',
      revision: snapshot.revision,
      valid: true,
      accepted: { revision: snapshot.revision, payload },
    };
  } catch (cause) {
    return {
      result: null,
      error: cause instanceof Error ? cause.message : 'Invalid calculation publication.',
      pending: false,
      revision: undefined,
      valid: false,
    };
  }
}

/** Every authored cell key of `workbook` that `isReady` accepts. */
function readyCellsOf(workbook: Workbook, isReady: (key: string) => boolean) {
  const readyCells = new Set<string>();
  for (const sheet of workbook.worksheets)
    sheet.data.forEach((row, y) =>
      row.forEach((_input, x) => {
        const key = calculationCellKey(sheet.id, x, y);
        if (isReady(key)) readyCells.add(key);
      }),
    );
  return readyCells;
}

const publicationStatus = (error: string | null, pending: boolean) =>
  error ? ('unavailable' as const) : pending ? ('pending' as const) : ('ready' as const);

export function usePublication(
  workbook: Workbook,
  { mode, provider, options, enabled = true }: PublicationSources,
) {
  const native = useCalculation(workbook, provider, {
    options,
    enabled: enabled && mode?.mode !== 'external',
    trusted: mode?.mode !== 'server',
  });
  const inputKey = useMemo(() => acceptedCalculationKey(workbook), [workbook]);
  const accepted = useRef<AcceptedRevision | null>(null);
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
      const external = externalPublication(snapshot, workbook, accepted.current);
      if (external.accepted) accepted.current = external.accepted;
      ({ result, error, pending, valid } = external);
      revision = external.revision ?? revision;
    }
    const isReady =
      kind === 'external'
        ? (key: string) => (snapshot?.cellStates?.[key] ?? snapshot?.status) === 'ready'
        : () => !error && !pending;
    const readyCells = result && valid ? readyCellsOf(workbook, isReady) : new Set<string>();
    return {
      result,
      error,
      pending,
      inputKey,
      revision,
      readyCells,
      status: publicationStatus(error, pending),
    };
  }, [workbook, kind, snapshot, native.result, native.error, native.pending, inputKey]);
}

const PRESENTATION_TEXT_FIELDS = ['text', 'tooltip', 'ariaDescription', 'defaultUnit'] as const;
const PRESENTATION_STATUSES = ['resolved', 'stale', 'missing', 'error'];

function assertPresentationEntry(entry: CellPresentationSnapshot['cells'][string]) {
  for (const field of PRESENTATION_TEXT_FIELDS) {
    const text = entry[field];
    if (text !== undefined && (typeof text !== 'string' || text.length > 32_768))
      throw new TypeError('Invalid cell presentation text.');
  }
  if (entry.status !== undefined && !PRESENTATION_STATUSES.includes(entry.status))
    throw new TypeError('Invalid cell presentation status.');
  if (entry.decoration !== undefined && !/^[a-z][a-z0-9-]{0,63}$/.test(entry.decoration))
    throw new TypeError('Invalid cell presentation decoration.');
}

export function acceptedPresentation(
  value: CellPresentationSnapshot | undefined,
  { inputKey, revision }: { inputKey: string; revision: number },
): CellPresentationSnapshot | undefined {
  if (!value || value.inputKey !== inputKey || value.revision !== revision) return undefined;
  if (!value.cells || typeof value.cells !== 'object' || Object.keys(value.cells).length > 250_000)
    throw new TypeError('Invalid cell presentation map.');
  for (const entry of Object.values(value.cells)) assertPresentationEntry(entry);
  return value;
}

const calculationInputs = (value: Worksheet) =>
  orderedJson({
    id: value.id,
    name: value.name,
    inputValues: value.inputValues,
    rowIds: value.rowIds,
    columnIds: value.columnIds,
    formats: calculationFormats(value),
  });

/** The keys of the cells whose authored value differs between two versions of a worksheet. */
function changedCells(sheet: Worksheet, next: Worksheet, cells: string[]) {
  for (let y = 0; y < Math.max(sheet.data.length, next.data.length); y++) {
    const previous = sheet.data[y],
      current = next.data[y];
    if (previous === current) continue;
    for (let x = 0; x < Math.max(previous?.length ?? 0, current?.length ?? 0); x++)
      if ((previous?.[x] ?? null) !== (current?.[x] ?? null))
        cells.push(calculationCellKey(sheet.id, x, y));
  }
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
    if (calculationInputs(sheet) !== calculationInputs(next)) return null;
    const a = worksheetCalculationExtent(sheet),
      b = worksheetCalculationExtent(next);
    if (a.rows !== b.rows || a.columns !== b.columns) extents.push(sheet.id);
    changedCells(sheet, next, cells);
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
  // The last ready result, kept while the next one is pending. It is cached in a ref during the
  // render that receives it: state would render the editor a second time for every result.
  /* eslint-disable react-hooks/refs -- a render-local cache of the previous ready result */
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
  /* eslint-enable react-hooks/refs */
}
