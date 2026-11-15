'use client';

import { placePeers, collaboratorSheet, lockedChange, selectionLock } from '../model/presence.js';
import { attributeCellChanges } from '../model/provenance.js';

import { useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { calculationCellKey, pendingCalculationCells } from '../calculation/publication.js';
import { commandOperations } from './command-operations.js';
import { historyEntry, rebuildWorkbook, type HistoryEntry } from './history.js';
import {
  changeWorkbookStructureWithOperations,
  collectCellOperations,
  invertWorkbookOperations,
  type WorkbookOperation,
  type WorkbookCommandResult,
  type WorkbookChangeContext,
} from '../model/operations.js';
import { createWorkbookEntityId, worksheetCalculationExtent } from '../model/extent.js';
import {
  acceptedPresentation,
  usePublication,
  useRecalculatingDisplay,
} from './use-publication.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorState } from './public-host.js';
import { createClientCalculator } from '../calculation/provider.js';
import {
  createSpreadsheetCopy,
  pasteSpreadsheetCopy,
  type SpreadsheetCopy,
} from '../model/clipboard.js';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import {
  parseSpreadsheetRange,
  formatSpreadsheetAddress,
} from '../model/spreadsheet-addressing.js';
import { formatCellValue } from '../calculation/format.js';
import { summarizeCellValues } from '../calculation/summary.js';
import {
  appendWorksheet,
  assertWorkbook,
  copyWorkbook,
  createWorkbook,
  type Workbook,
  type Worksheet,
} from '../model/workbook.js';
import {
  borderCells,
  clearSelection,
  formatCells,
  INITIAL_SELECTION,
  mergeSelection,
  normalizeSelection,
  parseCellStyle,
  parseDelimited,
  readSelection,
  selectionAddress,
  selectionCells,
  serializeDelimited,
  writeCells,
  type CellSelection,
} from '../model/editing.js';
import {
  deleteWorksheet,
  renameWorksheet,
  validateWorksheetName,
  type StructureChange,
} from '../model/structure.js';
import {
  addComment,
  addWorkbookComment,
  findOpenComment,
  previewSuggestions,
  decideSuggestion,
  proposeWorkbookChange,
  type CellFormat,
  type ReviewStamp,
} from '../model/review.js';
import { createDocumentInsertion } from '../document/insertion.js';
import type { GridController, SpreadsheetEditorProps, SpreadsheetMode } from './contracts.js';
import {
  FULL_CAPABILITIES,
  READ_ONLY_CAPABILITIES,
  resolveSpreadsheetMode,
} from '../host/capabilities.js';
import type { HostResult, WorkbookVersion } from '../host/adapters.js';
import { useWorkbookVersions } from './use-workbook-versions.js';
import { usePersonNames } from './use-person-names.js';
import { useWorkbookPresence } from './use-workbook-presence.js';
import { saveFailure, sessionSavePhase, useSessionState } from './use-session-state.js';

export type EditorDialog =
  'row-height' | 'column-width' | 'import' | 'export' | 'confirm-structure' | null;
const NO_CELLS: ReadonlySet<string> = new Set();
/** Cell addresses a command wrote, by worksheet id. */
type CellTargets = Map<string, Iterable<string>>;
function mergeTargets(a: CellTargets, b: CellTargets): CellTargets {
  const merged: CellTargets = new Map();
  for (const source of [a, b])
    for (const [id, addresses] of source)
      merged.set(id, new Set([...(merged.get(id) ?? []), ...addresses]));
  return merged;
}
function blockTargets(sheetId: string, x: number, y: number, values: unknown[][]): CellTargets {
  return new Map([
    [
      sheetId,
      values.flatMap((row, dy) => row.map((_cell, dx) => formatSpreadsheetAddress(x + dx, y + dy))),
    ],
  ]);
}
/** Ready cells less a draft's pending cells, without copying the ready set per keystroke. */
function withoutPending(
  ready: ReadonlySet<string>,
  pending: ReadonlySet<string>,
): ReadonlySet<string> {
  if (!pending.size) return ready;
  const view: ReadonlySet<string> = {
    get size() {
      return [...view].length;
    },
    has: (key) => ready.has(key) && !pending.has(key),
    *[Symbol.iterator]() {
      for (const key of ready) if (!pending.has(key)) yield key;
    },
    *keys() {
      yield* view;
    },
    *values() {
      yield* view;
    },
    *entries() {
      for (const key of view) yield [key, key];
    },
    forEach(callback, thisArg) {
      for (const key of view) callback.call(thisArg, key, key, view);
    },
  } as ReadonlySet<string>;
  return view;
}
function readyValues<T>(
  values: readonly (readonly T[])[] | undefined,
  ready: ReadonlySet<string>,
  sheetId: string,
): (T | null)[][] | undefined {
  return values?.map((row, y) =>
    row.map((value, x) => (ready.has(calculationCellKey(sheetId, x, y)) ? value : null)),
  );
}
/**
 * Blank a draft's pending cells; rows without one keep their identity. A row with a pending cell
 * is always a new row, even when that cell is already blank: the grid does not paint a row whose
 * identity is unchanged, so it would miss the cell's change of readiness (SE-PERF-1).
 */
function withoutPendingValues<T>(
  values: (T | null)[][] | undefined,
  pending: ReadonlySet<string>,
  sheetId: string,
): (T | null)[][] | undefined {
  if (!values || !pending.size) return values;
  let next: (T | null)[][] | undefined;
  for (const key of pending) {
    const [id, x, y] = JSON.parse(key) as [string, number, number];
    if (id !== sheetId || !values[y]) continue;
    next ??= [...values];
    if (next[y] === values[y]) next[y] = [...values[y]];
    if (x < next[y].length) next[y][x] = null;
  }
  return next ?? values;
}
function immutableBaseline(workbook: Workbook): Workbook {
  const snapshot = copyWorkbook(workbook),
    pending: object[] = [snapshot];
  while (pending.length) {
    const value = pending.pop()!;
    if (Object.isFrozen(value)) continue;
    for (const child of Object.values(value))
      if (child && typeof child === 'object') pending.push(child);
    Object.freeze(value);
  }
  return snapshot;
}
export function useSpreadsheetEditor(props: SpreadsheetEditorProps) {
  const [local, setLocal] = useState(() => copyWorkbook(props.defaultValue ?? createWorkbook()));
  const session = useSessionState(props.session ?? undefined);
  /** A saved version the reader opened, shown read-only in place of the workbook (SE-ADAPTERS-5.1). */
  const [viewedVersion, setViewedVersion] = useState<{
    version: WorkbookVersion;
    content: Workbook;
  } | null>(null);
  const workbook = viewedVersion?.content ?? props.value ?? session.workbook ?? local;
  const [history, setHistory] = useState<{
    head: Workbook | null;
    past: HistoryEntry[];
    future: HistoryEntry[];
  }>({ head: null, past: [], future: [] });
  const currentHistory =
    history.head === workbook ? history : { head: workbook, past: [], future: [] };
  const [localMode, setLocalMode] = useState<SpreadsheetMode>(props.defaultMode ?? 'author');
  const requestedMode = props.mode ?? localMode;
  const hostCapabilities = props.capabilities ?? FULL_CAPABILITIES;
  const capabilities = viewedVersion ? READ_ONLY_CAPABILITIES : hostCapabilities;
  const highestMode = resolveSpreadsheetMode(capabilities, { historical: !!viewedVersion });
  /** Nothing is allowed: no mode can be chosen. */
  const locked = highestMode === 'read';
  const mode: SpreadsheetMode =
    highestMode === 'author'
      ? requestedMode
      : highestMode === 'review' &&
          ((capabilities.suggest && props.features?.suggestions !== false) ||
            (capabilities.comment && props.features?.comments !== false))
        ? 'review'
        : 'read';
  const canEdit = mode === 'author' && capabilities.edit;
  const canSuggest =
    mode === 'review' && capabilities.suggest && props.features?.suggestions !== false;
  const canWrite = canEdit || canSuggest;
  const canComment = mode !== 'read' && capabilities.comment && props.features?.comments !== false;
  const canDecide = canEdit && capabilities.decide && props.features?.suggestions !== false;
  const actor = useMemo(() => {
    const person = props.identity?.currentPerson();
    return person ? { id: person.id, name: person.name } : { id: 'local-user', name: 'You' };
  }, [props.identity]);
  const [sheetId, setSheetId] = useState(props.initialSheetId ?? workbook.worksheets[0].id);
  const sheet = workbook.worksheets.find((sheet) => sheet.id === sheetId) ?? workbook.worksheets[0];
  const [selection, setSelection] = useState<CellSelection>(INITIAL_SELECTION);
  const [editingCell, setEditingCell] = useState(false);
  const place = { sheetId: sheet.id, range: selectionAddress(selection) };
  const remotePeers = useWorkbookPresence(props.presence, props.workbookId, {
    at: place,
    editLock: editingCell ? place : null,
  });
  const peers = useMemo(
    () => placePeers(workbook, remotePeers, actor.id),
    [workbook, remotePeers, actor.id],
  );
  const lockedSelection = selectionLock(workbook, sheet.id, selection, peers);
  const [draft, setDraftState] = useState<{ key: string; value: string } | null>(null);
  const draftRef = useRef(draft);
  function setDraft(next: typeof draft) {
    draftRef.current = next;
    setDraftState(next);
    const target = next && /^(.*):(\d+):(\d+)$/.exec(next.key);
    props.onDraftChange?.(
      next && target
        ? { sheetId: target[1], x: Number(target[2]), y: Number(target[3]), text: next.value }
        : null,
    );
  }
  const grid = useRef<GridController | null>(null);
  const copied = useRef<SpreadsheetCopy | null>(null);
  const [panel, setPanel] = useState<string | null>(null),
    [panelOpen, setPanelOpen] = useState(false);
  const versionList = useWorkbookVersions(
    hostCapabilities.history ? props.versions : undefined,
    props.workbookId,
    panelOpen && panel === 'history',
  );
  const versionAuthors = useMemo(
    () => [viewedVersion?.version.author ?? null, ...versionList.versions.map((v) => v.author)],
    [viewedVersion, versionList.versions],
  );
  const personName = usePersonNames(props.identity, workbook.comments, actor, versionAuthors);
  const [commentComposer, setCommentComposer] = useState(false);
  const [commentScope, setCommentScope] = useState<'range' | 'workbook'>('range');
  const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({});
  const [showContributors, setShowContributors] = useState(false);
  const [showFind, setShowFind] = useState(false),
    [query, setQueryState] = useState(''),
    [replacement, setReplacement] = useState('');
  const [activeMatch, setActiveMatch] = useState<string | null>(null);
  function setQuery(value: string) {
    setQueryState(value);
    setActiveMatch(null);
  }
  const [zoom, setZoom] = useState(100),
    [showGridlines, setShowGridlines] = useState(true),
    [showFormulas, setShowFormulas] = useState(false);
  const [notice, setNoticeState] = useState<string | null>(null);
  const errorNotice = useRef<string | null>(null);
  const [noticeTone, setNoticeTone] = useState<'danger' | 'warning'>('danger');
  const [noticeRevision, setNoticeRevision] = useState(0);
  function setNotice(message: string | null) {
    errorNotice.current = message;
    setNoticeTone('danger');
    setNoticeState(message);
  }
  function warn(message: string) {
    errorNotice.current = null;
    setNoticeTone('warning');
    setNoticeState(message);
    setNoticeRevision((value) => value + 1);
  }
  useEffect(() => {
    if (notice && noticeTone === 'warning') {
      const timer = setTimeout(() => setNoticeState(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [notice, noticeTone, noticeRevision]);
  const pendingStructure = useRef<(() => void) | null>(null);
  const [dialog, setDialog] = useState<EditorDialog>(null),
    [dialogValue, setDialogValue] = useState('');
  const listeners = useRef(new Set<() => void>());
  const [flushing, setFlushing] = useState(false);
  const savePhase = sessionSavePhase(session.status);
  const sheetViews = useRef(
    new Map<string, { selection: CellSelection; viewport: { left: number; top: number } }>(),
  );
  const [nativeDraft, setNativeDraftState] = useState<{
    x: number;
    y: number;
    text: string;
  } | null>(null);
  const nativeDraftRef = useRef(nativeDraft);
  function setNativeDraft(next: typeof nativeDraft) {
    nativeDraftRef.current = next;
    setNativeDraftState(next);
    props.onDraftChange?.(next ? { sheetId: sheet.id, ...next } : null);
  }
  const alive = useRef(true);
  useEffect(() => {
    alive.current = true;
    return () => {
      alive.current = false;
      listeners.current.clear();
    };
  }, []);
  const latestWorkbook = useRef(workbook);
  const live = useRef({ canEdit, canWrite, props });
  useLayoutEffect(() => {
    live.current = { canEdit, canWrite, props };
  }, [canEdit, canWrite, props]);
  // Read the callback through `live` so a host's inline callback is not a dependency (SE-HOST-2).
  useEffect(() => {
    live.current.props.onPanelChange?.({ id: panel, expanded: panelOpen });
  }, [panel, panelOpen]);
  useEffect(() => {
    latestWorkbook.current = workbook;
  }, [workbook]);
  const client = useMemo(
    () =>
      createClientCalculator(
        props.calculation?.mode === 'client' ? props.calculation.host : undefined,
      ),
    [props.calculation?.mode === 'client' ? props.calculation.host : undefined],
  );
  const [refreshVersion, setRefreshVersion] = useState(0);
  const baseProvider = props.calculation?.mode === 'server' ? props.calculation.provider : client;
  const provider = useMemo(
    () => ({ calculate: baseProvider.calculate.bind(baseProvider) }),
    [baseProvider, refreshVersion],
  );
  const calculationOptions = useMemo(
    () => ({
      timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      ...(props.calculation?.mode === 'external' ? undefined : props.calculation?.options),
    }),
    [props.calculation?.mode === 'external' ? undefined : props.calculation?.options],
  );
  const calculation = usePublication(workbook, props.calculation, provider, calculationOptions);
  const previewWorkbook = useMemo(
    () => (canSuggest ? previewSuggestions(workbook, actor.id) : workbook),
    [workbook, actor.id, canSuggest],
  );
  const previewCalculation = usePublication(
    previewWorkbook,
    props.calculation?.mode === 'external'
      ? { mode: 'external', snapshot: props.reviewCalculation ?? null }
      : props.calculation,
    provider,
    calculationOptions,
    previewWorkbook !== workbook,
  );
  const visibleCalculation = previewWorkbook !== workbook ? previewCalculation : calculation;
  const shown = useRecalculatingDisplay(
    previewWorkbook,
    visibleCalculation,
    props.calculation?.mode !== 'external',
  );
  const publication = useMemo(
    () => ({ inputKey: visibleCalculation.inputKey, revision: visibleCalculation.revision }),
    [visibleCalculation.inputKey, visibleCalculation.revision],
  );
  const previewSheet = previewWorkbook.worksheets.find((value) => value.id === sheet.id)!;
  const authored = previewSheet.data[selection.y1]?.[selection.x1] ?? '';
  const draftKey = `${sheet.id}:${selection.x1}:${selection.y1}`;
  const draftValue = draft?.key === draftKey ? draft.value : String(authored);
  const draftDirty = draft?.key === draftKey && draft.value !== String(authored);
  const draftCell = draftDirty
    ? { x: selection.x1, y: selection.y1, text: draft!.value }
    : nativeDraft;
  const calculationExtent = useMemo(() => worksheetCalculationExtent(sheet), [sheet]);
  const pendingCells = useMemo(
    () =>
      draftCell
        ? pendingCalculationCells(
            shown.result,
            [calculationCellKey(sheet.id, draftCell.x, draftCell.y)],
            draftCell.x >= calculationExtent.columns || draftCell.y >= calculationExtent.rows
              ? [sheet.id]
              : [],
          )
        : NO_CELLS,
    [
      shown.result,
      sheet.id,
      calculationExtent.columns,
      calculationExtent.rows,
      draftCell?.x,
      draftCell?.y,
      draftCell?.text,
    ],
  );
  const readyCells = useMemo(
    () => withoutPending(shown.readyCells, pendingCells),
    [shown.readyCells, pendingCells],
  );
  const presentation = acceptedPresentation(
    props.cellPresentation,
    visibleCalculation.inputKey,
    visibleCalculation.revision,
  );
  const acceptedOutput = shown.result?.worksheets.find((result) => result.id === sheet.id)?.values;
  // Ready values change with the result; a draft only blanks the rows its pending cells are in.
  const readyOutput = useMemo(
    () => readyValues(acceptedOutput, shown.readyCells, sheet.id),
    [acceptedOutput, shown.readyCells, sheet.id],
  );
  const output = useMemo(
    () => withoutPendingValues(readyOutput, pendingCells, sheet.id),
    [readyOutput, pendingCells, sheet.id],
  );
  const style = parseCellStyle(
    previewSheet.style?.[selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })],
  );
  const cellFormat = previewSheet.formats?.[
    selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })
  ] ?? { kind: 'general' as const };

  /** Shares whether the reader is writing in the selected cells, which locks them for others. */
  function publishPresence(editing: boolean) {
    setEditingCell(editing);
  }
  function followCollaborator(id: string) {
    const person = peers.find((person) => person.id === id),
      target = person && collaboratorSheet(workbook, person);
    if (person?.selection && target) select(person.selection, target.id);
  }
  function stamp(): ReviewStamp {
    return { id: globalThis.crypto.randomUUID(), time: new Date().toISOString(), actor };
  }
  function publish(
    next: Workbook,
    operations: WorkbookOperation[] = commandOperations(workbook, next),
    source: WorkbookChangeContext['source'] = 'user',
  ) {
    assertWorkbook(next);
    const previous = workbook;
    let baseline: Workbook | undefined;
    const context: WorkbookChangeContext = {
      transactionId: (props.idFactory ?? createWorkbookEntityId)('transaction'),
      source,
      // The detached, frozen copy is made only for a host that reads it (SE-PERF-1).
      get baseline() {
        return (baseline ??= immutableBaseline(previous));
      },
      operations: structuredClone(operations),
    };
    if (viewedVersion) {
      setNotice('A saved version is read-only.');
      throw new Error('A saved version is read-only.');
    }
    if (props.session && !props.session.workbook()) {
      setNotice('The workbook is still opening.');
      throw new Error('The workbook is still opening.');
    }
    try {
      props.onChange?.(next, context);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'The workbook change was rejected.');
      throw error;
    }
    latestWorkbook.current = next;
    if (props.session) props.session.change(next, context);
    else if (!props.value) setLocal(next);
  }
  /** `targets` names the cells a command wrote, so its patches need no whole-workbook diff. */
  function commit(
    next: Workbook,
    label = 'Edit workbook',
    operations?: WorkbookOperation[],
    targets?: CellTargets,
  ) {
    if (next === workbook) return;
    const acceptedOperations = operations ?? commandOperations(workbook, next, targets);
    publish(next, acceptedOperations);
    setHistory({
      head: next,
      past: [
        ...currentHistory.past.slice(-49),
        historyEntry(workbook, next, label, structuredClone(acceptedOperations)),
      ],
      future: [],
    });
    setEditingCell(false);
  }
  function mutate(
    next: Workbook,
    label: string,
    operation?: { sheetId: string; change: StructureChange },
    operations?: WorkbookOperation[],
    targets?: CellTargets,
  ) {
    if (!alive.current) return false;
    if (!canWrite) {
      warn('This workbook is read-only.');
      return false;
    }
    if (latestWorkbook.current !== workbook) {
      setNotice('The workbook changed before this action completed. Try again.');
      return false;
    }
    const locked = lockedChange(workbook, next, peers);
    if (locked) {
      warn(`${locked.name} is editing a cell in this selection. Try again after they finish.`);
      return false;
    }
    const accepted = canSuggest
      ? proposeWorkbookChange(workbook, next, label, stamp(), operation)
      : operation
        ? next
        : attributeCellChanges(workbook, next, {
            actor,
            source: 'human',
            at: new Date().toISOString(),
          });
    commit(
      accepted,
      canSuggest ? `Suggest: ${label}` : label,
      canSuggest ? undefined : operations,
      targets,
    );
    if (canSuggest) {
      setPanel('suggestions');
      setPanelOpen(true);
    }
    return true;
  }
  /** Compose a protected formula draft and a chrome command into one history entry. */
  function mutateWithDraft(
    transform: (base: Workbook) => Workbook | WorkbookCommandResult,
    label: string,
    operation?: { sheetId: string; change: StructureChange },
    written?: CellTargets,
  ): boolean {
    const pending = draftRef.current,
      draftTargets: CellTargets = new Map();
    let base = workbook;
    if (pending?.key === draftKey && pending.value !== String(authored)) {
      draftTargets.set(sheet.id, [formatSpreadsheetAddress(selection.x1, selection.y1)]);
      base = writeCells(
        base,
        sheet.id,
        selection.x1,
        selection.y1,
        [[pending.value]],
        props.idFactory,
      );
      const locked = lockedChange(workbook, base, peers);
      if (locked) {
        warn(`${locked.name} is editing a cell in this selection. Try again after they finish.`);
        return false;
      }
      // Structural transforms move this attribution together with the authored cell; other
      // commands are attributed once, in `mutate`.
      if (canEdit && operation)
        base = attributeCellChanges(workbook, base, {
          actor,
          source: 'human',
          at: new Date().toISOString(),
        });
    }
    const transformed = transform(base);
    const result =
      'workbook' in transformed ? transformed : { workbook: transformed, operations: undefined };
    if (
      !mutate(
        result.workbook,
        label,
        operation,
        result.operations
          ? [...collectCellOperations(workbook, base, draftTargets), ...result.operations]
          : undefined,
        written && mergeTargets(written, draftTargets),
      )
    )
      return false;
    if (draftRef.current === pending && pending?.key === draftKey) setDraft(null);
    return true;
  }
  function replaceWorkbook(
    next: Workbook,
    label: string,
    reason: 'import' | 'new-workbook' | 'reset' = 'import',
  ): boolean {
    if (
      (reason === 'import' && props.features?.import === false) ||
      (reason === 'new-workbook' && props.features?.newWorkbook === false)
    )
      return false;
    if (
      !canEdit ||
      !mutate(next, label, undefined, [
        { kind: 'replace', reason, before: copyWorkbook(workbook), after: copyWorkbook(next) },
      ])
    )
      return false;
    setDraft(null);
    return true;
  }
  function attempt(action: () => unknown) {
    try {
      void Promise.resolve(action()).catch((error) =>
        setNotice(error instanceof Error ? error.message : 'The action failed.'),
      );
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'The action failed.');
    }
  }
  function setMode(next: SpreadsheetMode) {
    if (locked || (!capabilities.edit && next !== mode)) return;
    if (!applyFormula()) return;
    setLocalMode(next);
    props.onModeChange?.(next);
    setDraft(null);
    grid.current?.releaseKeyboard();
  }
  function select(range: CellSelection, targetSheetId = sheet.id) {
    const normalized = normalizeSelection(range);
    selectionCells(normalized);
    if (!workbook.worksheets.some((value) => value.id === targetSheetId)) return false;
    grid.current?.commitDraft?.();
    if (!applyFormula()) return false;
    if (targetSheetId !== sheet.id) rememberSheetView();
    setSheetId(targetSheetId);
    setSelection(normalized);
    setDraft(null);
    if (targetSheetId === sheet.id) grid.current?.select(normalized);
    return true;
  }
  function activateSheet(id: string) {
    if (id === sheet.id) return true;
    if (!workbook.worksheets.some((value) => value.id === id)) return false;
    grid.current?.commitDraft?.();
    if (!applyFormula()) return false;
    rememberSheetView();
    setSheetId(id);
    setSelection(sheetViews.current.get(id)?.selection ?? INITIAL_SELECTION);
    setDraft(null);
    return true;
  }
  function rememberSheetView() {
    sheetViews.current.set(sheet.id, {
      selection: { ...selection },
      viewport: grid.current?.getViewport?.() ?? { left: 0, top: 0 },
    });
    grid.current?.releaseKeyboard();
  }
  function commitDraft(): Workbook {
    if (!alive.current) throw new Error('The editor session has ended.');
    grid.current?.commitDraft?.();
    if (!applyFormula()) throw new Error('The active cell edit could not be committed.');
    return copyWorkbook(latestWorkbook.current);
  }
  function cancelDraft() {
    if (!alive.current) throw new Error('The editor session has ended.');
    grid.current?.cancelDraft?.();
    setNativeDraft(null);
    setDraft(null);
  }
  function onGridSelection(range: CellSelection) {
    attempt(() => {
      const normalized = normalizeSelection(range);
      selectionCells(normalized);
      if (
        selection.x1 === normalized.x1 &&
        selection.x2 === normalized.x2 &&
        selection.y1 === normalized.y1 &&
        selection.y2 === normalized.y2
      )
        return;
      if (!applyFormula()) {
        grid.current?.select(selection);
        return;
      }
      setSelection((previous) =>
        previous.x1 === normalized.x1 &&
        previous.x2 === normalized.x2 &&
        previous.y1 === normalized.y1 &&
        previous.y2 === normalized.y2
          ? previous
          : normalized,
      );
      setDraft(null);
    });
  }
  function updateSheet(next: Worksheet) {
    const previous = workbook.worksheets.find((value) => value.id === next.id);
    if (!previous) return false;
    if (JSON.stringify(next.data) === JSON.stringify(previous.data))
      return mutateWithDraft(
        (base) => ({
          ...base,
          worksheets: base.worksheets.map((value) =>
            value.id === next.id
              ? { ...next, data: value.data, cellProvenance: value.cellProvenance }
              : value,
          ),
        }),
        'Update worksheet layout',
      );
    // Native cell commits already contain the native editor's protected draft.
    const changed = mutate(
      {
        ...workbook,
        worksheets: workbook.worksheets.map((sheet) => (sheet.id === next.id ? next : sheet)),
      },
      'Edit cells',
    );
    if (changed) setDraft(null);
    return changed;
  }
  function undo() {
    if (!canWrite) return;
    if (props.undo) {
      attempt(() => props.undo!.undo());
      return;
    }
    const previous = currentHistory.past.at(-1);
    if (!previous) return;
    const target =
      'snapshot' in previous ? previous.snapshot : rebuildWorkbook(workbook, previous, 'before');
    publish(target, invertWorkbookOperations(previous.operations), 'undo');
    setDraft(null);
    setHistory({
      head: target,
      past: currentHistory.past.slice(0, -1),
      future: [
        'snapshot' in previous ? { ...previous, snapshot: copyWorkbook(workbook) } : previous,
        ...currentHistory.future,
      ],
    });
  }
  function redo() {
    if (!canWrite) return;
    if (props.undo) {
      attempt(() => props.undo!.redo());
      return;
    }
    const next = currentHistory.future[0];
    if (!next) return;
    const target = 'snapshot' in next ? next.snapshot : rebuildWorkbook(workbook, next, 'after');
    publish(target, next.operations, 'redo');
    setDraft(null);
    setHistory({
      head: target,
      past: [
        ...currentHistory.past,
        'snapshot' in next ? { ...next, snapshot: copyWorkbook(workbook) } : next,
      ],
      future: currentHistory.future.slice(1),
    });
  }
  function applyFormula(): boolean {
    const pending = draftRef.current;
    if (pending?.key !== draftKey || pending.value === String(authored)) return true;
    return mutateWithDraft(
      (base) => base,
      `Edit ${selectionAddress(selection)}`,
      undefined,
      new Map(),
    );
  }
  function setDraftValue(value: string) {
    setDraft({ key: draftKey, value });
  }
  function proposeCells(
    values: import('../model/workbook.js').CellInput[][],
    x: number,
    y: number,
  ) {
    const changed = mutate(
      writeCells(workbook, sheet.id, x, y, values, props.idFactory),
      'Edit cells',
      undefined,
      undefined,
      blockTargets(sheet.id, x, y, values),
    );
    if (changed) setDraft(null);
    return changed;
  }
  function format(properties: Record<string, string | null>, format?: CellFormat | null) {
    if (format && props.formatting?.kinds && !props.formatting.kinds.includes(format.kind))
      throw new Error('The host does not support this number format.');
    if (format?.decimals !== undefined && props.formatting?.decimalPlaces === false)
      throw new Error('The host does not support decimal-place overrides.');
    mutateWithDraft(
      (base) => formatCells(base, sheet.id, selection, properties, format),
      `Format ${selectionAddress(selection)}`,
      undefined,
      new Map([[sheet.id, selectionCells(selection).map((cell) => cell.address)]]),
    );
  }
  function toggleFormat(property: string, value: string) {
    format({ [property]: style[property] === value ? null : value });
  }
  function border(kind: 'all' | 'outside' | 'none' | 'top' | 'right' | 'bottom' | 'left') {
    mutateWithDraft((base) => borderCells(base, sheet.id, selection, kind), `Set ${kind} borders`);
  }
  function structure(axis: StructureChange['axis'], kind: StructureChange['kind'], after = false) {
    if (!canEdit) {
      warn('Worksheet structure can be changed only in Authoring mode.');
      return;
    }
    const first = axis === 'row' ? selection.y1 : selection.x1,
      last = axis === 'row' ? selection.y2 : selection.x2;
    const change = { axis, kind, index: after ? last + 1 : first, count: last - first + 1 };
    const apply = () =>
      mutateWithDraft(
        (base) => changeWorkbookStructureWithOperations(base, sheet.id, change, props.idFactory),
        `${kind === 'insert' ? 'Insert' : 'Delete'} ${axis}${last > first ? 's' : ''}`,
        { sheetId: sheet.id, change },
      );
    const crossesMerge = Object.entries(sheet.mergeCells ?? {}).some(
      ([address, [columns, rows]]) => {
        const range = parseSpreadsheetRange(address)!;
        const start = axis === 'row' ? range.start.y : range.start.x;
        const end = start + (axis === 'row' ? rows : columns) - 1;
        return kind === 'insert'
          ? change.index > start && change.index <= end
          : first <= end && last >= start;
      },
    );
    if (crossesMerge) {
      pendingStructure.current = apply;
      setDialog('confirm-structure');
    } else apply();
  }
  function openPanel(id: string) {
    setPanel(id);
    setPanelOpen(true);
    grid.current?.releaseKeyboard();
  }
  function setHostPanel(id: string | null) {
    if (id === null) setPanelOpen(false);
    else openPanel(id);
  }
  function restoreDraft(value: { sheetId: string; x: number; y: number; text: string }) {
    if (
      !canWrite ||
      !Number.isSafeInteger(value.x) ||
      !Number.isSafeInteger(value.y) ||
      value.x < 0 ||
      value.y < 0 ||
      value.x > 16383 ||
      value.y > 1048575 ||
      value.text.length > 10000
    )
      throw new Error('The draft cannot be restored in this session.');
    if (!select({ x1: value.x, x2: value.x, y1: value.y, y2: value.y }, value.sheetId))
      throw new Error('The draft target is unavailable.');
    setDraft({ key: `${value.sheetId}:${value.x}:${value.y}`, value: value.text });
  }
  function openComment(id: string) {
    const thread = workbook.comments?.find((thread) => thread.id === id);
    if (!thread) return;
    if (thread.scope === 'workbook' || thread.state === 'resolved' || thread.orphaned) {
      setActiveCommentId(thread.id);
      return;
    }
    const range = parseSpreadsheetRange(thread.range ?? '');
    if (range && thread.sheetId)
      select(
        { x1: range.start.x, x2: range.end.x, y1: range.start.y, y2: range.end.y },
        thread.sheetId,
      );
    setCommentComposer(false);
    setActiveCommentId(id);
    grid.current?.releaseKeyboard();
  }
  function beginComment() {
    if (!canComment) return;
    const existing = findOpenComment(workbook, sheet.id, selectionAddress(selection));
    if (existing) {
      openComment(existing.id);
      return;
    }
    setCommentScope('range');
    setCommentComposer(true);
    setActiveCommentId('draft');
    grid.current?.releaseKeyboard();
  }
  function beginWorkbookComment() {
    if (!canComment) return;
    openPanel('comments');
    setCommentScope('workbook');
    setActiveCommentId(null);
    setCommentComposer(true);
  }
  function postComment(text: string) {
    if (!canComment) return;
    const meta = stamp();
    commit(
      commentScope === 'workbook'
        ? addWorkbookComment(workbook, text, meta)
        : addComment(workbook, sheet.id, selectionAddress(selection), text, meta),
      'Add comment',
    );
    setCommentComposer(false);
    if (commentScope === 'range') setActiveCommentId(meta.id);
  }
  useEffect(() => {
    if (props.initialCommentId) openComment(props.initialCommentId);
  }, [props.initialCommentId]);
  function openSuggestion(id: string) {
    const item = workbook.suggestions?.find((suggestion) => suggestion.id === id);
    if (item?.commentThreadId) openComment(item.commentThreadId);
  }
  useEffect(() => {
    const thread = workbook.comments?.find((thread) => thread.id === activeCommentId);
    if (activeCommentId !== 'draft' && (!thread || thread.state === 'resolved' || thread.orphaned))
      setActiveCommentId(null);
  }, [workbook]);
  function decide(id: string, action: 'accepted' | 'rejected') {
    if (canDecide)
      commit(
        decideSuggestion(workbook, id, action, stamp()),
        `${action === 'accepted' ? 'Accept' : 'Reject'} suggestion`,
      );
  }
  function insert(kind: 'table' | 'spreadsheet') {
    if (kind === 'table') {
      if (!calculation.result || draftCell)
        throw new Error('A document table needs current calculation results.');
      for (const { x, y } of selectionCells(selection)) {
        const input = sheet.data[y]?.[x],
          key = calculationCellKey(sheet.id, x, y);
        const explicit =
          props.calculation?.mode === 'external'
            ? props.calculation.snapshot?.cellStates?.[key]
            : undefined;
        if (
          !readyCells.has(key) &&
          ((input != null && input !== '') || (explicit ?? calculation.status) !== 'ready')
        )
          throw new Error('Every inserted cell must have a current ready result.');
      }
    }
    props.onInsert?.(
      createDocumentInsertion(workbook, calculation.result, {
        mode: kind,
        sheetId: sheet.id,
        range: selectionAddress(selection),
        display: props.display,
      }),
    );
  }
  const clipboard = props.clipboard ?? {
    async readText() {
      if (!globalThis.navigator?.clipboard)
        throw new Error('Clipboard access is unavailable. Use the grid’s paste shortcut.');
      return globalThis.navigator.clipboard.readText();
    },
    async writeText(text: string) {
      if (!globalThis.navigator?.clipboard) throw new Error('Clipboard access is unavailable.');
      await globalThis.navigator.clipboard.writeText(text);
    },
  };
  async function copy(cut = false) {
    if (cut && !canWrite) return;
    const pending = draftRef.current,
      hasDraft = pending?.key === draftKey && pending.value !== String(authored),
      source = hasDraft
        ? writeCells(
            previewWorkbook,
            sheet.id,
            selection.x1,
            selection.y1,
            [[pending.value]],
            props.idFactory,
          )
        : previewWorkbook;
    const cells = readSelection(source, sheet.id, selection);
    const displayed = cells.map((row, dy) =>
      row.map((cell, dx) => {
        const x = selection.x1 + dx,
          y = selection.y1 + dy;
        if (hasDraft && x === selection.x1 && y === selection.y1) return pending.value;
        if (!readyCells.has(calculationCellKey(sheet.id, x, y))) return cell;
        return formatWorksheetCell(
          shown.result?.worksheets.find((item) => item.id === sheet.id)?.values[y]?.[x] ?? null,
          sheet.formats?.[formatSpreadsheetAddress(x, y)],
          props.display,
        );
      }),
    );
    const text = serializeDelimited(displayed);
    await clipboard.writeText(text);
    if (cut && draftRef.current !== pending) {
      copied.current = createSpreadsheetCopy(source, sheet.id, selection, text, false);
      warn('The cell edit changed while copying. Nothing was cut. Try Cut again.');
      return;
    }
    copied.current = createSpreadsheetCopy(source, sheet.id, selection, text, cut);
  }
  function pasteCells(
    values: import('../model/workbook.js').CellInput[][],
    x = selection.x1,
    y = selection.y1,
    originalText?: string,
  ) {
    if (!canWrite) return;
    const text = originalText ?? serializeDelimited(values);
    const cached = copied.current,
      matching = cached && cached.text.replace(/\r\n/g, '\n') === text.replace(/\r\n/g, '\n');
    const changed = mutateWithDraft(
      (base) =>
        matching
          ? pasteSpreadsheetCopy(base, sheet.id, x, y, cached, selection, props.idFactory)
          : writeCells(base, sheet.id, x, y, values, props.idFactory),
      'Paste cells',
      undefined,
      matching ? undefined : blockTargets(sheet.id, x, y, values),
    );
    if (changed && cached?.cut && matching) copied.current = null;
  }
  async function paste() {
    if (canWrite) {
      const text = await clipboard.readText();
      pasteCells(parseDelimited(text), selection.x1, selection.y1, text);
    }
  }
  async function refresh(selected = false) {
    if (!canEdit || !props.onRefresh) return;
    await props.onRefresh(
      selected ? { sheetId: sheet.id, range: selectionAddress(selection) } : undefined,
    );
    if (live.current.canEdit) setRefreshVersion((version) => version + 1);
  }
  async function importExcel() {
    if (!canEdit || !props.onImportXlsx) return;
    const snapshot = workbook;
    const imported = await props.onImportXlsx();
    if (!imported) return;
    if (!alive.current || !live.current.canEdit) return;
    if (latestWorkbook.current !== snapshot)
      throw new Error('The workbook changed while importing. Try again.');
    assertWorkbook(imported);
    replaceWorkbook(imported, 'Import Excel workbook');
  }
  async function exportExcel() {
    if (!props.onExportXlsx || !calculation.result || draftCell) return;
    for (const target of workbook.worksheets)
      for (let y = 0; y < target.data.length; y++)
        for (let x = 0; x < target.data[y].length; x++)
          if (
            target.data[y][x] != null &&
            target.data[y][x] !== '' &&
            !calculation.readyCells.has(calculationCellKey(target.id, x, y))
          )
            throw new Error('Excel export requires current ready results for every authored cell.');
    await props.onExportXlsx(copyWorkbook(workbook), structuredClone(calculation.result));
  }
  /** Saves through the workbook session now (SE-SESSION-4). */
  async function save() {
    if (!alive.current) throw new Error('The editor session has ended.');
    commitDraft();
    const current = live.current.props.session;
    if (!current) return;
    if (!live.current.canEdit) {
      warn('Saving is available only in Authoring mode.');
      return;
    }
    try {
      await current.save();
      if (alive.current) setNotice(null);
    } catch (error) {
      if (alive.current) setNotice(saveFailure(error, current));
      throw error;
    }
  }
  /** Commits the draft and waits until the session holds no unsaved change (SE-SESSION-4). */
  async function flush(): Promise<void> {
    if (!alive.current) throw new Error('The editor session has ended.');
    commitDraft();
    const current = live.current.props.session;
    if (!current) return;
    setFlushing(true);
    try {
      await current.flush();
      if (alive.current) setNotice(null);
    } catch (error) {
      if (alive.current) setNotice(saveFailure(error, current));
      throw error;
    } finally {
      if (alive.current) setFlushing(false);
    }
  }
  async function resolveConflict(choice: 'mine' | 'stored') {
    const current = live.current.props.session;
    if (!current) return;
    try {
      await (choice === 'mine' ? current.keepMine() : current.takeStored());
      if (alive.current) setNotice(null);
    } catch (error) {
      if (alive.current) setNotice(saveFailure(error, current));
    }
  }
  /**
   * Commits the active cell edit and saves every change made so far, so a version action never
   * leaves an edit behind. Throws when an edit stays unsaved.
   */
  async function settleChanges() {
    commitDraft();
    const current = live.current.props.session;
    if (current?.status().dirty || current?.status().phase === 'saving') {
      await lifecycle.current.flush();
      if (current.status().dirty)
        throw new Error('Changes are not saved. Save or discard them first.');
    }
  }
  /** Runs a version action, showing its failure as a notice. Resolves whether it succeeded. */
  async function versionAction(action: () => Promise<void>, fallback: string) {
    try {
      await action();
      return true;
    } catch (error) {
      if (alive.current) setNotice(error instanceof Error ? error.message : fallback);
      return false;
    }
  }
  function versionStore() {
    const { versions, workbookId } = live.current.props;
    if (!versions || workbookId === undefined) throw new Error('This workbook has no versions.');
    return { versions, workbookId };
  }
  function unwrap<T>(result: HostResult<T>, fallback: string): T {
    if (!result.ok) throw new Error(result.error.message ?? fallback);
    return result.value;
  }
  const versionRequest = useRef(0);
  /** Saves pending changes, then shows `version` read-only in place of the workbook. */
  function openVersion(version: WorkbookVersion) {
    const request = ++versionRequest.current;
    return versionAction(async () => {
      if (!viewedVersion) await settleChanges();
      const { versions, workbookId } = versionStore();
      const content = unwrap(
        await versions.load(workbookId, version.revision),
        'Could not open the version.',
      );
      if (alive.current && request === versionRequest.current) {
        setNotice(null);
        setViewedVersion({ version, content });
      }
    }, 'Could not open the version.');
  }
  function returnToCurrent() {
    versionRequest.current++;
    setViewedVersion(null);
  }
  /** Saves pending changes, then saves the workbook as a new version named `label`. */
  function saveVersion(label: string | null) {
    return versionAction(async () => {
      await settleChanges();
      const { versions, workbookId } = versionStore();
      if (!versions.create) throw new Error('This workbook cannot save versions.');
      unwrap(
        await versions.create(workbookId, { content: commitDraft(), label }),
        'Could not save the version.',
      );
      versionList.refresh();
    }, 'Could not save the version.');
  }
  function renameVersion(version: WorkbookVersion, label: string | null) {
    return versionAction(async () => {
      const { versions, workbookId } = versionStore();
      if (!versions.rename) throw new Error('This workbook cannot rename versions.');
      const renamed = unwrap(
        await versions.rename(workbookId, version.revision, label),
        'Could not rename the version.',
      );
      if (alive.current)
        setViewedVersion((viewed) =>
          viewed?.version.revision === renamed.revision ? { ...viewed, version: renamed } : viewed,
        );
      versionList.refresh();
    }, 'Could not rename the version.');
  }
  /** Makes `version` the workbook's current content and returns to it. */
  function restoreVersion(version: WorkbookVersion) {
    return versionAction(async () => {
      const { versions, workbookId } = versionStore();
      if (!versions.restore) throw new Error('This workbook cannot restore versions.');
      if (!viewedVersion) await settleChanges();
      unwrap(
        await versions.restore(workbookId, version.revision),
        'Could not restore the version.',
      );
      await live.current.props.session?.takeStored();
      if (alive.current) returnToCurrent();
      versionList.refresh();
    }, 'Could not restore the version.');
  }
  function validateSheetName(id: string, name: string): string | null {
    try {
      validateWorksheetName(workbook, id, name);
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : 'Choose a valid worksheet name.';
    }
  }
  function renameSheet(id: string, name: string): boolean {
    if (!canEdit || !workbook.worksheets.some((value) => value.id === id)) return false;
    return mutateWithDraft((base) => renameWorksheet(base, id, name), 'Rename worksheet');
  }
  function sheetCommand(
    command: 'add' | 'insert' | 'duplicate' | 'delete' | 'left' | 'right',
    targetSheetId = sheet.id,
  ): boolean {
    if (!canEdit || !workbook.worksheets.some((value) => value.id === targetSheetId)) return false;
    let activatedId: string | undefined;
    const changed = mutateWithDraft((base) => {
      let next = copyWorkbook(base);
      const targetSheet = next.worksheets.find((value) => value.id === targetSheetId)!;
      if (command === 'add' || command === 'insert' || command === 'duplicate') {
        next = appendWorksheet(base);
        const added = next.worksheets.at(-1)!;
        added.id = (props.idFactory ?? createWorkbookEntityId)('sheet');
        if (command === 'insert') {
          next.worksheets.pop();
          next.worksheets.splice(
            next.worksheets.findIndex((value) => value.id === targetSheetId),
            0,
            added,
          );
        }
        if (command === 'duplicate')
          next.worksheets[next.worksheets.length - 1] = {
            ...structuredClone(targetSheet),
            id: added.id,
            name: added.name,
          };
        activatedId = added.id;
      }
      if (command === 'delete') {
        const index = next.worksheets.findIndex((value) => value.id === targetSheetId);
        next = deleteWorksheet(base, targetSheetId);
        if (targetSheetId === sheet.id)
          activatedId = next.worksheets[Math.min(index, next.worksheets.length - 1)].id;
      }
      if (command === 'left' || command === 'right') {
        const index = next.worksheets.findIndex((value) => value.id === targetSheetId),
          target =
            (index + (command === 'left' ? -1 : 1) + next.worksheets.length) %
            next.worksheets.length;
        [next.worksheets[index], next.worksheets[target]] = [
          next.worksheets[target],
          next.worksheets[index],
        ];
      }
      return next;
    }, `${command} worksheet`);
    if (changed && activatedId) {
      setSheetId(activatedId);
      setSelection(INITIAL_SELECTION);
    }
    return changed;
  }
  function openDialog(kind: Exclude<EditorDialog, null>) {
    setDialog(kind);
    setDialogValue(
      kind === 'row-height'
        ? String(sheet.rows?.[selection.y1]?.height ?? 26)
        : kind === 'column-width'
          ? String(sheet.columns?.[selection.x1]?.width ?? 100)
          : kind === 'export'
            ? JSON.stringify(workbook, null, 2)
            : '',
    );
    grid.current?.releaseKeyboard();
  }
  const dialogError =
    dialog === 'row-height' || dialog === 'column-width'
      ? !Number.isFinite(Number(dialogValue)) ||
        Number(dialogValue) <= 0 ||
        Number(dialogValue) > 2000
        ? 'Enter a positive size up to 2000 pixels.'
        : null
      : null;
  function applyDialog() {
    if (dialogError) return;
    if (dialog === 'confirm-structure') {
      pendingStructure.current?.();
      pendingStructure.current = null;
      setDialog(null);
      return;
    }
    if (dialog === 'export') {
      setDialog(null);
      return;
    }
    if (!canEdit) return;
    if (dialog === 'import') {
      const value = dialogValue.trimStart().startsWith('{')
        ? (JSON.parse(dialogValue) as Workbook)
        : createWorkbook(parseDelimited(dialogValue, ','));
      assertWorkbook(value);
      replaceWorkbook(value, 'Import workbook');
    }
    if (dialog === 'row-height' || dialog === 'column-width') {
      const measure = Number(dialogValue);
      if (!Number.isFinite(measure) || measure <= 0 || measure > 2000)
        throw new RangeError('Enter a positive size up to 2000 pixels.');
      mutateWithDraft(
        (base) => {
          const next = copyWorkbook(base),
            target = next.worksheets.find((value) => value.id === sheet.id)!;
          if (dialog === 'row-height') {
            target.rows = { ...target.rows };
            for (let y = selection.y1; y <= selection.y2; y++) target.rows[y] = { height: measure };
          } else {
            target.columns = [...(target.columns ?? [])];
            while (target.columns.length <= selection.x2) target.columns.push({});
            for (let x = selection.x1; x <= selection.x2; x++)
              target.columns[x] = { ...target.columns[x], width: measure };
          }
          return next;
        },
        dialog === 'row-height' ? 'Set row height' : 'Set column width',
      );
    }
    setDialog(null);
  }
  const matches = useMemo(() => {
    if (!query) return [];
    const matches: { sheetId: string; x: number; y: number; authored: boolean }[] = [];
    workbook.worksheets.forEach((sheet) =>
      sheet.data.forEach((row, y) =>
        row.forEach((value, x) => {
          const authored = String(value ?? '')
            .toLowerCase()
            .includes(query.toLowerCase());
          const key = calculationCellKey(sheet.id, x, y),
            ready = readyCells.has(key);
          const rawResult =
            shown.result?.worksheets.find((entry) => entry.id === sheet.id)?.values[y]?.[x] ?? null;
          const cellPresentation = presentation?.cells[key];
          const text = ready
            ? (cellPresentation?.text ??
              formatWorksheetCell(
                rawResult,
                sheet.formats?.[formatSpreadsheetAddress(x, y)] ??
                  (cellPresentation?.defaultUnit
                    ? { kind: 'unit', unit: cellPresentation.defaultUnit }
                    : undefined),
                props.display,
              ))
            : '';
          if (authored || text.toLowerCase().includes(query.toLowerCase()))
            matches.push({ sheetId: sheet.id, x, y, authored });
        }),
      ),
    );
    return matches;
  }, [workbook, query, shown.result, readyCells, presentation, props.display]);
  const matchIndex =
    activeMatch === null
      ? -1
      : matches.findIndex(
          (match) => calculationCellKey(match.sheetId, match.x, match.y) === activeMatch,
        );
  function navigateMatch(direction: number) {
    if (!matches.length) return;
    const current = matchIndex >= 0 ? matchIndex : direction < 0 ? 0 : -1;
    const index = (current + direction + matches.length) % matches.length,
      match = matches[index];
    if (select({ x1: match.x, x2: match.x, y1: match.y, y2: match.y }, match.sheetId))
      setActiveMatch(calculationCellKey(match.sheetId, match.x, match.y));
  }
  function replace(all: boolean, text: string) {
    if (!canWrite || !query) return;
    const targets = (
      all ? matches : matches[matchIndex] ? [matches[matchIndex]] : matches.slice(0, 1)
    ).filter((match) => match.authored);
    if (!targets.length) return 0;
    mutateWithDraft(
      (base) => {
        // Matches are authored cells, so every replacement lands in one copy (SE-8).
        const next = copyWorkbook(base),
          pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), all ? 'gi' : 'i');
        for (const match of targets) {
          const target = next.worksheets.find((sheet) => sheet.id === match.sheetId)!;
          target.data[match.y][match.x] = String(target.data[match.y][match.x] ?? '').replace(
            pattern,
            () => text,
          );
          if (target.inputValues)
            delete target.inputValues[formatSpreadsheetAddress(match.x, match.y)];
        }
        assertWorkbook(next);
        return next;
      },
      all ? 'Replace all matches' : 'Replace match',
      undefined,
      targets.reduce<CellTargets>((cells, match) => {
        const list = (cells.get(match.sheetId) as string[] | undefined) ?? [];
        list.push(formatSpreadsheetAddress(match.x, match.y));
        return cells.set(match.sheetId, list);
      }, new Map()),
    );
    return targets.length;
  }
  const summary = useMemo(() => {
    const values = selectionCells(selection).map(({ x, y }) => output?.[y]?.[x] ?? null);
    const { count, sum, average } = summarizeCellValues(values);
    return {
      count,
      sum: sum ? formatCellValue(sum, props.display) : '—',
      average: average ? formatCellValue(average, props.display) : '—',
    };
  }, [selection, output, props.display]);
  const lifecycleState: SpreadsheetEditorState = {
    dirty: !!session.status?.dirty || draftDirty || !!nativeDraft,
    draftDirty: draftDirty || !!nativeDraft,
    saving: savePhase === 'saving',
    flushing,
    error: noticeTone === 'danger' ? notice : null,
    calculation: {
      inputKey: visibleCalculation.inputKey,
      revision: visibleCalculation.revision,
      status: visibleCalculation.status,
    },
  };
  const lifecycle = useRef({
    setHostPanel,
    restoreDraft,
    commitDraft,
    cancelDraft,
    save,
    flush,
    select,
    draftTarget: { sheetId: sheet.id, x: selection.x1, y: selection.y1 },
    state: lifecycleState,
  });
  lifecycle.current = {
    setHostPanel,
    restoreDraft,
    commitDraft,
    cancelDraft,
    save,
    flush,
    select,
    draftTarget: { sheetId: sheet.id, x: selection.x1, y: selection.y1 },
    state: lifecycleState,
  };
  const handle = useMemo<SpreadsheetEditorHandle>(() => {
    const requireSession = () => {
      if (!alive.current) throw new Error('The editor session has ended.');
    };
    return {
      setPanel(id) {
        requireSession();
        lifecycle.current.setHostPanel(id);
      },
      restoreDraft(draft) {
        requireSession();
        lifecycle.current.restoreDraft(draft);
      },
      commitDraft() {
        requireSession();
        return lifecycle.current.commitDraft();
      },
      cancelDraft() {
        requireSession();
        lifecycle.current.cancelDraft();
      },
      getSnapshot() {
        requireSession();
        return copyWorkbook(latestWorkbook.current);
      },
      getState() {
        requireSession();
        const pending = draftRef.current;
        const current = lifecycle.current.state;
        const target = lifecycle.current.draftTarget;
        const draftDirty =
          !!nativeDraftRef.current ||
          (!!pending &&
            pending.value !==
              String(
                latestWorkbook.current.worksheets.find((value) => value.id === target.sheetId)
                  ?.data[target.y]?.[target.x] ?? '',
              ));
        return structuredClone({
          ...current,
          dirty: !!live.current.props.session?.status().dirty || draftDirty,
          draftDirty,
          saving: live.current.props.session?.status().phase === 'saving',
          flushing: current.flushing,
          error: errorNotice.current,
        });
      },
      select(target) {
        requireSession();
        const range = parseSpreadsheetRange(target.range);
        if (!range) throw new TypeError('Invalid selection range.');
        if (
          !lifecycle.current.select(
            { x1: range.start.x, y1: range.start.y, x2: range.end.x, y2: range.end.y },
            target.sheetId,
          )
        )
          throw new Error('Could not select the requested cell.');
      },
      subscribe(listener) {
        requireSession();
        listeners.current.add(listener);
        return () => listeners.current.delete(listener);
      },
      async save() {
        requireSession();
        await lifecycle.current.save();
      },
      async flush() {
        requireSession();
        await lifecycle.current.flush();
      },
    };
  }, []);
  useImperativeHandle(props.editorRef, () => handle, [handle]);
  const stateIdentity = JSON.stringify(lifecycleState);
  useLayoutEffect(() => {
    for (const listener of listeners.current) listener();
  }, [stateIdentity]);
  useEffect(() => {
    if (!workbook.worksheets.some((value) => value.id === sheetId)) {
      grid.current?.releaseKeyboard();
      setSheetId(workbook.worksheets[0].id);
      setSelection(
        sheetViews.current.get(workbook.worksheets[0].id)?.selection ?? INITIAL_SELECTION,
      );
    }
    const ids = new Set(workbook.worksheets.map((value) => value.id));
    for (const id of sheetViews.current.keys()) if (!ids.has(id)) sheetViews.current.delete(id);
  }, [workbook, sheetId]);
  return {
    props,
    workbook,
    sheet,
    selection,
    select,
    onGridSelection,
    activateSheet,
    mode,
    setMode,
    peers,
    lockedSelection,
    pasteCells,
    followCollaborator,
    publishPresence,
    capabilities,
    locked,
    canEdit,
    canSuggest,
    canWrite,
    canComment,
    canDecide,
    actor,
    personName,
    stamp,
    draftKey,
    draftValue,
    setDraftValue,
    draftDirty,
    draftCell,
    readyCells,
    presentation,
    initialViewport: sheetViews.current.get(sheet.id)?.viewport,
    setNativeDraft,
    handle,
    lifecycleState,
    applyFormula,
    flushDraft: applyFormula,
    cancelFormula: () => setDraft(null),
    calculation,
    visibleCalculation,
    publication,
    previewSheet,
    output,
    style,
    cellFormat,
    summary,
    grid,
    updateSheet,
    proposeCells,
    commit,
    mutate,
    mutateWithDraft,
    replaceWorkbook,
    attempt,
    undo,
    redo,
    history: currentHistory,
    canUndo: props.undo?.canUndo ?? !!currentHistory.past.length,
    canRedo: props.undo?.canRedo ?? !!currentHistory.future.length,
    format,
    toggleFormat,
    border,
    structure,
    copy,
    paste,
    save,
    flush,
    importExcel,
    refresh,
    exportExcel,
    savePhase,
    saveAutomatic: !!session.status?.automatic,
    sessionStatus: session.status,
    sessionDraft: session.draft,
    resolveConflict,
    versionList,
    viewedVersion: viewedVersion?.version ?? null,
    canSaveVersion: !viewedVersion && hostCapabilities.saveVersion && !!props.versions?.create,
    canRenameVersion: hostCapabilities.saveVersion && !!props.versions?.rename,
    canRestoreVersion: hostCapabilities.edit && !!props.versions?.restore,
    openVersion,
    returnToCurrent,
    saveVersion,
    renameVersion,
    restoreVersion,
    panel,
    panelOpen,
    setPanelOpen,
    openPanel,
    activeCommentId,
    setActiveCommentId,
    openComment,
    openSuggestion,
    replyDrafts,
    setReplyDrafts,
    commentScope,
    beginWorkbookComment,
    commentComposer,
    setCommentComposer,
    beginComment,
    postComment,
    decide,
    showContributors,
    setShowContributors,
    showFind,
    setShowFind,
    query,
    setQuery,
    replacement,
    setReplacement,
    matches,
    matchIndex,
    navigateMatch,
    replace,
    zoom,
    setZoom,
    showGridlines,
    setShowGridlines,
    showFormulas,
    setShowFormulas,
    notice,
    noticeTone,
    noticeRevision,
    warn,
    setNotice,
    dialog,
    dialogValue,
    dialogError,
    setDialogValue,
    setDialog,
    openDialog,
    applyDialog,
    sheetCommand,
    renameSheet,
    validateSheetName,
    insert,
    clear: (formatting = false) =>
      mutateWithDraft(
        (base) => clearSelection(base, sheet.id, selection, formatting),
        formatting ? 'Clear formatting' : 'Clear cells',
      ),
    merge: (unmerge = false) => {
      if (!canEdit) return;
      mutateWithDraft(
        (base) => mergeSelection(base, sheet.id, selection, unmerge),
        unmerge ? 'Unmerge cells' : 'Merge cells',
      );
    },
  };
}

export type SpreadsheetEditorController = ReturnType<typeof useSpreadsheetEditor>;
