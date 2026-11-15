'use client';

import { useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import { formatCellValue } from '../calculation/format.js';
import { calculationCellKey } from '../calculation/publication.js';
import { summarizeCellValues } from '../calculation/summary.js';
import { createDocumentInsertion } from '../document/insertion.js';
import type { WorkbookVersion } from '../host/adapters.js';
import {
  FULL_CAPABILITIES,
  READ_ONLY_CAPABILITIES,
  resolveSpreadsheetMode,
} from '../host/capabilities.js';
import {
  createSpreadsheetCopy,
  pasteSpreadsheetCopy,
  type SpreadsheetCopy,
} from '../model/clipboard.js';
import { parseDelimited, serializeDelimited } from '../model/delimited.js';
import {
  borderCells,
  type CellSelection,
  clearSelection,
  formatCells,
  INITIAL_SELECTION,
  mergeSelection,
  normalizeSelection,
  readSelection,
  selectionAddress,
  selectionCells,
  writeCells,
} from '../model/editing.js';
import { createWorkbookEntityId } from '../model/extent.js';
import {
  changeWorkbookStructureWithOperations,
  collectCellOperations,
  type WorkbookChangeContext,
  type WorkbookCommandResult,
  type WorkbookOperation,
} from '../model/operations.js';
import { collaboratorSheet, lockedChange, placePeers, selectionLock } from '../model/presence.js';
import { attributeCellChanges } from '../model/provenance.js';
import { requireDefined } from '../model/require-defined.js';
import { type CellFormat, proposeWorkbookChange, type ReviewStamp } from '../model/review.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import {
  renameWorksheet,
  type StructureChange,
  validateWorksheetName,
} from '../model/structure.js';
import {
  assertWorkbook,
  type CellInput,
  copyWorkbook,
  createWorkbook,
  type Workbook,
  type Worksheet,
} from '../model/workbook.js';
import { commandOperations } from './command-operations.js';
import type { GridController, SpreadsheetEditorProps, SpreadsheetMode } from './contracts.js';
import { blockTargets, type CellTargets, immutableBaseline, mergeTargets } from './edit-targets.js';
import { createEditorHandle, type EditorLifecycle } from './editor-handle.js';
import { type EditorHistory, historyEntry, redoStep, undoStep } from './history.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorState } from './public-host.js';
import { applySheetCommand, resizeSelection, type SheetCommand } from './sheet-commands.js';
import { useCellDisplay } from './use-cell-display.js';
import { useComments } from './use-comments.js';
import { matchTargets, replaceMatches, useFind } from './use-find.js';
import { useNotice } from './use-notice.js';
import { usePersonNames } from './use-person-names.js';
import { useSessionActions } from './use-session-actions.js';
import { sessionErrorMessage, sessionSavePhase, useSessionState } from './use-session-state.js';
import { useVersionActions } from './use-version-actions.js';
import { useWorkbookPresence } from './use-workbook-presence.js';
import { useWorkbookVersions } from './use-workbook-versions.js';

type EditorDialog =
  'row-height' | 'column-width' | 'import' | 'export' | 'confirm-structure' | null;
export function useSpreadsheetEditor(props: SpreadsheetEditorProps) {
  const [local, setLocal] = useState(() => copyWorkbook(props.defaultValue ?? createWorkbook()));
  const session = useSessionState(props.session ?? undefined);
  /** A saved version the reader opened, shown read-only in place of the workbook (SE-ADAPTERS-5.1). */
  const [viewedVersion, setViewedVersion] = useState<{
    version: WorkbookVersion;
    content: Workbook;
  } | null>(null);
  const workbook = viewedVersion?.content ?? props.value ?? session.workbook ?? local;
  const [history, setHistory] = useState<EditorHistory>({ head: null, past: [], future: [] });
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
  const lockedSelection = selectionLock(workbook, { sheetId: sheet.id, selection }, peers);
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
  function attachGrid(controller: GridController | null) {
    grid.current = controller;
  }
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
  const [showContributors, setShowContributors] = useState(false);
  const [zoom, setZoom] = useState(100),
    [showGridlines, setShowGridlines] = useState(true),
    [showFormulas, setShowFormulas] = useState(false);
  const { notice, noticeTone, noticeRevision, errorNotice, setNotice, warn } = useNotice();
  const pendingStructure = useRef<(() => void) | null>(null);
  const [dialog, setDialog] = useState<EditorDialog>(null),
    [dialogValue, setDialogValue] = useState('');
  const listeners = useRef(new Set<() => void>());
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
    const subscribed = listeners.current;
    return () => {
      alive.current = false;
      subscribed.clear();
    };
  }, []);
  const latestWorkbook = useRef(workbook);
  const live = useRef({ canEdit, canWrite, props });
  useLayoutEffect(() => {
    live.current = { canEdit, canWrite, props };
  }, [canEdit, canWrite, props]);
  // Read the callback through `live` so a host's inline callback is not a dependency (SE-HOST-2.1).
  useEffect(() => {
    live.current.props.onPanelChange?.({ id: panel, expanded: panelOpen });
  }, [panel, panelOpen]);
  useEffect(() => {
    latestWorkbook.current = workbook;
  }, [workbook]);
  const [refreshVersion, setRefreshVersion] = useState(0);
  const {
    authored,
    calculation,
    cellFormat,
    draftCell,
    draftDirty,
    draftKey,
    draftValue,
    output,
    presentation,
    previewSheet,
    previewWorkbook,
    publication,
    readyCells,
    shown,
    style,
    visibleCalculation,
  } = useCellDisplay({
    props,
    workbook,
    sheet,
    selection,
    draft,
    nativeDraft,
    canSuggest,
    actorId: actor.id,
    refreshVersion,
  });

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
      // The detached, frozen copy is made only for a host that reads it (SE-PERF-2.1).
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
      base = writeCells(base, [[pending.value]], {
        sheetId: sheet.id,
        x: selection.x1,
        y: selection.y1,
        idFactory: props.idFactory,
      });
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
    const host = props.undo;
    if (host) {
      attempt(() => host.undo());
      return;
    }
    const step = undoStep(currentHistory, workbook);
    if (!step) return;
    publish(step.target, step.operations, 'undo');
    setDraft(null);
    setHistory(step.history);
  }
  function redo() {
    if (!canWrite) return;
    const host = props.undo;
    if (host) {
      attempt(() => host.redo());
      return;
    }
    const step = redoStep(currentHistory, workbook);
    if (!step) return;
    publish(step.target, step.operations, 'redo');
    setDraft(null);
    setHistory(step.history);
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
  function proposeCells(values: CellInput[][], x: number, y: number) {
    const changed = mutate(
      writeCells(workbook, values, { sheetId: sheet.id, x, y, idFactory: props.idFactory }),
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
      (base) => formatCells(base, { sheetId: sheet.id, range: selection }, { properties, format }),
      `Format ${selectionAddress(selection)}`,
      undefined,
      new Map([[sheet.id, selectionCells(selection).map((cell) => cell.address)]]),
    );
  }
  function toggleFormat(property: string, value: string) {
    format({ [property]: style[property] === value ? null : value });
  }
  function border(kind: 'all' | 'outside' | 'none' | 'top' | 'right' | 'bottom' | 'left') {
    mutateWithDraft(
      (base) => borderCells(base, { sheetId: sheet.id, range: selection }, kind),
      `Set ${kind} borders`,
    );
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
        (base) =>
          changeWorkbookStructureWithOperations(
            base,
            { sheetId: sheet.id, idFactory: props.idFactory },
            change,
          ),
        `${kind === 'insert' ? 'Insert' : 'Delete'} ${axis}${last > first ? 's' : ''}`,
        { sheetId: sheet.id, change },
      );
    const crossesMerge = Object.entries(sheet.mergeCells ?? {}).some(
      ([address, [columns, rows]]) => {
        const range = requireDefined(parseSpreadsheetRange(address));
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
  const {
    commentComposer,
    setCommentComposer,
    commentScope,
    activeCommentId,
    setActiveCommentId,
    replyDrafts,
    setReplyDrafts,
    openComment,
    beginComment,
    beginWorkbookComment,
    postComment,
    openSuggestion,
    decide,
  } = useComments({
    workbook,
    sheet,
    selection,
    initialCommentId: props.initialCommentId,
    canComment,
    canDecide,
    grid,
    select,
    openPanel,
    commit,
    stamp,
  });
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
        ? writeCells(previewWorkbook, [[pending.value]], {
            sheetId: sheet.id,
            x: selection.x1,
            y: selection.y1,
            idFactory: props.idFactory,
          })
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
      copied.current = createSpreadsheetCopy(
        source,
        { sheetId: sheet.id, range: selection },
        { text, cut: false },
      );
      warn('The cell edit changed while copying. Nothing was cut. Try Cut again.');
      return;
    }
    copied.current = createSpreadsheetCopy(
      source,
      { sheetId: sheet.id, range: selection },
      { text, cut },
    );
  }
  function pasteCells(
    values: CellInput[][],
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
          ? pasteSpreadsheetCopy(base, cached, {
              sheetId: sheet.id,
              x,
              y,
              destination: selection,
              idFactory: props.idFactory,
            })
          : writeCells(base, values, { sheetId: sheet.id, x, y, idFactory: props.idFactory }),
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
  const { flushing, save, flush, resolveConflict } = useSessionActions({
    commitDraft,
    live,
    alive,
    setNotice,
    warn,
  });
  const { openVersion, returnToCurrent, saveVersion, renameVersion, restoreVersion } =
    useVersionActions({
      viewedVersion,
      setViewedVersion,
      versionList,
      commitDraft,
      flush: () => lifecycle.current.flush(),
      props: () => live.current.props,
      alive,
      setNotice,
    });
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
  function sheetCommand(command: SheetCommand, targetSheetId = sheet.id): boolean {
    if (!canEdit || !workbook.worksheets.some((value) => value.id === targetSheetId)) return false;
    let activatedId: string | undefined;
    const changed = mutateWithDraft((base) => {
      const result = applySheetCommand(base, command, targetSheetId, sheet.id, props.idFactory);
      activatedId = result.activatedId;
      return result.workbook;
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
        (base) => resizeSelection(base, sheet.id, selection, dialog, measure),
        dialog === 'row-height' ? 'Set row height' : 'Set column width',
      );
    }
    setDialog(null);
  }
  const {
    showFind,
    setShowFind,
    query,
    setQuery,
    replacement,
    setReplacement,
    setActiveMatch,
    matches,
    matchIndex,
  } = useFind(workbook, shown.result, readyCells, presentation, props.display);
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
      (base) => replaceMatches(base, targets, query, text, all),
      all ? 'Replace all matches' : 'Replace match',
      undefined,
      matchTargets(targets),
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
    error: sessionErrorMessage(session.status) ?? (noticeTone === 'danger' ? notice : null),
    calculation: {
      inputKey: visibleCalculation.inputKey,
      revision: visibleCalculation.revision,
      status: visibleCalculation.status,
    },
  };
  const lifecycle = useRef<EditorLifecycle>({
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
  const handle = useMemo<SpreadsheetEditorHandle>(
    () =>
      createEditorHandle({
        alive,
        lifecycle,
        latestWorkbook,
        draftRef,
        nativeDraftRef,
        session: () => live.current.props.session,
        errorNotice,
        listeners,
      }),
    [errorNotice],
  );
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
    attachGrid,
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
        (base) => clearSelection(base, { sheetId: sheet.id, range: selection }, formatting),
        formatting ? 'Clear formatting' : 'Clear cells',
      ),
    merge: (unmerge = false) => {
      if (!canEdit) return;
      mutateWithDraft(
        (base) => mergeSelection(base, { sheetId: sheet.id, range: selection }, unmerge),
        unmerge ? 'Unmerge cells' : 'Merge cells',
      );
    },
  };
}

export type SpreadsheetEditorController = ReturnType<typeof useSpreadsheetEditor>;
