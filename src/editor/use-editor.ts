'use client';

import {
  remoteCollaborators,
  collaboratorSheet,
  lockedChange,
  selectionLock,
} from '../model/presence.js';
import { attributeCellChanges } from '../model/provenance.js';

import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
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
import { createArithmetic } from '../calculation/rational.js';
import { isNumeric } from '../calculation/values.js';
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
  changeWorkbookStructure,
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
import { useCalculation } from './use-calculation.js';

export type EditorDialog =
  'row-height' | 'column-width' | 'import' | 'export' | 'confirm-structure' | null;
export function useSpreadsheetEditor(props: SpreadsheetEditorProps) {
  const [local, setLocal] = useState(() => copyWorkbook(props.defaultValue ?? createWorkbook()));
  const workbook = props.value ?? local;
  const [history, setHistory] = useState<{
    head: Workbook | null;
    past: { workbook: Workbook; label: string }[];
    future: { workbook: Workbook; label: string }[];
  }>({ head: null, past: [], future: [] });
  const currentHistory =
    history.head === workbook ? history : { head: workbook, past: [], future: [] };
  const [localMode, setLocalMode] = useState<SpreadsheetMode>(props.defaultMode ?? 'author');
  const requestedMode = props.mode ?? localMode;
  const mode: SpreadsheetMode = props.readOnly
    ? 'read'
    : props.permissions?.edit === false
      ? props.permissions.suggest === false || props.features?.suggestions === false
        ? 'read'
        : 'review'
      : requestedMode;
  const canEdit = !props.readOnly && mode === 'author' && props.permissions?.edit !== false;
  const canSuggest =
    !props.readOnly &&
    mode === 'review' &&
    props.permissions?.suggest !== false &&
    props.features?.suggestions !== false;
  const canWrite = canEdit || canSuggest;
  const canComment =
    !props.readOnly &&
    mode !== 'read' &&
    props.permissions?.comment !== false &&
    props.features?.comments !== false;
  const canDecide =
    canEdit &&
    props.permissions?.decideSuggestions !== false &&
    props.features?.suggestions !== false;
  const actor = props.currentUser ?? { id: 'local-user', name: 'You' };
  const peers = useMemo(
    () =>
      remoteCollaborators(props.collaborators ?? [], actor.id).map((person) => ({
        ...person,
        sheetId: collaboratorSheet(workbook, person)?.id ?? person.sheetId,
      })),
    [props.collaborators, actor.id, workbook],
  );
  const [sheetId, setSheetId] = useState(props.initialSheetId ?? workbook.worksheets[0].id);
  const sheet = workbook.worksheets.find((sheet) => sheet.id === sheetId) ?? workbook.worksheets[0];
  const [selection, setSelection] = useState<CellSelection>(INITIAL_SELECTION);
  const lockedSelection = selectionLock(workbook, sheet.id, selection, peers);
  const [draft, setDraftState] = useState<{ key: string; value: string } | null>(null);
  const draftRef = useRef(draft);
  function setDraft(next: typeof draft) {
    draftRef.current = next;
    setDraftState(next);
  }
  const grid = useRef<GridController | null>(null);
  const copied = useRef<SpreadsheetCopy | null>(null);
  const [panel, setPanel] = useState<string | null>(props.historical ? 'history' : null),
    [panelOpen, setPanelOpen] = useState(!!props.historical);
  const [commentComposer, setCommentComposer] = useState(false);
  const [commentScope, setCommentScope] = useState<'range' | 'workbook'>('range');
  const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
  const [replyDrafts, setReplyDrafts] = useState<Record<string, string>>({});
  const [showContributors, setShowContributors] = useState(false);
  const [showFind, setShowFind] = useState(false),
    [query, setQuery] = useState(''),
    [replacement, setReplacement] = useState('');
  const [matchIndex, setMatchIndex] = useState(-1);
  const [zoom, setZoom] = useState(100),
    [showGridlines, setShowGridlines] = useState(true),
    [showFormulas, setShowFormulas] = useState(false);
  const [notice, setNoticeState] = useState<string | null>(null);
  const [noticeTone, setNoticeTone] = useState<'danger' | 'warning'>('danger');
  const [noticeRevision, setNoticeRevision] = useState(0);
  function setNotice(message: string | null) {
    setNoticeTone('danger');
    setNoticeState(message);
  }
  function warn(message: string) {
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
  const [savePhase, setSavePhase] = useState<'idle' | 'dirty' | 'saving' | 'saved' | 'error'>(
    'idle',
  );
  const alive = useRef(true);
  useEffect(() => {
    alive.current = true;
    return () => {
      alive.current = false;
    };
  }, []);
  const latestWorkbook = useRef(workbook);
  const live = useRef({ canEdit, canWrite, props });
  useLayoutEffect(() => {
    live.current = { canEdit, canWrite, props };
  }, [canEdit, canWrite, props]);
  const activeSave = useRef<{
    promise: Promise<void>;
    controller: AbortController;
    kind: 'manual' | 'autosave';
  } | null>(null);
  const lastSaveError = useRef<unknown>(null);
  const [saveKind, setSaveKind] = useState<'manual' | 'autosave'>('manual');
  useEffect(() => {
    if (!canEdit) {
      activeSave.current?.controller.abort();
      setSavePhase((phase) => (phase === 'saving' ? 'dirty' : phase));
    }
    return () => {
      activeSave.current?.controller.abort();
    };
  }, [canEdit]);
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
      ...props.calculation?.options,
    }),
    [props.calculation?.options],
  );
  const calculation = useCalculation(workbook, provider, calculationOptions);
  const previewWorkbook = useMemo(
    () => (canSuggest ? previewSuggestions(workbook, actor.id) : workbook),
    [workbook, actor.id, canSuggest],
  );
  const previewCalculation = useCalculation(
    previewWorkbook,
    provider,
    calculationOptions,
    previewWorkbook !== workbook,
  );
  const visibleCalculation = previewWorkbook !== workbook ? previewCalculation : calculation;
  const previewSheet = previewWorkbook.worksheets.find((value) => value.id === sheet.id)!;
  const authored = previewSheet.data[selection.y1]?.[selection.x1] ?? '';
  const draftKey = `${sheet.id}:${selection.x1}:${selection.y1}`;
  const draftValue = draft?.key === draftKey ? draft.value : String(authored);
  const output = visibleCalculation.result?.worksheets.find(
    (result) => result.id === sheet.id,
  )?.values;
  const style = parseCellStyle(
    previewSheet.style?.[selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })],
  );
  const cellFormat = previewSheet.formats?.[
    selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })
  ] ?? { kind: 'general' as const };

  function publishPresence(activity: 'Editing' | 'Reviewing' | 'Viewing', editing = false) {
    props.onPresenceChange?.({
      sheetId: sheet.id,
      selection,
      editLock: editing ? selection : undefined,
      activity,
      timestamp: Date.now(),
    });
  }
  useEffect(() => {
    publishPresence('Viewing');
  }, [sheet.id, selection]);
  function followCollaborator(id: string) {
    const person = peers.find((person) => person.id === id),
      target = person && collaboratorSheet(workbook, person);
    if (person?.selection && target) select(person.selection, target.id);
  }
  function stamp(): ReviewStamp {
    return { id: globalThis.crypto.randomUUID(), time: new Date().toISOString(), actor };
  }
  function publish(next: Workbook) {
    assertWorkbook(next);
    latestWorkbook.current = next;
    if (!props.value) setLocal(next);
    props.onChange?.(next);
    setSavePhase('dirty');
  }
  function commit(next: Workbook, label = 'Edit workbook') {
    if (next === workbook) return;
    assertWorkbook(next);
    setHistory({
      head: next,
      past: [...currentHistory.past.slice(-49), { workbook: copyWorkbook(workbook), label }],
      future: [],
    });
    publish(next);
    publishPresence(/comment|reply|reaction|suggest/i.test(label) ? 'Reviewing' : 'Editing');
  }
  function mutate(
    next: Workbook,
    label: string,
    operation?: { sheetId: string; change: StructureChange },
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
    commit(
      canSuggest
        ? proposeWorkbookChange(workbook, next, label, stamp(), operation)
        : operation
          ? next
          : attributeCellChanges(workbook, next, {
              actor,
              source: 'human',
              at: new Date().toISOString(),
            }),
      canSuggest ? `Suggest: ${label}` : label,
    );
    if (canSuggest) {
      setPanel('suggestions');
      setPanelOpen(true);
    }
    return true;
  }
  /** Compose a protected formula draft and a chrome command into one history entry. */
  function mutateWithDraft(
    transform: (base: Workbook) => Workbook,
    label: string,
    operation?: { sheetId: string; change: StructureChange },
  ): boolean {
    const pending = draftRef.current;
    let base = workbook;
    if (pending?.key === draftKey && pending.value !== String(authored)) {
      base = writeCells(base, sheet.id, selection.x1, selection.y1, [[pending.value]]);
      const locked = lockedChange(workbook, base, peers);
      if (locked) {
        warn(`${locked.name} is editing a cell in this selection. Try again after they finish.`);
        return false;
      }
      // Structural transforms move this attribution together with the authored cell.
      if (canEdit)
        base = attributeCellChanges(workbook, base, {
          actor,
          source: 'human',
          at: new Date().toISOString(),
        });
    }
    if (!mutate(transform(base), label, operation)) return false;
    if (draftRef.current === pending && pending?.key === draftKey) setDraft(null);
    return true;
  }
  function replaceWorkbook(next: Workbook, label: string): boolean {
    if (!canEdit || !mutate(next, label)) return false;
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
    if (
      props.readOnly ||
      (props.permissions?.edit === false &&
        next !== (props.permissions.suggest === false ? 'read' : 'review'))
    )
      return;
    if (!applyFormula()) return;
    setLocalMode(next);
    props.onModeChange?.(next);
    setDraft(null);
    grid.current?.releaseKeyboard();
  }
  function select(range: CellSelection, targetSheetId = sheet.id) {
    const normalized = normalizeSelection(range);
    selectionCells(normalized);
    if (!workbook.worksheets.some((value) => value.id === targetSheetId) || !applyFormula())
      return false;
    setSheetId(targetSheetId);
    setSelection(normalized);
    setDraft(null);
    if (targetSheetId === sheet.id) grid.current?.select(normalized);
    return true;
  }
  function activateSheet(id: string) {
    if (id === sheet.id) return true;
    if (!workbook.worksheets.some((value) => value.id === id) || !applyFormula()) return false;
    setSheetId(id);
    setSelection(INITIAL_SELECTION);
    setDraft(null);
    return true;
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
    setDraft(null);
    const previous = currentHistory.past.at(-1);
    if (!previous) return;
    setHistory({
      head: previous.workbook,
      past: currentHistory.past.slice(0, -1),
      future: [
        { workbook: copyWorkbook(workbook), label: previous.label },
        ...currentHistory.future,
      ],
    });
    publish(previous.workbook);
  }
  function redo() {
    if (!canWrite) return;
    const next = currentHistory.future[0];
    if (!next) return;
    setDraft(null);
    setHistory({
      head: next.workbook,
      past: [...currentHistory.past, { workbook: copyWorkbook(workbook), label: next.label }],
      future: currentHistory.future.slice(1),
    });
    publish(next.workbook);
  }
  function applyFormula(): boolean {
    const pending = draftRef.current;
    if (pending?.key !== draftKey || pending.value === String(authored)) return true;
    return mutateWithDraft((base) => base, `Edit ${selectionAddress(selection)}`);
  }
  function setDraftValue(value: string) {
    setDraft({ key: draftKey, value });
  }
  function proposeCells(
    values: import('../model/workbook.js').CellInput[][],
    x: number,
    y: number,
  ) {
    const changed = mutate(writeCells(workbook, sheet.id, x, y, values), 'Edit cells');
    if (changed) setDraft(null);
    return changed;
  }
  function format(properties: Record<string, string | null>, format?: CellFormat | null) {
    mutateWithDraft(
      (base) => formatCells(base, sheet.id, selection, properties, format),
      `Format ${selectionAddress(selection)}`,
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
        (base) => changeWorkbookStructure(base, sheet.id, change),
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
    if (id === 'history') attempt(() => props.history?.onRefresh?.());
    grid.current?.releaseKeyboard();
  }
  function openComment(id: string) {
    const thread = workbook.comments?.find((thread) => thread.id === id);
    if (!thread) return;
    if (thread.scope === 'workbook' || thread.resolved || thread.orphaned) {
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
    if (activeCommentId !== 'draft' && (!thread || thread.resolved || thread.orphaned))
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
        ? writeCells(previewWorkbook, sheet.id, selection.x1, selection.y1, [[pending.value]])
        : previewWorkbook;
    const cells = readSelection(source, sheet.id, selection);
    const displayed = cells.map((row, dy) =>
      row.map((_cell, dx) => {
        const x = selection.x1 + dx,
          y = selection.y1 + dy;
        if (hasDraft && x === selection.x1 && y === selection.y1) return pending.value;
        return formatWorksheetCell(
          visibleCalculation.result?.worksheets.find((item) => item.id === sheet.id)?.values[y]?.[
            x
          ] ?? null,
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
    if (cut) mutateWithDraft((base) => clearSelection(base, sheet.id, selection), 'Cut cells');
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
          ? pasteSpreadsheetCopy(base, sheet.id, x, y, cached)
          : writeCells(base, sheet.id, x, y, values),
      'Paste cells',
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
    if (
      !props.onExportXlsx ||
      !calculation.result ||
      calculation.pending ||
      calculation.error ||
      draft?.key === draftKey
    )
      return;
    await props.onExportXlsx(copyWorkbook(workbook), structuredClone(calculation.result));
  }
  async function save(kind: 'manual' | 'autosave' = 'manual') {
    if (!live.current.props.onSave) return;
    if (!live.current.canEdit) {
      warn('Saving is available only in Authoring mode.');
      return;
    }
    if (kind === 'manual' && !applyFormula()) return;
    if (activeSave.current) {
      if (kind === 'manual' && activeSave.current.kind === 'autosave') {
        await activeSave.current.promise;
        if (live.current.canEdit) return save('manual');
      }
      return activeSave.current?.promise;
    }
    const snapshot = latestWorkbook.current,
      controller = new AbortController();
    const context = { signal: controller.signal, kind, requestId: globalThis.crypto.randomUUID() };
    setSaveKind(kind);
    setSavePhase('saving');
    lastSaveError.current = null;
    const operation = Promise.resolve().then(async () => {
      const started = Date.now();
      let retry = 500;
      try {
        while (true) {
          if (controller.signal.aborted || !live.current.canEdit) return;
          try {
            await live.current.props.onSave!(copyWorkbook(snapshot), context);
            break;
          } catch (error) {
            const transient =
              error && typeof error === 'object' && 'status' in error && error.status === 503;
            if (kind !== 'autosave' || !transient || Date.now() - started >= 30000) throw error;
            await new Promise<void>((resolve) => {
              const timer = setTimeout(resolve, retry);
              controller.signal.addEventListener(
                'abort',
                () => {
                  clearTimeout(timer);
                  resolve();
                },
                { once: true },
              );
            });
            retry = Math.min(4000, retry * 2);
          }
        }
        if (!controller.signal.aborted) {
          setSavePhase(latestWorkbook.current === snapshot ? 'saved' : 'dirty');
          setNotice(null);
        }
      } catch (error) {
        if (!controller.signal.aborted) {
          lastSaveError.current = error;
          setSavePhase('error');
          setNotice(error instanceof Error ? error.message : 'Could not save workbook.');
        }
      } finally {
        activeSave.current = null;
      }
    });
    activeSave.current = { promise: operation, controller, kind };
    return operation;
  }
  useEffect(() => {
    if (savePhase !== 'dirty' || !canEdit || !props.onSave || props.autosaveMs === undefined)
      return;
    if (!Number.isFinite(props.autosaveMs) || props.autosaveMs < 1) return;
    const timer = setTimeout(() => {
      void save('autosave');
    }, props.autosaveMs);
    return () => clearTimeout(timer);
  }, [workbook, savePhase, canEdit, props.autosaveMs]);
  async function openHistorical(entry: import('./contracts.js').SpreadsheetHistoryEntry) {
    const pending = draftRef.current,
      hadDraft = pending?.key === draftKey && pending.value !== String(authored);
    if (hadDraft && !applyFormula()) return;
    if (hadDraft || savePhase === 'dirty' || activeSave.current) {
      await save();
      if (lastSaveError.current) return;
    }
    if (alive.current) await props.history?.onOpen(entry);
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
    const matches: { sheetId: string; x: number; y: number }[] = [];
    workbook.worksheets.forEach((sheet) =>
      sheet.data.forEach((row, y) =>
        row.forEach((value, x) => {
          if (
            String(value ?? '')
              .toLowerCase()
              .includes(query.toLowerCase())
          )
            matches.push({ sheetId: sheet.id, x, y });
        }),
      ),
    );
    return matches;
  }, [workbook, query]);
  function navigateMatch(direction: number) {
    if (!matches.length) return;
    const index = (matchIndex + direction + matches.length) % matches.length,
      match = matches[index];
    setMatchIndex(index);
    select({ x1: match.x, x2: match.x, y1: match.y, y2: match.y }, match.sheetId);
  }
  function replace(all: boolean, text: string) {
    if (!canWrite || !query) return;
    const targets = all
      ? matches
      : matches[matchIndex]
        ? [matches[matchIndex]]
        : matches.slice(0, 1);
    mutateWithDraft(
      (base) => {
        let next = base;
        for (const match of targets) {
          const value = String(
            next.worksheets.find((sheet) => sheet.id === match.sheetId)!.data[match.y][match.x] ??
              '',
          );
          const pattern = new RegExp(
            query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
            all ? 'gi' : 'i',
          );
          next = writeCells(next, match.sheetId, match.x, match.y, [
            [value.replace(pattern, () => text)],
          ]);
        }
        return next;
      },
      all ? 'Replace all matches' : 'Replace match',
    );
  }
  const summary = useMemo(() => {
    const values = selectionCells(selection).map(({ x, y }) => output?.[y]?.[x] ?? null),
      numeric = values.filter(isNumeric);
    const arithmetic = createArithmetic();
    try {
      const sum = numeric.reduce(arithmetic.add, arithmetic.parse('0'));
      return {
        count: values.filter((value) => value !== null && value !== '').length,
        sum: formatCellValue(sum, props.display),
        average: numeric.length
          ? formatCellValue(
              arithmetic.divide(sum, arithmetic.parse(String(numeric.length))),
              props.display,
            )
          : '—',
      };
    } catch {
      return { count: values.length, sum: '—', average: '—' };
    }
  }, [selection, output, props.display]);
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
    canEdit,
    canSuggest,
    canWrite,
    canComment,
    canDecide,
    actor,
    stamp,
    draftKey,
    draftValue,
    setDraftValue,
    draftDirty: draft?.key === draftKey && draft.value !== String(authored),
    applyFormula,
    flushDraft: applyFormula,
    cancelFormula: () => setDraft(null),
    calculation,
    visibleCalculation,
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
    format,
    toggleFormat,
    border,
    structure,
    copy,
    paste,
    save,
    importExcel,
    refresh,
    exportExcel,
    savePhase,
    saveKind,
    openHistorical,
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
