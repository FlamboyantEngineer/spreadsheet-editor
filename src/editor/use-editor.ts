'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { createClientCalculator } from '../calculation/provider.js';
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
  type StructureChange,
} from '../model/structure.js';
import {
  addComment,
  decideSuggestion,
  proposeWorkbookChange,
  type CellFormat,
  type ReviewStamp,
} from '../model/review.js';
import { createDocumentInsertion } from '../document/insertion.js';
import type { GridController, SpreadsheetEditorProps, SpreadsheetMode } from './contracts.js';
import { useCalculation } from './use-calculation.js';

export type EditorDialog =
  'rename-sheet' | 'row-height' | 'column-width' | 'import' | 'export' | null;
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
  const mode: SpreadsheetMode = props.readOnly ? 'read' : (props.mode ?? localMode);
  const canEdit = !props.readOnly && mode === 'author' && props.permissions?.edit !== false;
  const canSuggest = !props.readOnly && mode === 'review' && props.permissions?.suggest !== false;
  const canWrite = canEdit || canSuggest;
  const canComment = !props.readOnly && mode !== 'read' && props.permissions?.comment !== false;
  const canDecide =
    !props.readOnly && mode !== 'read' && props.permissions?.decideSuggestions !== false;
  const actor = props.currentUser ?? { id: 'local-user', name: 'You' };
  const [sheetId, setSheetId] = useState(props.initialSheetId ?? workbook.worksheets[0].id);
  const sheet = workbook.worksheets.find((sheet) => sheet.id === sheetId) ?? workbook.worksheets[0];
  const [selection, setSelection] = useState<CellSelection>(INITIAL_SELECTION);
  const [draft, setDraft] = useState<{ key: string; value: string } | null>(null);
  const authored = sheet.data[selection.y1]?.[selection.x1] ?? '';
  const draftKey = `${sheet.id}:${selection.x1}:${selection.y1}:${String(authored)}`;
  const draftValue = draft?.key === draftKey ? draft.value : String(authored);
  const grid = useRef<GridController | null>(null);
  const [panel, setPanel] = useState('comments'),
    [panelOpen, setPanelOpen] = useState(false);
  const [commentComposer, setCommentComposer] = useState(false);
  const [showFind, setShowFind] = useState(false),
    [query, setQuery] = useState(''),
    [replacement, setReplacement] = useState('');
  const [matchIndex, setMatchIndex] = useState(-1);
  const [zoom, setZoom] = useState(100),
    [showGridlines, setShowGridlines] = useState(true),
    [showFormulas, setShowFormulas] = useState(false);
  const [notice, setNotice] = useState<string | null>(null);
  const [dialog, setDialog] = useState<EditorDialog>(null),
    [dialogValue, setDialogValue] = useState('');
  const [savePhase, setSavePhase] = useState<'idle' | 'dirty' | 'saving' | 'saved' | 'error'>(
    'idle',
  );
  const latestWorkbook = useRef(workbook);
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
  const provider = props.calculation?.mode === 'server' ? props.calculation.provider : client;
  const calculation = useCalculation(workbook, provider, props.calculation?.options);
  const output = calculation.result?.worksheets.find((result) => result.id === sheet.id)?.values;
  const style = parseCellStyle(
    sheet.style?.[selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })],
  );
  const cellFormat = sheet.formats?.[
    selectionAddress({ ...selection, x2: selection.x1, y2: selection.y1 })
  ] ?? { kind: 'general' as const };

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
    setNotice(null);
  }
  function mutate(
    next: Workbook,
    label: string,
    operation?: { sheetId: string; change: StructureChange },
  ) {
    if (!canWrite) return;
    commit(
      canSuggest ? proposeWorkbookChange(workbook, next, label, stamp(), operation) : next,
      canSuggest ? `Suggest: ${label}` : label,
    );
    if (canSuggest) {
      setPanel('suggestions');
      setPanelOpen(true);
    }
  }
  function attempt(action: () => void | Promise<void>) {
    try {
      const result = action();
      if (result)
        result.catch((error) =>
          setNotice(error instanceof Error ? error.message : 'The action failed.'),
        );
    } catch (error) {
      setNotice(error instanceof Error ? error.message : 'The action failed.');
    }
  }
  function setMode(next: SpreadsheetMode) {
    if (props.readOnly) return;
    setLocalMode(next);
    props.onModeChange?.(next);
    setDraft(null);
    grid.current?.releaseKeyboard();
  }
  function select(range: CellSelection, targetSheetId = sheet.id) {
    const normalized = normalizeSelection(range);
    selectionCells(normalized);
    setSheetId(targetSheetId);
    setSelection(normalized);
    setDraft(null);
    if (targetSheetId === sheet.id) grid.current?.select(normalized);
  }
  function activateSheet(id: string) {
    setSheetId(id);
    setSelection(INITIAL_SELECTION);
    setDraft(null);
  }
  function onGridSelection(range: CellSelection) {
    attempt(() => {
      const normalized = normalizeSelection(range);
      selectionCells(normalized);
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
    mutate(
      {
        ...workbook,
        worksheets: workbook.worksheets.map((sheet) => (sheet.id === next.id ? next : sheet)),
      },
      'Edit cells',
    );
  }
  function undo() {
    if (!canWrite) return;
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
    setHistory({
      head: next.workbook,
      past: [...currentHistory.past, { workbook: copyWorkbook(workbook), label: next.label }],
      future: currentHistory.future.slice(1),
    });
    publish(next.workbook);
  }
  function applyFormula() {
    if (canWrite && draft?.key === draftKey) {
      mutate(
        writeCells(workbook, sheet.id, selection.x1, selection.y1, [[draft.value]]),
        `Edit ${selectionAddress(selection)}`,
      );
      setDraft(null);
    }
  }
  function setDraftValue(value: string) {
    setDraft({ key: draftKey, value });
  }
  function format(properties: Record<string, string | null>, format?: CellFormat | null) {
    mutate(
      formatCells(workbook, sheet.id, selection, properties, format),
      `Format ${selectionAddress(selection)}`,
    );
  }
  function toggleFormat(property: string, value: string) {
    format({ [property]: style[property] === value ? null : value });
  }
  function border(kind: 'all' | 'outside' | 'none' | 'bottom') {
    mutate(borderCells(workbook, sheet.id, selection, kind), `Set ${kind} borders`);
  }
  function structure(axis: StructureChange['axis'], kind: StructureChange['kind'], after = false) {
    const first = axis === 'row' ? selection.y1 : selection.x1,
      last = axis === 'row' ? selection.y2 : selection.x2;
    const change = { axis, kind, index: after ? last + 1 : first, count: last - first + 1 };
    mutate(
      changeWorkbookStructure(workbook, sheet.id, change),
      `${kind === 'insert' ? 'Insert' : 'Delete'} ${axis}${last > first ? 's' : ''}`,
      { sheetId: sheet.id, change },
    );
  }
  function openPanel(id: string) {
    setPanel(id);
    setPanelOpen(true);
    grid.current?.releaseKeyboard();
  }
  function beginComment() {
    if (canComment) {
      openPanel('comments');
      setCommentComposer(true);
    }
  }
  function postComment(text: string) {
    if (canComment) {
      commit(
        addComment(workbook, sheet.id, selectionAddress(selection), text, stamp()),
        'Add comment',
      );
      setCommentComposer(false);
    }
  }
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
    await clipboard.writeText(serializeDelimited(readSelection(workbook, sheet.id, selection)));
    if (cut) mutate(clearSelection(workbook, sheet.id, selection), 'Cut cells');
  }
  async function paste() {
    if (canWrite) {
      const text = await clipboard.readText();
      mutate(
        writeCells(workbook, sheet.id, selection.x1, selection.y1, parseDelimited(text)),
        'Paste cells',
      );
    }
  }
  async function save() {
    if (!props.onSave || props.readOnly) return;
    const snapshot = workbook;
    setSavePhase('saving');
    try {
      await props.onSave(copyWorkbook(snapshot));
      setSavePhase(latestWorkbook.current === snapshot ? 'saved' : 'dirty');
    } catch (error) {
      setSavePhase('error');
      setNotice(error instanceof Error ? error.message : 'Could not save workbook.');
    }
  }
  function sheetCommand(command: 'add' | 'duplicate' | 'delete' | 'left' | 'right') {
    if (!canWrite) return;
    let next = copyWorkbook(workbook);
    if (command === 'add') {
      next = appendWorksheet(workbook);
      if (!canSuggest) activateSheet(next.worksheets.at(-1)!.id);
    }
    if (command === 'duplicate') {
      next = appendWorksheet(workbook);
      const added = next.worksheets.at(-1)!;
      next.worksheets[next.worksheets.length - 1] = {
        ...structuredClone(sheet),
        id: added.id,
        name: added.name,
      };
      if (!canSuggest) activateSheet(added.id);
    }
    if (command === 'delete') next = deleteWorksheet(workbook, sheet.id);
    if (command === 'left' || command === 'right') {
      const index = next.worksheets.findIndex((candidate) => candidate.id === sheet.id),
        target = index + (command === 'left' ? -1 : 1);
      if (target < 0 || target >= next.worksheets.length) return;
      [next.worksheets[index], next.worksheets[target]] = [
        next.worksheets[target],
        next.worksheets[index],
      ];
    }
    mutate(next, `${command} worksheet`);
  }
  function openDialog(kind: Exclude<EditorDialog, null>) {
    setDialog(kind);
    setDialogValue(
      kind === 'rename-sheet'
        ? sheet.name
        : kind === 'row-height'
          ? String(sheet.rows?.[selection.y1]?.height ?? 26)
          : kind === 'column-width'
            ? String(sheet.columns?.[selection.x1]?.width ?? 100)
            : kind === 'export'
              ? JSON.stringify(workbook, null, 2)
              : '',
    );
    grid.current?.releaseKeyboard();
  }
  function applyDialog() {
    if (dialog === 'export') {
      setDialog(null);
      return;
    }
    if (!canWrite) return;
    if (dialog === 'rename-sheet')
      mutate(renameWorksheet(workbook, sheet.id, dialogValue), 'Rename worksheet');
    if (dialog === 'import') {
      const value = dialogValue.trimStart().startsWith('{')
        ? (JSON.parse(dialogValue) as Workbook)
        : createWorkbook(parseDelimited(dialogValue, ','));
      assertWorkbook(value);
      mutate(value, 'Import workbook');
    }
    if (dialog === 'row-height' || dialog === 'column-width') {
      const measure = Number(dialogValue);
      if (!Number.isFinite(measure) || measure < 10 || measure > 2000)
        throw new RangeError('Enter a size from 10 to 2000 pixels.');
      const next = structuredClone(sheet);
      if (dialog === 'row-height') {
        next.rows = { ...next.rows };
        for (let y = selection.y1; y <= selection.y2; y++) next.rows[y] = { height: measure };
      } else {
        next.columns = [...(next.columns ?? [])];
        while (next.columns.length <= selection.x2) next.columns.push({});
        for (let x = selection.x1; x <= selection.x2; x++)
          next.columns[x] = { ...next.columns[x], width: measure };
      }
      updateSheet(next);
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
    let next = workbook;
    for (const match of targets) {
      const value = String(
        next.worksheets.find((sheet) => sheet.id === match.sheetId)!.data[match.y][match.x] ?? '',
      );
      const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), all ? 'gi' : 'i');
      next = writeCells(next, match.sheetId, match.x, match.y, [
        [value.replace(pattern, () => text)],
      ]);
    }
    mutate(next, all ? 'Replace all matches' : 'Replace match');
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
    draftDirty: draft?.key === draftKey,
    applyFormula,
    calculation,
    output,
    style,
    cellFormat,
    summary,
    grid,
    updateSheet,
    commit,
    mutate,
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
    savePhase,
    panel,
    panelOpen,
    setPanelOpen,
    openPanel,
    commentComposer,
    setCommentComposer,
    beginComment,
    postComment,
    decide,
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
    setNotice,
    dialog,
    dialogValue,
    setDialogValue,
    setDialog,
    openDialog,
    applyDialog,
    sheetCommand,
    insert,
    clear: (formatting = false) =>
      mutate(
        clearSelection(workbook, sheet.id, selection, formatting),
        formatting ? 'Clear formatting' : 'Clear cells',
      ),
    merge: (unmerge = false) =>
      mutate(
        mergeSelection(workbook, sheet.id, selection, unmerge),
        unmerge ? 'Unmerge cells' : 'Merge cells',
      ),
  };
}

export type SpreadsheetEditorController = ReturnType<typeof useSpreadsheetEditor>;
