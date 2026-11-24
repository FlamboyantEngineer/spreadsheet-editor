'use client';

import { type RefObject, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { IdentityAdapter } from '../host/adapters.js';
import { equal } from '../model/cell-operations.js';
import { type CellSelection, INITIAL_SELECTION, selectionAddress } from '../model/editing.js';
import { placePeers, selectionLock } from '../model/presence.js';
import type { ReviewActor } from '../model/review.js';
import { copyWorkbook, createWorkbook, type Workbook } from '../model/workbook.js';
import type { GridController, SpreadsheetEditorProps } from './contracts.js';
import type {
  DraftText,
  EditorDialog,
  LiveEditor,
  NativeDraft,
  SheetView,
  ViewedVersion,
} from './editor-core.js';
import type { EditorHistory } from './history.js';
import { usePersonNames } from './use-person-names.js';
import { useSessionState } from './use-session-state.js';
import { useWorkbookPresence } from './use-workbook-presence.js';
import { useWorkbookVersions } from './use-workbook-versions.js';

/** The workbook the editor shows: a viewed version, the host's value, the session's or its own. */
export function useEditorWorkbook(
  props: SpreadsheetEditorProps,
  provided?: ReturnType<typeof useSessionState>,
) {
  const [local, setLocal] = useState(() => copyWorkbook(props.defaultValue ?? createWorkbook()));
  const subscribed = useSessionState(provided ? undefined : (props.session ?? undefined));
  const session = provided ?? subscribed;
  /** A saved version the reader opened, shown read-only in place of the workbook (SE-ADAPTERS-5.1). */
  const [viewedVersion, setViewedVersion] = useState<ViewedVersion>(null);
  const workbook = viewedVersion?.content ?? props.value ?? session.workbook ?? local;
  const [history, setHistory] = useState<EditorHistory>({ head: null, past: [], future: [] });
  const controlled = !!props.value;
  const currentHistory = useMemo(() => {
    if (history.head === workbook) return history;
    if (
      controlled &&
      (history.past.length || history.future.length) &&
      equal(history.head, workbook)
    )
      return { ...history, head: workbook };
    return { head: workbook, past: [], future: [] };
  }, [controlled, history, workbook]);
  return {
    setLocal,
    session,
    viewedVersion,
    setViewedVersion,
    workbook,
    currentHistory,
    setHistory,
  };
}

export function useEditorActor(identity: IdentityAdapter | undefined) {
  return useMemo(() => {
    const person = identity?.currentPerson();
    return person ? { id: person.id, name: person.name } : { id: 'local-user', name: 'You' };
  }, [identity]);
}

/** Forgets the views of removed worksheets, and leaves a removed active worksheet. */
function useSheetCleanup(
  workbook: Workbook,
  sheetId: string,
  {
    grid,
    sheetViews,
    setSheetId,
    setSelection,
  }: {
    grid: RefObject<GridController | null>;
    sheetViews: RefObject<Map<string, SheetView>>;
    setSheetId(id: string): void;
    setSelection(selection: CellSelection): void;
  },
) {
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
  }, [workbook, sheetId, grid, sheetViews, setSheetId, setSelection]);
}

/** The active worksheet and selection, the views of other worksheets, and the collaborators. */
export function useEditorPlace(
  props: SpreadsheetEditorProps,
  workbook: Workbook,
  { actorId, grid }: { actorId: string; grid: RefObject<GridController | null> },
) {
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
    () => placePeers(workbook, remotePeers, actorId),
    [workbook, remotePeers, actorId],
  );
  const lockedSelection = selectionLock(workbook, { sheetId: sheet.id, selection }, peers);
  const sheetViews = useRef(new Map<string, SheetView>());
  useSheetCleanup(workbook, sheetId, { grid, sheetViews, setSheetId, setSelection });
  return {
    setSheetId,
    sheet,
    selection,
    setSelection,
    setEditingCell,
    peers,
    lockedSelection,
    sheetViews,
  };
}

/** The formula bar's draft and the grid editor's draft, mirrored in refs for the handle. */
export function useEditorDrafts(props: SpreadsheetEditorProps, sheetId: string) {
  const [draft, setDraftState] = useState<DraftText>(null);
  const draftRef = useRef(draft);
  function setDraft(next: DraftText) {
    draftRef.current = next;
    setDraftState(next);
    const target = next && /^(.*):(\d+):(\d+)$/.exec(next.key);
    props.onDraftChange?.(
      next && target
        ? { sheetId: target[1], x: Number(target[2]), y: Number(target[3]), text: next.value }
        : null,
    );
  }
  const [nativeDraft, setNativeDraftState] = useState<NativeDraft>(null);
  const nativeDraftRef = useRef(nativeDraft);
  function setNativeDraft(next: NativeDraft) {
    nativeDraftRef.current = next;
    setNativeDraftState(next);
    props.onDraftChange?.(next ? { sheetId, ...next } : null);
  }
  return { draft, draftRef, setDraft, nativeDraft, nativeDraftRef, setNativeDraft };
}

/** The panel the editor shows beside the grid, and whether it is open. */
export function useEditorPanels() {
  const [panel, setPanel] = useState<string | null>(null),
    [panelOpen, setPanelOpen] = useState(false);
  return { panel, setPanel, panelOpen, setPanelOpen };
}

/**
 * State that outlives a render: whether the editor is mounted, its handle's listeners, the latest
 * workbook and the latest props, which callbacks read so a host's inline callback is no dependency.
 */
export function useEditorLive(
  workbook: Workbook,
  current: LiveEditor,
  { panel, panelOpen }: { panel: string | null; panelOpen: boolean },
) {
  const listeners = useRef(new Set<() => void>());
  const alive = useRef(false);
  useLayoutEffect(() => {
    alive.current = true;
    const subscribed = listeners.current;
    return () => {
      alive.current = false;
      subscribed.clear();
    };
  }, []);
  const latestWorkbook = useRef(workbook);
  const live = useRef(current);
  const { canEdit, canWrite, props, peers } = current;
  useLayoutEffect(() => {
    live.current = { canEdit, canWrite, props, peers };
  }, [canEdit, canWrite, props, peers]);
  // Read the callback through `live` so a host's inline callback is not a dependency (SE-HOST-2.1).
  useEffect(() => {
    live.current.props.onPanelChange?.({ id: panel, expanded: panelOpen });
  }, [panel, panelOpen]);
  useLayoutEffect(() => {
    latestWorkbook.current = workbook;
  }, [workbook]);
  return { listeners, alive, latestWorkbook, live };
}

/** The reader's view toggles: contributors, zoom, gridlines and formulas. */
export function useEditorView() {
  const [showContributors, setShowContributors] = useState(false);
  const [zoom, setZoom] = useState(100),
    [showGridlines, setShowGridlines] = useState(true),
    [showFormulas, setShowFormulas] = useState(false);
  return {
    showContributors,
    setShowContributors,
    zoom,
    setZoom,
    showGridlines,
    setShowGridlines,
    showFormulas,
    setShowFormulas,
  };
}

/** The open dialog, its value, and a structure change waiting for the reader's confirmation. */
export function useEditorDialogState() {
  const pendingStructure = useRef<(() => void) | null>(null);
  const [dialog, setDialog] = useState<EditorDialog>(null),
    [dialogValue, setDialogValue] = useState('');
  return { pendingStructure, dialog, setDialog, dialogValue, setDialogValue };
}

/** The workbook's saved versions, and the names of the people the editor shows. */
export function useEditorPeople(
  props: SpreadsheetEditorProps,
  {
    historyAllowed,
    viewedVersion,
    historyOpen,
    comments,
    actor,
  }: {
    historyAllowed: boolean;
    viewedVersion: ViewedVersion;
    historyOpen: boolean;
    comments: Workbook['comments'];
    actor: ReviewActor;
  },
) {
  const versionList = useWorkbookVersions(
    historyAllowed ? props.versions : undefined,
    props.workbookId,
    historyOpen,
  );
  const versionAuthors = useMemo(
    () => [viewedVersion?.version.author ?? null, ...versionList.versions.map((v) => v.author)],
    [viewedVersion, versionList.versions],
  );
  const personName = usePersonNames(props.identity, { comments, actor, others: versionAuthors });
  return { versionList, personName };
}
