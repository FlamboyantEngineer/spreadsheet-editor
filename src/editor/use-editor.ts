'use client';

import { type RefObject, useMemo, useRef, useState } from 'react';
import { type DisplayOptions, formatCellValue } from '../calculation/format.js';
import { summarizeCellValues } from '../calculation/summary.js';
import { type CellSelection, selectionCells } from '../model/editing.js';
import type { Workbook } from '../model/workbook.js';
import type { ClipboardCopy } from './clipboard-payload.js';
import type { GridController, SpreadsheetEditorProps, SpreadsheetMode } from './contracts.js';
import { editorAccess } from './editor-access.js';
import type { EditorCore, ViewedVersion } from './editor-core.js';
import { useCellDisplay } from './use-cell-display.js';
import { useComments } from './use-comments.js';
import {
  useEditorEditActions,
  useEditorFindActions,
  useEditorNavigationActions,
} from './use-editor-actions.js';
import { editorState, useEditorHandle } from './use-editor-handle.js';
import {
  useEditorActor,
  useEditorDialogState,
  useEditorDrafts,
  useEditorLive,
  useEditorPanels,
  useEditorPeople,
  useEditorPlace,
  useEditorView,
  useEditorWorkbook,
} from './use-editor-state.js';
import { useFind } from './use-find.js';
import { useNotice } from './use-notice.js';
import { useSessionActions } from './use-session-actions.js';
import { sessionSavePhase, type useSessionState } from './use-session-state.js';
import { useVersionActions } from './use-version-actions.js';

/** The count, sum and average of the selection's shown values. */
function useSelectionSummary(
  selection: CellSelection,
  output: ReturnType<typeof useCellDisplay>['output'],
  display: DisplayOptions | undefined,
) {
  return useMemo(() => {
    const values = selectionCells(selection).map(({ x, y }) => output?.[y]?.[x] ?? null);
    const { count, sum, average } = summarizeCellValues(values);
    return {
      count,
      sum: sum ? formatCellValue(sum, display) : '—',
      average: average ? formatCellValue(average, display) : '—',
    };
  }, [selection, output, display]);
}

/** The session's save state and which version commands the host offers. */
function sessionView(
  props: SpreadsheetEditorProps,
  { status, draft }: ReturnType<typeof useEditorWorkbook>['session'],
  {
    viewedVersion,
    hostCapabilities,
  }: { viewedVersion: ViewedVersion; hostCapabilities: EditorCore['hostCapabilities'] },
) {
  return {
    savePhase: sessionSavePhase(status),
    saveAutomatic: !!status?.automatic,
    sessionStatus: status,
    sessionDraft: draft,
    canSaveVersion: !viewedVersion && hostCapabilities.saveVersion && !!props.versions?.create,
    canRenameVersion: hostCapabilities.saveVersion && !!props.versions?.rename,
    canRestoreVersion: hostCapabilities.edit && !!props.versions?.restore,
  };
}

/** The editor's state, refs and commands for one render: everything but its React chrome. */
function useEditorCore(
  props: SpreadsheetEditorProps,
  grid: RefObject<GridController | null>,
  session: ReturnType<typeof useSessionState> | undefined,
) {
  const data = useEditorWorkbook(props, session);
  const { workbook, viewedVersion } = data;
  const [localMode, setLocalMode] = useState<SpreadsheetMode>(props.defaultMode ?? 'author');
  const access = editorAccess(props, {
    historical: !!viewedVersion,
    requestedMode: props.mode ?? localMode,
  });
  const actor = useEditorActor(props.identity);
  const place = useEditorPlace(props, workbook, { actorId: actor.id, grid });
  const drafts = useEditorDrafts(props, place.sheet.id);
  const copied = useRef<ClipboardCopy | null>(null);
  const panels = useEditorPanels();
  const notices = useNotice();
  const dialogs = useEditorDialogState();
  const { canEdit, canWrite, canSuggest } = access;
  const liveRefs = useEditorLive(
    workbook,
    { canEdit, canWrite, props, peers: place.peers },
    panels,
  );
  const [refreshVersion, setRefreshVersion] = useState(0);
  const display = useCellDisplay({
    props,
    workbook,
    sheet: place.sheet,
    selection: place.selection,
    draft: drafts.draft,
    nativeDraft: drafts.nativeDraft,
    canSuggest,
    actorId: actor.id,
    refreshVersion,
  });
  const core: EditorCore = {
    props,
    ...data,
    ...access,
    setLocalMode,
    ...place,
    actor,
    grid,
    ...drafts,
    display,
    ...liveRefs,
    copied,
    ...dialogs,
    setRefreshVersion,
    setNotice: notices.setNotice,
    warn: notices.warn,
    setPanel: panels.setPanel,
    setPanelOpen: panels.setPanelOpen,
  };
  return { core, data, place, drafts, panels, notices, liveRefs };
}

/** Comment threads, suggestion decisions, saved versions and the names of the people shown. */
function useEditorReview(
  core: EditorCore,
  {
    panels,
    edits,
    navigation,
  }: {
    panels: ReturnType<typeof useEditorPanels>;
    edits: ReturnType<typeof useEditorEditActions>;
    navigation: ReturnType<typeof useEditorNavigationActions>;
  },
) {
  const { props, workbook, actor } = core;
  const people = useEditorPeople(props, {
    historyAllowed: core.hostCapabilities.history,
    viewedVersion: core.viewedVersion,
    historyOpen: panels.panelOpen && panels.panel === 'history',
    comments: workbook.comments,
    actor,
  });
  const comments = useComments({
    workbook,
    sheet: core.sheet,
    selection: core.selection,
    initialCommentId: props.initialCommentId,
    idFactory: props.idFactory,
    canComment: core.canComment,
    canDecide: core.canDecide,
    grid: core.grid,
    select: navigation.select,
    openPanel: navigation.openPanel,
    commit: edits.commit,
    stamp: edits.stamp,
  });
  return { ...people, ...comments };
}

/** The find pane's state and its navigation and replacement commands. */
function useEditorFind(core: EditorCore) {
  const { display } = core;
  const find = useFind(
    core.workbook,
    {
      result: display.shown.result,
      readyCells: display.readyCells,
      presentation: display.presentation,
      display: core.props.display,
    },
    core.props.features?.search !== false,
  );
  return { ...find, ...useEditorFindActions(core, find) };
}

/** A version action must keep newer edits and drafts in the current editor. */
function captureVersionContext(core: EditorCore) {
  const read = () => {
    const { workbookId, versions, session } = core.live.current.props;
    return [
      workbookId,
      versions,
      session,
      session?.getSnapshot().localRevision ?? core.latestWorkbook.current,
      core.draftRef.current,
      core.nativeDraftRef.current,
    ];
  };
  const previous = read();
  return () => {
    if (!core.alive.current || read().some((value, index) => value !== previous[index]))
      throw new Error('The workbook changed while handling its saved version. Try again.');
  };
}

/** Saving through the session, and opening, saving and restoring versions. */
function useEditorSessionWork(
  core: EditorCore,
  {
    setViewedVersion,
    versionList,
    commitDraft,
    flush,
  }: {
    setViewedVersion: ReturnType<typeof useEditorWorkbook>['setViewedVersion'];
    versionList: ReturnType<typeof useEditorPeople>['versionList'];
    commitDraft: () => Workbook;
    flush: () => Promise<void>;
  },
) {
  const { alive, live, setNotice, warn } = core;
  const sessionActions = useSessionActions({ commitDraft, live, alive, setNotice, warn });
  const versions = useVersionActions({
    viewedVersion: core.viewedVersion,
    setViewedVersion,
    versionList,
    commitDraft,
    capture: () => captureVersionContext(core),
    flush,
    props: () => live.current.props,
    alive,
    setNotice,
  });
  return { ...sessionActions, ...versions };
}

export function useSpreadsheetEditor(
  props: SpreadsheetEditorProps,
  session?: ReturnType<typeof useSessionState>,
) {
  const grid = useRef<GridController | null>(null);
  const { core, data, place, drafts, panels, notices, liveRefs } = useEditorCore(
    props,
    grid,
    session,
  );
  const { sheet, selection, display } = core;
  const edits = useEditorEditActions(core);
  const navigation = useEditorNavigationActions(core);
  const review = useEditorReview(core, { panels, edits, navigation });
  const work = useEditorSessionWork(core, {
    setViewedVersion: data.setViewedVersion,
    versionList: review.versionList,
    commitDraft: navigation.commitDraft,
    flush: () => lifecycle.current.flush(),
  });
  const status = sessionView(props, data.session, core);
  const lifecycleState = editorState({
    status: data.session.status,
    display,
    nativeDraft: drafts.nativeDraft,
    flushing: work.flushing,
    savePhase: status.savePhase,
    notice: notices.notice,
    noticeTone: notices.noticeTone,
  });
  const { handle, lifecycle } = useEditorHandle(
    props.editorRef,
    {
      ...navigation,
      save: work.save,
      flush: work.flush,
      draftTarget: { sheetId: sheet.id, x: selection.x1, y: selection.y1 },
      state: lifecycleState,
    },
    { ...liveRefs, ...drafts, errorNotice: notices.errorNotice },
  );
  return {
    ...core,
    ...display,
    ...edits,
    ...navigation,
    ...review,
    ...useEditorFind(core),
    ...useEditorView(),
    ...notices,
    ...panels,
    ...work,
    ...status,
    lockedSelection: place.lockedSelection,
    initialViewport: core.sheetViews.current.get(sheet.id)?.viewport,
    attachGrid: (controller: GridController | null) => {
      grid.current = controller;
    },
    handle,
    lifecycleState,
    summary: useSelectionSummary(selection, display.output, props.display),
    history: data.currentHistory,
    canUndo: props.undo?.canUndo ?? !!data.currentHistory.past.length,
    canRedo: props.undo?.canRedo ?? !!data.currentHistory.future.length,
    viewedVersion: core.viewedVersion?.version ?? null,
  };
}

export type SpreadsheetEditorController = ReturnType<typeof useSpreadsheetEditor>;
