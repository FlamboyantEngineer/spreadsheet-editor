'use client';

import {
  type Ref,
  type RefObject,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
} from 'react';
import type { WorkbookSessionStatus } from '../host/workbook-session.js';
import type { Workbook } from '../model/workbook.js';
import type { DraftText, LiveEditor, NativeDraft } from './editor-core.js';
import { createEditorHandle, type EditorLifecycle } from './editor-handle.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorState } from './public-host.js';
import type { useCellDisplay } from './use-cell-display.js';
import { sessionErrorMessage } from './use-session-state.js';

/** What the editor reports to its host through `getState` (SE-HOST). */
export function editorState({
  status,
  display,
  nativeDraft,
  flushing,
  savePhase,
  notice,
  noticeTone,
}: {
  status: WorkbookSessionStatus | null;
  display: ReturnType<typeof useCellDisplay>;
  nativeDraft: NativeDraft;
  flushing: boolean;
  savePhase: string;
  notice: string | null;
  noticeTone: 'danger' | 'warning';
}): SpreadsheetEditorState {
  const { draftDirty, visibleCalculation } = display;
  return {
    dirty: !!status?.dirty || draftDirty || !!nativeDraft,
    draftDirty: draftDirty || !!nativeDraft,
    saving: savePhase === 'saving',
    flushing,
    error: sessionErrorMessage(status) ?? (noticeTone === 'danger' ? notice : null),
    calculation: {
      inputKey: visibleCalculation.inputKey,
      revision: visibleCalculation.revision,
      status: visibleCalculation.status,
    },
  };
}

/** The refs the host's handle reads between renders. */
interface EditorHandleRefs {
  alive: RefObject<boolean>;
  latestWorkbook: RefObject<Workbook>;
  draftRef: RefObject<DraftText>;
  nativeDraftRef: RefObject<NativeDraft>;
  live: RefObject<LiveEditor>;
  errorNotice: RefObject<string | null>;
  listeners: RefObject<Set<() => void>>;
}

/**
 * The handle hosts reach through `editorRef`. It reads this render's lifecycle through a ref, and
 * its listeners hear each change of the editor's state.
 */
export function useEditorHandle(
  editorRef: Ref<SpreadsheetEditorHandle> | undefined,
  current: EditorLifecycle,
  refs: EditorHandleRefs,
) {
  const lifecycle = useRef<EditorLifecycle>(current);
  // The handle reads the latest render's lifecycle, including a call made before this render
  // commits, so the ref follows each render rather than each commit.
  // eslint-disable-next-line react-hooks/refs -- the handle's view of the latest render
  lifecycle.current = current;
  const { alive, latestWorkbook, draftRef, nativeDraftRef, live, errorNotice, listeners } = refs;
  const handle = useMemo<SpreadsheetEditorHandle>(
    () =>
      // eslint-disable-next-line react-hooks/refs -- the handle stores the refs and reads them only when the host calls it
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
    [alive, latestWorkbook, draftRef, nativeDraftRef, live, errorNotice, listeners],
  );
  useImperativeHandle(editorRef, () => handle, [handle]);
  const stateIdentity = JSON.stringify(current.state);
  useLayoutEffect(() => {
    for (const listener of listeners.current) listener();
  }, [stateIdentity, listeners]);
  return { handle, lifecycle };
}
