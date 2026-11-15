import type { RefObject } from 'react';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import { copyWorkbook, type Workbook } from '../model/workbook.js';
import type { CellSelection } from '../model/editing.js';
import type { WorkbookSession } from '../host/workbook-session.js';
import type { SpreadsheetEditorHandle, SpreadsheetEditorState } from './public-host.js';

/** What the editor hands its imperative handle each render. */
export interface EditorLifecycle {
  setHostPanel(id: string | null): void;
  restoreDraft(value: { sheetId: string; x: number; y: number; text: string }): void;
  commitDraft(): Workbook;
  cancelDraft(): void;
  save(): Promise<void>;
  flush(): Promise<void>;
  select(range: CellSelection, targetSheetId?: string): boolean;
  draftTarget: { sheetId: string; x: number; y: number };
  state: SpreadsheetEditorState;
}

/** The handle hosts reach through `editorRef`; it reads the editor only through these refs. */
export function createEditorHandle(refs: {
  alive: RefObject<boolean>;
  lifecycle: RefObject<EditorLifecycle>;
  latestWorkbook: RefObject<Workbook>;
  draftRef: RefObject<{ key: string; value: string } | null>;
  nativeDraftRef: RefObject<{ x: number; y: number; text: string } | null>;
  session: () => WorkbookSession | null | undefined;
  errorNotice: RefObject<string | null>;
  listeners: RefObject<Set<() => void>>;
}): SpreadsheetEditorHandle {
  const {
    alive,
    lifecycle,
    latestWorkbook,
    draftRef,
    nativeDraftRef,
    session,
    errorNotice,
    listeners,
  } = refs;
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
              latestWorkbook.current.worksheets.find((value) => value.id === target.sheetId)?.data[
                target.y
              ]?.[target.x] ?? '',
            ));
      return structuredClone({
        ...current,
        dirty: !!session()?.status().dirty || draftDirty,
        draftDirty,
        saving: session()?.status().phase === 'saving',
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
}
