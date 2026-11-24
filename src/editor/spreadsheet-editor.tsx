'use client';

import { ContextMenu } from '@flamboyantengineer/ui-components/context-menu';
import { FindBar } from '@flamboyantengineer/ui-components/find-bar';
import { InlineError } from '@flamboyantengineer/ui-components/inline-error';
import { type RefObject, useLayoutEffect, useRef, useState } from 'react';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import { parseFormulaReferenceHighlights } from '../model/formula-reference-parser.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import {
  CalculationStatus,
  EditorNotices,
  HistoricalBanner,
  InsertionActions,
} from './editor-banners.js';
import { CommentBubble, useCommentBubbleAnchor } from './editor-comment-bubble.js';
import { useEditorContextMenu } from './editor-context-menu.js';
import { EditorDialogModal } from './editor-dialog.js';
import { EditorFormulaBar, type EditorFormulaBarHandle } from './editor-formula-bar.js';
import { handleEditorKey } from './editor-keyboard.js';
import { EditorSidebar } from './editor-sidebar.js';
import { EditorToolbar } from './editor-toolbar.js';
import { EditorWorkspace } from './editor-workspace.js';
import { type SpreadsheetEditorController, useSpreadsheetEditor } from './use-editor.js';
import { useSessionState } from './use-session-state.js';

export type { SpreadsheetClipboardData } from './clipboard-payload.js';
export type {
  CalculationMode,
  SpreadsheetChrome,
  SpreadsheetClipboard,
  SpreadsheetEditorProps,
  SpreadsheetFeatures,
  SpreadsheetHostPanel,
  SpreadsheetMode,
} from './contracts.js';
export * from './public-host.js';
export { useWorkbookSession } from './use-workbook-session.js';
export function SpreadsheetEditor(props: SpreadsheetEditorProps) {
  const session = useSessionState(props.session ?? undefined);
  if (props.session !== undefined && !session.workbook) {
    const failed = session.status?.phase === 'unavailable';
    return (
      <section
        className="se-editor"
        data-session={failed ? 'unavailable' : 'loading'}
        aria-label={props['aria-label'] ?? 'Spreadsheet editor'}
      >
        {failed ? (
          <InlineError className="se-editor__notice">
            {session.status?.error?.message ?? 'The workbook could not be opened.'}
          </InlineError>
        ) : (
          <p className="se-editor__loading" role="status">
            Opening workbook…
          </p>
        )}
      </section>
    );
  }
  return (
    <SpreadsheetEditorInstance
      key={`${props.workbookId ?? ''}\n${props.sessionKey ?? 'spreadsheet'}`}
      props={props}
      session={session}
    />
  );
}
/** The formula bar, wired to the editor's draft, selection and reference picking. */
function FormulaBarSection({
  editor: e,
  formulaBar,
  keepsEditing,
  onEditingChange,
}: {
  editor: SpreadsheetEditorController;
  formulaBar: RefObject<EditorFormulaBarHandle | null>;
  keepsEditing(target: EventTarget | null): boolean;
  onEditingChange(editing: boolean): void;
}) {
  const { props, selection } = e;
  const activeFormat =
    e.previewSheet.formats?.[formatSpreadsheetAddress(selection.x1, selection.y1)];
  return (
    <EditorFormulaBar
      ref={formulaBar}
      formulaFunctions={props.formulaFunctions}
      includeBuiltinFormulaFunctions={props.includeBuiltinFormulaFunctions}
      selection={selection}
      draftValue={e.draftValue}
      result={formatWorksheetCell(
        e.output?.[selection.y1]?.[selection.x1] ?? null,
        activeFormat,
        props.display,
      )}
      isReadOnly={!e.canWrite || !!e.lockedSelection}
      pending={e.calculation.pending}
      error={!!e.calculation.error}
      draftDirty={e.draftDirty}
      onDraftChange={e.setDraftValue}
      onCommit={(move) =>
        e.attempt(() => {
          if (e.applyFormula() === false) return;
          if (move === 'down' && selection.y1 < 1048575)
            e.select({ ...selection, y1: selection.y1 + 1, y2: selection.y1 + 1 });
          e.grid.current?.focus?.();
        })
      }
      onCancel={() => {
        e.cancelFormula();
        e.grid.current?.focus?.();
      }}
      onNavigate={(next) => {
        let selected = false;
        e.attempt(() => {
          selected = e.select(next);
        });
        return selected;
      }}
      onEditingChange={onEditingChange}
      keepsEditing={keepsEditing}
      onLeaveEditor={() =>
        e.attempt(() => {
          e.applyFormula();
        })
      }
    />
  );
}

function FindSection({
  editor: e,
  input,
}: {
  editor: SpreadsheetEditorController;
  input: RefObject<HTMLInputElement | null>;
}) {
  const active = e.matches[e.matchIndex] ?? e.matches[0];
  useLayoutEffect(() => {
    input.current?.focus({ preventScroll: true });
  }, [input]);
  return (
    <div className="se-editor__find">
      <FindBar
        inputRef={input}
        label="Find in workbook"
        query={e.query}
        onQueryChange={(value) => e.setQuery(value)}
        matchCount={e.query ? e.matches.length : undefined}
        currentIndex={e.matchIndex}
        onNext={() => e.navigateMatch(1)}
        onPrevious={() => e.navigateMatch(-1)}
        replacement={e.replacement}
        onReplacementChange={e.setReplacement}
        onReplace={(text) => e.attempt(() => e.replace(false, text))}
        onReplaceAll={(text) => e.attempt(() => e.replace(true, text))}
        isReplaceDisabled={!e.canWrite || !active?.authored}
        isReplaceAllDisabled={!e.canWrite || !e.matches.some((match) => match.authored)}
        onClose={() => {
          e.setShowFind(false);
          e.grid.current?.focus?.();
        }}
      />
    </div>
  );
}

/** Formula reference picking: the formula bar and the grid editor share one draft. */
function useReferencePicking(
  e: SpreadsheetEditorController,
  formulaBar: RefObject<EditorFormulaBarHandle | null>,
) {
  const [gridEditing, setGridEditing] = useState(false);
  const [formulaActive, setFormulaActive] = useState(false);
  const picking = (formulaActive || gridEditing) && e.draftValue.startsWith('=');
  return {
    grid: {
      picking,
      references: picking ? parseFormulaReferenceHighlights(e.draftValue) : [],
      onFormulaReference: (range: string) => formulaBar.current?.insertReference(range),
      onEditStateChange: (editing: boolean) => {
        setGridEditing(editing);
        e.publishPresence(editing);
      },
    },
    onFormulaEditingChange: (editing: boolean) => {
      setFormulaActive(editing);
      e.publishPresence(editing);
    },
  };
}

function SpreadsheetEditorInstance({
  props,
  session,
}: {
  props: SpreadsheetEditorProps;
  session: ReturnType<typeof useSessionState>;
}) {
  const { chrome = 'full', panelRef, panelId, panelData } = props;
  const e = useSpreadsheetEditor(props, session);
  const root = useRef<HTMLElement | null>(null);
  const formulaBar = useRef<EditorFormulaBarHandle | null>(null);
  const findInput = useRef<HTMLInputElement | null>(null);
  const menu = useEditorContextMenu(e);
  const bubble = useCommentBubbleAnchor(e);
  const picking = useReferencePicking(e, formulaBar);
  function openFind() {
    e.setShowFind(true);
    findInput.current?.focus({ preventScroll: true });
  }
  function keepsEditing(target: EventTarget | null) {
    if (!(target instanceof Element)) return false;
    return !!(
      root.current?.contains(target) ||
      target.closest('[role="menu"], [role="listbox"], [role="dialog"], .se-comment-bubble')
    );
  }
  return (
    <section
      ref={root}
      className={['se-editor', props.className].filter(Boolean).join(' ')}
      aria-label={props['aria-label'] ?? 'Spreadsheet editor'}
      data-chrome={chrome}
      data-mode={e.mode}
      data-calculation-mode={props.calculation?.mode ?? 'client'}
      data-pending={e.calculation.pending}
      onKeyDownCapture={(event) =>
        handleEditorKey(e, event, { openCellMenu: menu.openCellMenu, openFind })
      }
    >
      <EditorToolbar editor={e} />
      {e.viewedVersion && <HistoricalBanner editor={e} version={e.viewedVersion} />}
      {chrome !== 'minimal' && (
        <FormulaBarSection
          editor={e}
          formulaBar={formulaBar}
          keepsEditing={keepsEditing}
          onEditingChange={picking.onFormulaEditingChange}
        />
      )}
      {e.showFind && props.features?.search !== false && (
        <FindSection editor={e} input={findInput} />
      )}
      <div className="se-editor__workspace">
        <EditorWorkspace
          editor={e}
          picking={picking.grid}
          onContextMenu={menu.openContext}
          onViewportChange={bubble.positionComment}
        />
        {props.features?.panels !== false && (chrome !== 'minimal' || e.panelOpen) && (
          <EditorSidebar editor={e} panelRef={panelRef} panelId={panelId} panelData={panelData} />
        )}
      </div>
      {props.onInsert && <InsertionActions editor={e} />}
      <CalculationStatus editor={e} />
      <EditorNotices editor={e} />
      <CommentBubble editor={e} anchor={bubble.anchor} />
      <ContextMenu
        controller={menu.context}
        triggerRef={root}
        aria-label="Spreadsheet context menu"
      />
      <EditorDialogModal editor={e} />
    </section>
  );
}
