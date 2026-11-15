'use client';

import {
  lazy,
  Suspense,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type KeyboardEvent,
} from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { ContextMenu, useMenuController } from '@flamboyantengineer/ui-components/context-menu';
import { DockPanel, DockPanelPane } from '@flamboyantengineer/ui-components/dock-panel';
import { FindBar } from '@flamboyantengineer/ui-components/find-bar';
import { Banner } from '@flamboyantengineer/ui-components/banner';
import { InlineError } from '@flamboyantengineer/ui-components/inline-error';
import { ZoomControl } from '@flamboyantengineer/ui-components/zoom-control';
import { Popover } from '@flamboyantengineer/ui-components/popover';
import { findOpenComment } from '../model/review.js';
import { Modal } from '@flamboyantengineer/ui-components/modal';
import {
  InfoGlyph,
  ClockGlyph,
  ReplyGlyph,
  EditGlyph,
} from '@flamboyantengineer/ui-components/core';
import { VisuallyHidden } from '@flamboyantengineer/ui-components/visually-hidden';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { worksheetExtent } from '../model/extent.js';
import { parseFormulaReferenceHighlights } from '../model/formula-reference-parser.js';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
} from '../engine/jspreadsheet-host.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';
import { Tabs, TabPanel } from '@flamboyantengineer/ui-components/tabs';
import { EditorFormulaBar, type EditorFormulaBarHandle } from './editor-formula-bar.js';
import { WorksheetStrip } from './worksheet-strip.js';
import { EditorToolbar } from './editor-toolbar.js';
import { editorMenus } from './editor-menus.js';
import { CommentsPanel, SuggestionsPanel, HistoryPanel, PropertiesPanel } from './review-panels.js';

export type {
  SpreadsheetEditorProps,
  SpreadsheetChrome,
  SpreadsheetFeatures,
  SpreadsheetSaveContext,
  SpreadsheetHistory,
  SpreadsheetHistoryEntry,
  SpreadsheetHistoricalView,
  CalculationMode,
  SpreadsheetMode,
  SpreadsheetClipboard,
  SpreadsheetHostPanel,
} from './contracts.js';
export * from './public-host.js';
import { READ_ONLY_CAPABILITIES } from '../host/capabilities.js';
const WorkbookGrid = lazy(() => import('../engine/workbook-grid.js'));

export function SpreadsheetEditor(props: SpreadsheetEditorProps) {
  return (
    <SpreadsheetEditorInstance
      key={`${props.workbookId ?? ''}\n${props.sessionKey ?? 'spreadsheet'}`}
      {...props}
      capabilities={props.historical ? READ_ONLY_CAPABILITIES : props.capabilities}
    />
  );
}
function SpreadsheetEditorInstance(props: SpreadsheetEditorProps) {
  const e = useSpreadsheetEditor(props);
  const chrome = props.chrome ?? 'full';
  const context = useMenuController(),
    root = useRef<HTMLElement | null>(null);
  const hasMountedGrid = useRef(false);
  const [contextKind, setContextKind] = useState<'cell' | 'row' | 'column'>('cell');
  const contextItems = () => {
    const menus = editorMenus(e);
    return contextKind === 'row'
      ? menus.rowContext
      : contextKind === 'column'
        ? menus.columnContext
        : menus.context;
  };
  const [contextPoint, setContextPoint] = useState<{ x: number; y: number } | null>(null);
  const formulaBar = useRef<EditorFormulaBarHandle | null>(null);
  const [gridEditing, setGridEditing] = useState(false);
  const [formulaActive, setFormulaActive] = useState(false);
  const bubbleTrigger = useRef<HTMLSpanElement | null>(null);
  const [bubbleAnchor, setBubbleAnchor] = useState<{ left: number; top: number } | null>(null);
  function positionComment() {
    if (!e.activeCommentId) return;
    const rect = e.grid.current?.getAnchor(e.selection);
    if (rect)
      setBubbleAnchor((previous) =>
        previous?.left === rect.right && previous.top === rect.top
          ? previous
          : { left: rect.right, top: rect.top },
      );
  }
  useLayoutEffect(() => {
    positionComment();
  }, [e.activeCommentId, e.selection, e.workbook]);
  const references =
    (formulaActive || gridEditing) && e.draftValue.startsWith('=')
      ? parseFormulaReferenceHighlights(e.draftValue)
      : [];
  function insertReference(range: string) {
    formulaBar.current?.insertReference(range);
  }
  function keepsEditing(target: EventTarget | null) {
    if (!(target instanceof Element)) return false;
    return !!(
      root.current?.contains(target) ||
      target.closest('[role="menu"], [role="listbox"], [role="dialog"], .se-comment-bubble')
    );
  }
  function openContext(coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect) {
    setContextKind(coords.x < 0 ? 'row' : coords.y < 0 ? 'column' : 'cell');
    if (coords.x < 0 && coords.y >= 0) {
      if (coords.y < e.selection.y1 || coords.y > e.selection.y2)
        e.select({
          x1: 0,
          x2: Math.max(
            (e.sheet.minDimensions?.[0] ?? 26) - 1,
            ...e.sheet.data.map((row) => row.length - 1),
          ),
          y1: coords.y,
          y2: coords.y,
        });
    } else if (coords.y < 0 && coords.x >= 0) {
      if (coords.x < e.selection.x1 || coords.x > e.selection.x2)
        e.select({
          x1: coords.x,
          x2: coords.x,
          y1: 0,
          y2: Math.max((e.sheet.minDimensions?.[1] ?? 100) - 1, e.sheet.data.length - 1),
        });
    } else if (
      coords.x >= 0 &&
      coords.y >= 0 &&
      (coords.x < e.selection.x1 ||
        coords.x > e.selection.x2 ||
        coords.y < e.selection.y1 ||
        coords.y > e.selection.y2)
    )
      e.select({ x1: coords.x, x2: coords.x, y1: coords.y, y2: coords.y });
    e.grid.current?.releaseKeyboard();
    setContextPoint({ x: rect.left + 10, y: rect.bottom });
  }
  useEffect(() => {
    if (contextPoint) {
      context.open(contextPoint, { items: contextItems() });
      setContextPoint(null);
    }
  }, [contextPoint, context, contextKind]);
  useEffect(() => {
    context.setItems(contextItems());
  }, [e.workbook, e.selection, e.mode, e.capabilities, context, contextKind]);
  function onKeyDown(event: KeyboardEvent<HTMLElement>) {
    if (event.nativeEvent.isComposing || event.repeat) return;
    if (event.altKey && !((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'm'))
      return;
    const modifier = event.ctrlKey || event.metaKey,
      key = event.key.toLowerCase();
    let action: (() => void | Promise<void>) | undefined;
    if (modifier && key === 'f' && props.features?.search !== false)
      action = () => e.setShowFind(true);
    if (modifier && key === 's' && props.onSave) action = () => e.save();
    if (modifier && event.altKey && key === 'm') action = e.beginComment;
    const textTarget =
      event.target instanceof HTMLInputElement ||
      event.target instanceof HTMLTextAreaElement ||
      (event.target instanceof HTMLElement && event.target.isContentEditable);
    const gridTarget =
      event.target instanceof Element && !!event.target.closest('.se-editor__grid');
    if (!textTarget && gridTarget) {
      if (
        e.lockedSelection &&
        ((key.length === 1 && !modifier) ||
          key === 'enter' ||
          key === 'delete' ||
          key === 'backspace')
      ) {
        event.preventDefault();
        event.stopPropagation();
        e.warn(`${e.lockedSelection.name} is editing a cell in this selection.`);
        return;
      }
      if (modifier && key === 'c') action = () => e.copy();
      if (modifier && key === 'x') action = () => e.copy(true);
      if (modifier && key === 'v') action = e.paste;
      if (modifier && key === 'z') action = event.shiftKey ? e.redo : e.undo;
      if (modifier && key === 'y') action = e.redo;
      if (modifier && key === 'b' && e.canWrite)
        action = () => e.toggleFormat('font-weight', 'bold');
      if (modifier && key === 'i' && e.canWrite)
        action = () => e.toggleFormat('font-style', 'italic');
      if (key === 'contextmenu' || (key === 'f10' && event.shiftKey))
        action = () => {
          setContextKind('cell');
          const rect = e.grid.current?.getAnchor(e.selection);
          if (rect) {
            e.grid.current?.releaseKeyboard();
            setContextPoint({ x: rect.left, y: rect.bottom });
          }
        };
    }
    if (action) {
      event.preventDefault();
      event.stopPropagation();
      e.attempt(action);
    }
  }
  const titleForDialog =
    e.dialog === 'confirm-structure'
      ? 'Change merged cells'
      : e.dialog === 'row-height'
        ? 'Row height'
        : e.dialog === 'column-width'
          ? 'Column width'
          : e.dialog === 'import'
            ? 'Import workbook'
            : 'Export workbook';
  const activeFormat =
    e.previewSheet.formats?.[formatSpreadsheetAddress(e.selection.x1, e.selection.y1)];
  return (
    <section
      ref={root}
      className={['se-editor', props.className].filter(Boolean).join(' ')}
      aria-label={props['aria-label'] ?? 'Spreadsheet editor'}
      data-chrome={chrome}
      data-mode={e.mode}
      data-calculation-mode={props.calculation?.mode ?? 'client'}
      data-pending={e.calculation.pending}
      onKeyDownCapture={onKeyDown}
      onFocusCapture={(event) => {
        if (event.target instanceof Element && !event.target.closest('.se-editor__grid'))
          e.grid.current?.releaseKeyboard();
      }}
    >
      <EditorToolbar editor={e} />
      {props.historical && (
        <Banner
          className="se-editor__historical"
          tone="info"
          description={
            <span>
              You are viewing a historical version from{' '}
              {new Intl.DateTimeFormat(undefined, {
                dateStyle: 'medium',
                timeStyle: 'long',
              }).format(new Date(props.historical.timestamp))}
              .
            </span>
          }
          actions={
            <Button size="sm" onPress={() => e.attempt(() => props.historical!.onReturnCurrent())}>
              View Current Version
            </Button>
          }
        />
      )}
      {chrome !== 'minimal' && (
        <EditorFormulaBar
          ref={formulaBar}
          formulaFunctions={props.formulaFunctions}
          includeBuiltinFormulaFunctions={props.includeBuiltinFormulaFunctions}
          selection={e.selection}
          draftValue={e.draftValue}
          result={formatWorksheetCell(
            e.output?.[e.selection.y1]?.[e.selection.x1] ?? null,
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
              if (move === 'down' && e.selection.y1 < 1048575)
                e.select({ ...e.selection, y1: e.selection.y1 + 1, y2: e.selection.y1 + 1 });
              e.grid.current?.focus?.();
            })
          }
          onCancel={() => {
            e.cancelFormula();
            e.grid.current?.focus?.();
          }}
          onNavigate={(selection) => e.attempt(() => e.select(selection))}
          onEditingChange={(editing) => {
            setFormulaActive(editing);
            e.publishPresence(editing);
          }}
          keepsEditing={keepsEditing}
          onLeaveEditor={() =>
            e.attempt(() => {
              e.applyFormula();
            })
          }
        />
      )}
      {e.showFind && (
        <div className="se-editor__find" onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
          <FindBar
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
            isReplaceDisabled={!e.canWrite || !(e.matches[e.matchIndex] ?? e.matches[0])?.authored}
            isReplaceAllDisabled={!e.canWrite || !e.matches.some((match) => match.authored)}
            onClose={() => e.setShowFind(false)}
          />
        </div>
      )}
      <div className="se-editor__workspace">
        <Tabs
          className="se-editor__main se-editor__worksheet-tabs"
          selectedKey={e.sheet.id}
          onSelectionChange={(id) => e.attempt(() => e.activateSheet(String(id)))}
          keyboardActivation="manual"
        >
          <TabPanel id={e.sheet.id} className="se-editor__worksheet-panel">
            <div
              className="se-editor__grid"
              role="group"
              aria-label="Workbook grid"
              tabIndex={0}
              data-gridlines={e.showGridlines}
            >
              <Suspense fallback={<p role="status">Loading spreadsheet…</p>}>
                <WorkbookGrid
                  key={e.sheet.id}
                  subscribeWindowBlur={props.events?.subscribeWindowBlur}
                  worksheet={e.previewSheet}
                  logicalExtent={worksheetExtent(e.previewSheet, e.workbook)}
                  zoom={e.zoom}
                  values={e.output}
                  draft={e.draftCell ?? undefined}
                  readyCells={e.readyCells}
                  publication={e.publication}
                  cellPresentation={e.presentation}
                  renderCellContent={props.renderCellContent}
                  formulaFunctions={props.formulaFunctions}
                  idFactory={props.idFactory}
                  includeBuiltinFormulaFunctions={props.includeBuiltinFormulaFunctions}
                  initialViewport={e.initialViewport}
                  focusOnMount={hasMountedGrid.current}
                  onNativeDraft={e.setNativeDraft}
                  onError={(error) =>
                    e.attempt(() => {
                      throw error;
                    })
                  }
                  formulaReferences={references}
                  pickingFormula={(formulaActive || gridEditing) && e.draftValue.startsWith('=')}
                  onFormulaReference={insertReference}
                  worksheetIndex={e.workbook.worksheets.findIndex(
                    (sheet) => sheet.id === e.sheet.id,
                  )}
                  collaborators={e.peers}
                  showContributors={e.showContributors}
                  onViewportChange={positionComment}
                  onEditStateChange={(editing) => {
                    setGridEditing(editing);
                    e.publishPresence(editing);
                  }}
                  onHeaderDelete={(axis) => e.attempt(() => e.structure(axis, 'delete'))}
                  onEditDraft={(value) =>
                    value === null ? e.cancelFormula() : e.setDraftValue(value)
                  }
                  display={props.display}
                  readOnly={!e.canWrite}
                  reviewMode={e.canSuggest}
                  onReviewChange={e.proposeCells}
                  onPasteCells={e.pasteCells}
                  onBlockedAction={e.warn}
                  showFormulas={e.showFormulas}
                  comments={e.workbook.comments}
                  suggestions={e.workbook.suggestions}
                  selection={e.selection}
                  onController={(controller) => {
                    e.grid.current = controller;
                    if (controller) hasMountedGrid.current = true;
                    if (e.dialog) controller?.releaseKeyboard();
                  }}
                  unavailableText={e.calculation.error ? 'Unavailable' : '…'}
                  onChange={(next) => e.updateSheet(next)}
                  onSelectionChange={e.onGridSelection}
                  onContextMenu={openContext}
                  onCommentRequest={(coords) => {
                    const thread = findOpenComment(
                      e.workbook,
                      e.sheet.id,
                      formatSpreadsheetAddress(coords.x, coords.y),
                    );
                    if (thread) e.openComment(thread.id);
                    else {
                      const suggestion = e.workbook.suggestions?.find(
                        (item) =>
                          item.status === 'pending' &&
                          !item.outdated &&
                          item.changes.some(
                            (change) =>
                              change.sheetId === e.sheet.id &&
                              change.address === formatSpreadsheetAddress(coords.x, coords.y),
                          ),
                      );
                      if (suggestion) e.openSuggestion(suggestion.id);
                    }
                  }}
                />
              </Suspense>
            </div>
          </TabPanel>
          <WorksheetStrip
            sheets={e.workbook.worksheets}
            selectedId={e.sheet.id}
            canEdit={e.canEdit}
            onRename={(id, name) => e.attempt(() => e.renameSheet(id, name))}
            validateName={e.validateSheetName}
            onCommand={(command, id) => e.attempt(() => e.sheetCommand(command, id))}
          >
            <ZoomControl
              value={e.zoom}
              onChange={e.setZoom}
              levels={[50, 75, 100, 125, 150]}
              aria-label="Spreadsheet zoom"
            />
          </WorksheetStrip>
        </Tabs>
        {props.features?.panels !== false && (chrome !== 'minimal' || e.panelOpen) && (
          <div
            className="se-editor__sidebar"
            {...props.panelData}
            id={props.panelId}
            ref={props.panelRef}
            onFocusCapture={() => e.grid.current?.releaseKeyboard()}
          >
            <DockPanel
              label="Spreadsheet panels"
              selectedKey={e.panel ?? undefined}
              onSelectionChange={(key) => e.openPanel(String(key))}
              isExpanded={e.panelOpen}
              onExpandedChange={e.setPanelOpen}
              defaultSize={340}
              minSize={260}
              maxSize={600}
            >
              <DockPanelPane id="properties" label="Properties" icon={InfoGlyph}>
                {props.panels?.find((panel) => panel.id === 'properties')?.content ?? (
                  <PropertiesPanel editor={e} />
                )}
              </DockPanelPane>
              <DockPanelPane id="history" label="History" icon={ClockGlyph}>
                {props.panels?.find((panel) => panel.id === 'history')?.content ?? (
                  <HistoryPanel editor={e} />
                )}
              </DockPanelPane>
              <DockPanelPane
                id="comments"
                label="Comments"
                icon={ReplyGlyph}
                count={
                  (e.workbook.comments?.filter(
                    (thread) => thread.state === 'open' && !thread.suggestion,
                  ).length ?? 0) +
                  (e.workbook.suggestions?.filter((item) => item.status === 'pending').length ?? 0)
                }
              >
                {props.panels?.find((panel) => panel.id === 'comments')?.content ?? (
                  <CommentsPanel editor={e} />
                )}
              </DockPanelPane>
              <DockPanelPane
                id="suggestions"
                label="Suggestions"
                icon={EditGlyph}
                count={e.workbook.suggestions?.filter((item) => item.status === 'pending').length}
              >
                <SuggestionsPanel editor={e} />
              </DockPanelPane>
              {props.panels
                ?.filter(
                  (panel) =>
                    !['properties', 'history', 'comments', 'suggestions'].includes(panel.id),
                )
                .map((panel) => (
                  <DockPanelPane key={panel.id} id={panel.id} label={panel.label} icon={InfoGlyph}>
                    {panel.content}
                  </DockPanelPane>
                ))}
            </DockPanel>
          </div>
        )}
      </div>
      {props.onInsert && (
        <div
          className="se-editor__insertion"
          role="group"
          aria-label="Insert selected cells"
          onFocusCapture={() => e.grid.current?.releaseKeyboard()}
        >
          <Button
            onPress={() => e.attempt(() => e.insert('table'))}
            isDisabled={e.draftDirty || e.calculation.pending || !!e.calculation.error}
          >
            Insert as document table
          </Button>
          <Button onPress={() => e.attempt(() => e.insert('spreadsheet'))}>
            Insert as live spreadsheet
          </Button>
        </div>
      )}
      <VisuallyHidden>
        <p className="se-editor__status" aria-label="Calculation status" role="status">
          {e.calculation.error ??
            (e.calculation.pending
              ? 'Calculating…'
              : props.calculation?.mode === 'external'
                ? 'Published by host'
                : `Calculated ${props.calculation?.mode === 'server' ? 'on server' : 'on client'}`)}
        </p>
        <span aria-label="Selection summary">
          Count: {e.summary.count} · Sum: {e.summary.sum} · Average: {e.summary.average}
        </span>
      </VisuallyHidden>
      {e.calculation.error && (
        <InlineError className="se-editor__notice">{e.calculation.error}</InlineError>
      )}
      {e.notice && (
        <Banner
          className="se-editor__notice"
          tone={e.noticeTone}
          key={e.noticeRevision}
          description={e.notice}
          onDismiss={() => e.setNotice(null)}
          dismissLabel="Dismiss"
        />
      )}
      <span
        ref={bubbleTrigger}
        aria-hidden="true"
        style={{
          position: 'fixed',
          left: bubbleAnchor?.left ?? 0,
          top: bubbleAnchor?.top ?? 0,
          width: 0,
          height: 0,
          pointerEvents: 'none',
        }}
      />
      <Popover
        className="se-comment-bubble"
        triggerRef={bubbleTrigger}
        placement="right top"
        offset={8}
        isNonModal
        isOpen={!!e.activeCommentId && !!bubbleAnchor}
        onOpenChange={(open) => {
          if (!open) {
            e.setActiveCommentId(null);
            e.setCommentComposer(false);
          }
        }}
        title={e.activeCommentId === 'draft' ? 'New comment' : 'Comment'}
        showCloseButton
      >
        <div onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
          <CommentsPanel editor={e} threadId={e.activeCommentId ?? undefined} />
        </div>
      </Popover>
      <ContextMenu controller={context} triggerRef={root} aria-label="Spreadsheet context menu" />
      <Modal
        isOpen={e.dialog !== null}
        onOpenChange={(open) => {
          if (!open) e.setDialog(null);
        }}
        title={titleForDialog}
        primaryAction={{
          label:
            e.dialog === 'confirm-structure'
              ? 'Continue'
              : e.dialog === 'export'
                ? 'Done'
                : e.dialog === 'import'
                  ? 'Import'
                  : 'Apply',
          onPress: () => e.applyDialog(),
          isDisabled: !!e.dialogError,
        }}
        cancelAction={e.dialog === 'export' ? null : {}}
        description={
          e.dialog === 'import'
            ? 'Paste CSV data or a serialized workbook. All digits in CSV numeric fields are preserved.'
            : undefined
        }
      >
        <div onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
          {e.dialog === 'confirm-structure' ? (
            <p>
              This change intersects merged cells. Continue with the row or column change and
              adjusted merge spans?
            </p>
          ) : (
            <TextField
              autoFocus
              label={
                e.dialog === 'import'
                  ? 'CSV or workbook JSON'
                  : e.dialog === 'export'
                    ? 'Workbook JSON'
                    : 'Size in pixels'
              }
              tracksChanges={false}
              value={e.dialogValue}
              onChange={e.setDialogValue}
              multiline={e.dialog === 'import' || e.dialog === 'export'}
              isReadOnly={e.dialog === 'export'}
              minRows={8}
              maxRows={18}
            />
          )}
          {e.dialogError && <p role="alert">{e.dialogError}</p>}
        </div>
      </Modal>
    </section>
  );
}
