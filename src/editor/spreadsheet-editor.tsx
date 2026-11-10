'use client';

import {
  lazy,
  Suspense,
  useEffect,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
} from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { FormulaBar } from '@flamboyantengineer/ui-components/formula-bar';
import { Menu } from '@flamboyantengineer/ui-components/menu';
import {
  ContextMenu,
  ContextMenuTrigger,
  useMenuController,
} from '@flamboyantengineer/ui-components/context-menu';
import { DockPanel, DockPanelPane } from '@flamboyantengineer/ui-components/dock-panel';
import { FindBar } from '@flamboyantengineer/ui-components/find-bar';
import { ZoomControl } from '@flamboyantengineer/ui-components/zoom-control';
import { Modal } from '@flamboyantengineer/ui-components/modal';
import {
  InfoGlyph,
  ClockGlyph,
  ReplyGlyph,
  EditGlyph,
} from '@flamboyantengineer/ui-components/core';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import {
  parseSpreadsheetRange,
  formatSpreadsheetAddress,
} from '../model/spreadsheet-addressing.js';
import { selectionAddress } from '../model/editing.js';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
} from '../engine/jspreadsheet-host.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { useSpreadsheetEditor } from './use-editor.js';
import { EditorToolbar } from './editor-toolbar.js';
import { editorMenus } from './editor-menus.js';
import { CommentsPanel, SuggestionsPanel, HistoryPanel, PropertiesPanel } from './review-panels.js';

export type {
  SpreadsheetEditorProps,
  CalculationMode,
  SpreadsheetMode,
  SpreadsheetPermissions,
  SpreadsheetClipboard,
  SpreadsheetHostPanel,
} from './contracts.js';
const WorkbookGrid = lazy(() => import('../engine/workbook-grid.js'));
const COMPLETIONS = SUPPORTED_FUNCTIONS.map((name) => ({ label: name, insertText: `${name}(` }));

export function SpreadsheetEditor(props: SpreadsheetEditorProps) {
  const e = useSpreadsheetEditor(props);
  const context = useMenuController(),
    root = useRef<HTMLElement | null>(null);
  const [contextPoint, setContextPoint] = useState<{ x: number; y: number } | null>(null);
  const [addressDraft, setAddressDraft] = useState<string | null>(null);
  const address = selectionAddress(e.selection);
  function openContext(coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect) {
    if (
      coords.x < e.selection.x1 ||
      coords.x > e.selection.x2 ||
      coords.y < e.selection.y1 ||
      coords.y > e.selection.y2
    )
      e.select({ x1: coords.x, x2: coords.x, y1: coords.y, y2: coords.y });
    e.grid.current?.releaseKeyboard();
    setContextPoint({ x: rect.left + 10, y: rect.bottom });
  }
  useEffect(() => {
    if (contextPoint) {
      context.open(contextPoint, { items: editorMenus(e).context });
      setContextPoint(null);
    }
  }, [contextPoint, context]);
  useEffect(() => {
    context.setItems(editorMenus(e).context);
  }, [e.workbook, e.selection, e.mode, props.readOnly, props.permissions, context]);
  function jumpToAddress() {
    e.attempt(() => {
      const range = parseSpreadsheetRange(addressDraft ?? address);
      if (!range) throw new RangeError('Enter a valid cell or range, such as B2 or A1:C3.');
      e.select({ x1: range.start.x, x2: range.end.x, y1: range.start.y, y2: range.end.y });
      setAddressDraft(null);
    });
  }
  function onKeyDown(event: KeyboardEvent<HTMLElement>) {
    const modifier = event.ctrlKey || event.metaKey,
      key = event.key.toLowerCase();
    let action: (() => void | Promise<void>) | undefined;
    if (modifier && key === 'f') action = () => e.setShowFind(true);
    if (modifier && key === 's' && props.onSave) action = e.save;
    if (modifier && event.altKey && key === 'm') action = e.beginComment;
    const textTarget =
      event.target instanceof HTMLInputElement ||
      event.target instanceof HTMLTextAreaElement ||
      (event.target instanceof HTMLElement && event.target.isContentEditable);
    if (!textTarget) {
      if (modifier && key === 'z') action = event.shiftKey ? e.redo : e.undo;
      if (modifier && key === 'y') action = e.redo;
      if (modifier && key === 'b' && e.canWrite)
        action = () => e.toggleFormat('font-weight', 'bold');
      if (modifier && key === 'i' && e.canWrite)
        action = () => e.toggleFormat('font-style', 'italic');
      if (key === 'contextmenu' || (key === 'f10' && event.shiftKey))
        action = () => {
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
  const surfaceKey = JSON.stringify([
    e.sheet.id,
    !e.canWrite,
    e.sheet.columns,
    e.sheet.rows,
    e.sheet.mergeCells,
    e.sheet.style,
    e.sheet.freezeColumns,
  ]);
  const titleForDialog =
    e.dialog === 'rename-sheet'
      ? 'Rename worksheet'
      : e.dialog === 'row-height'
        ? 'Row height'
        : e.dialog === 'column-width'
          ? 'Column width'
          : e.dialog === 'import'
            ? 'Import workbook'
            : 'Export workbook';
  const activeFormat = e.sheet.formats?.[formatSpreadsheetAddress(e.selection.x1, e.selection.y1)];
  const menus = editorMenus(e);
  return (
    <section
      ref={root}
      className={['se-editor', props.className].filter(Boolean).join(' ')}
      aria-label={props['aria-label'] ?? 'Spreadsheet editor'}
      data-mode={e.mode}
      data-calculation-mode={props.calculation?.mode ?? 'client'}
      data-pending={e.calculation.pending}
      onKeyDownCapture={onKeyDown}
    >
      <EditorToolbar editor={e} />
      <div
        className="se-editor__formula-bar"
        onFocusCapture={() => e.grid.current?.releaseKeyboard()}
      >
        <FormulaBar
          cell={{ row: e.selection.y1, column: e.selection.x1 }}
          value={e.draftValue}
          isReadOnly={!e.canWrite}
          controls={{
            reference: (
              <TextField
                className="se-editor__address"
                label="Selected cell or range"
                value={addressDraft ?? address}
                onChange={setAddressDraft}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    jumpToAddress();
                  }
                }}
              />
            ),
            editor: (
              <div className="se-editor__formula-controls">
                <TextField
                  label="Cell value or formula"
                  value={e.draftValue}
                  onChange={e.setDraftValue}
                  isReadOnly={!e.canWrite}
                  completions={COMPLETIONS}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter' && !event.defaultPrevented) {
                      event.preventDefault();
                      e.attempt(e.applyFormula);
                    }
                  }}
                />
                <Button
                  onPress={() => e.attempt(e.applyFormula)}
                  isDisabled={!e.canWrite || !e.draftDirty}
                >
                  Apply
                </Button>
              </div>
            ),
          }}
        />
        <output aria-label="Calculated cell value">
          {e.calculation.pending
            ? 'Calculating…'
            : e.calculation.error
              ? 'Unavailable'
              : formatWorksheetCell(
                  e.output?.[e.selection.y1]?.[e.selection.x1] ?? null,
                  activeFormat,
                  props.display,
                )}
        </output>
      </div>
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
            isReplaceDisabled={!e.canWrite}
            isReplaceAllDisabled={!e.canWrite}
            onClose={() => e.setShowFind(false)}
          />
        </div>
      )}
      <div className="se-editor__workspace">
        <div className="se-editor__main">
          <div
            className="se-editor__grid"
            role="group"
            aria-label="Workbook grid"
            tabIndex={0}
            data-gridlines={e.showGridlines}
            style={{ '--se-zoom': e.zoom / 100 } as CSSProperties}
          >
            <Suspense fallback={<p role="status">Loading spreadsheet…</p>}>
              <WorkbookGrid
                key={surfaceKey}
                worksheet={e.sheet}
                values={e.output}
                display={props.display}
                readOnly={!e.canWrite}
                reviewMode={e.canSuggest}
                showFormulas={e.showFormulas}
                comments={e.workbook.comments}
                suggestions={e.workbook.suggestions}
                selection={e.selection}
                onController={(controller) => {
                  e.grid.current = controller;
                }}
                unavailableText={e.calculation.error ? 'Unavailable' : '…'}
                onChange={(next) => e.attempt(() => e.updateSheet(next))}
                onSelectionChange={e.onGridSelection}
                onContextMenu={openContext}
                onCommentRequest={() => e.openPanel('comments')}
              />
            </Suspense>
          </div>
          <div
            className="se-editor__sheets"
            role="group"
            aria-label="Worksheets"
            onFocusCapture={() => e.grid.current?.releaseKeyboard()}
          >
            {e.workbook.worksheets.map((sheet) => (
              <ContextMenuTrigger key={sheet.id} aria-label={`${sheet.name} worksheet actions`}>
                <Button
                  isSelected={sheet.id === e.sheet.id}
                  onPress={() => e.activateSheet(sheet.id)}
                  render={(domProps) => (
                    <button
                      {...domProps}
                      onContextMenu={(event) => {
                        e.activateSheet(sheet.id);
                        domProps.onContextMenu?.(event);
                      }}
                    />
                  )}
                >
                  {sheet.name}
                </Button>
                <Menu aria-label="Worksheet actions" items={menus.sheets} />
              </ContextMenuTrigger>
            ))}
          </div>
        </div>
        <div
          className="se-editor__sidebar"
          onFocusCapture={() => e.grid.current?.releaseKeyboard()}
        >
          <DockPanel
            label="Spreadsheet panels"
            selectedKey={e.panel}
            onSelectionChange={(key) => e.openPanel(String(key))}
            isExpanded={e.panelOpen}
            onExpandedChange={e.setPanelOpen}
            defaultSize={340}
            minSize={260}
            maxSize={600}
          >
            <DockPanelPane id="properties" label="Properties" icon={InfoGlyph}>
              <PropertiesPanel editor={e} />
            </DockPanelPane>
            <DockPanelPane id="history" label="History" icon={ClockGlyph}>
              <HistoryPanel editor={e} />
            </DockPanelPane>
            <DockPanelPane
              id="comments"
              label="Comments"
              icon={ReplyGlyph}
              count={e.workbook.comments?.filter((thread) => !thread.resolved).length}
            >
              <CommentsPanel editor={e} />
            </DockPanelPane>
            <DockPanelPane
              id="suggestions"
              label="Suggestions"
              icon={EditGlyph}
              count={e.workbook.suggestions?.filter((item) => item.status === 'pending').length}
            >
              <SuggestionsPanel editor={e} />
            </DockPanelPane>
            {props.panels?.map((panel) => (
              <DockPanelPane key={panel.id} id={panel.id} label={panel.label} icon={InfoGlyph}>
                {panel.content}
              </DockPanelPane>
            ))}
          </DockPanel>
        </div>
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
            isDisabled={e.calculation.pending || !!e.calculation.error}
          >
            Insert as document table
          </Button>
          <Button onPress={() => e.attempt(() => e.insert('spreadsheet'))}>
            Insert as live spreadsheet
          </Button>
        </div>
      )}
      <footer
        className="se-editor__statusbar"
        onFocusCapture={() => e.grid.current?.releaseKeyboard()}
      >
        <p
          className="se-editor__status"
          aria-label="Calculation status"
          role={e.calculation.error ? 'alert' : 'status'}
        >
          {e.calculation.error ??
            (e.calculation.pending
              ? 'Calculating…'
              : `Calculated ${props.calculation?.mode === 'server' ? 'on server' : 'on client'}`)}
        </p>
        <span className="se-editor__summary" aria-label="Selection summary">
          Count: {e.summary.count} · Sum: {e.summary.sum} · Average: {e.summary.average}
        </span>
        <ZoomControl
          value={e.zoom}
          onChange={e.setZoom}
          levels={[50, 75, 90, 100, 110, 125, 150, 175, 200]}
          resetValue={100}
          aria-label="Spreadsheet zoom"
        />
      </footer>
      {e.notice && (
        <div className="se-editor__notice" role="alert">
          <p>{e.notice}</p>
          <Button onPress={() => e.setNotice(null)}>Dismiss</Button>
        </div>
      )}
      <ContextMenu controller={context} triggerRef={root} aria-label="Spreadsheet context menu" />
      <Modal
        isOpen={e.dialog !== null}
        onOpenChange={(open) => {
          if (!open) e.setDialog(null);
        }}
        title={titleForDialog}
        primaryAction={{
          label: e.dialog === 'export' ? 'Done' : e.dialog === 'import' ? 'Import' : 'Apply',
          onPress: () => e.applyDialog(),
        }}
        cancelAction={e.dialog === 'export' ? null : {}}
        description={
          e.dialog === 'import'
            ? 'Paste CSV data or a serialized workbook. All digits in CSV numeric fields are preserved.'
            : undefined
        }
      >
        <div onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
          <TextField
            label={
              e.dialog === 'rename-sheet'
                ? 'Worksheet name'
                : e.dialog === 'import'
                  ? 'CSV or workbook JSON'
                  : e.dialog === 'export'
                    ? 'Workbook JSON'
                    : 'Size in pixels'
            }
            value={e.dialogValue}
            onChange={e.setDialogValue}
            multiline={e.dialog === 'import' || e.dialog === 'export'}
            isReadOnly={e.dialog === 'export'}
            minRows={8}
            maxRows={18}
          />
        </div>
      </Modal>
    </section>
  );
}
