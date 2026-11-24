'use client';

import { Tabs, TabsContent } from '@flamboyantengineer/ui-components/tabs';
import { ZoomControl } from '@flamboyantengineer/ui-components/zoom-control';
import { lazy, Suspense, useState } from 'react';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
} from '../engine/jspreadsheet-types.js';
import { worksheetExtent } from '../model/extent.js';
import type { FormulaReferenceHighlight } from '../model/formula-reference-parser.js';
import { openCellComment } from './editor-comment-bubble.js';
import { SheetBar } from './sheet-bar.js';
import type { SpreadsheetEditorController } from './use-editor.js';

const WorkbookGrid = lazy(() =>
  import('../engine/workbook-grid.js').then((module) => ({ default: module.WorkbookGrid })),
);

/** Formula reference picking: the draft's references and where a picked reference goes. */
interface ReferencePicking {
  picking: boolean;
  references: FormulaReferenceHighlight[];
  onFormulaReference(range: string): void;
  onEditStateChange(editing: boolean): void;
}

interface GridProps {
  editor: SpreadsheetEditorController;
  picking: ReferencePicking;
  onContextMenu(coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect): void;
  onViewportChange(): void;
}

/** The active worksheet's grid, wired to the editor. Grids mounted after the first take focus. */
function EditorGrid({ editor: e, picking, onContextMenu, onViewportChange }: GridProps) {
  const { props } = e;
  // Grids mounted after the first one may take focus (SE-UI-3.1).
  const [gridMounted, setGridMounted] = useState(false);
  return (
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
      focusOnMount={gridMounted}
      onNativeDraft={e.setNativeDraft}
      onError={(error) =>
        e.attempt(() => {
          throw error;
        })
      }
      formulaReferences={picking.references}
      pickingFormula={picking.picking}
      onFormulaReference={picking.onFormulaReference}
      collaborators={e.peers}
      showContributors={e.showContributors}
      onViewportChange={onViewportChange}
      onEditStateChange={picking.onEditStateChange}
      onDelete={(axis) => e.attempt(() => (axis ? e.structure(axis, 'delete') : e.clear()))}
      onEditDraft={(value) => (value === null ? e.cancelFormula() : e.setDraftValue(value))}
      display={props.display}
      readOnly={!e.canWrite}
      reviewMode={e.canSuggest}
      onReviewChange={e.proposeCells}
      onClipboardCopy={e.transferCopy}
      onClipboardPaste={e.transferPaste}
      onBlockedAction={e.warn}
      showFormulas={e.showFormulas}
      comments={e.workbook.comments}
      suggestions={e.workbook.suggestions}
      selection={e.selection}
      onController={(controller) => {
        e.attachGrid(controller);
        if (controller) setGridMounted(true);
        if (e.dialog) controller?.releaseKeyboard();
      }}
      unavailableText={e.calculation.error ? 'Unavailable' : '…'}
      onChange={e.updateSheet}
      onSelectionChange={e.onGridSelection}
      onContextMenu={onContextMenu}
      onCommentRequest={(coords) => openCellComment(e, coords)}
    />
  );
}

/** The worksheet tabs around the active worksheet's grid, with the sheet bar and zoom. */
export function EditorWorkspace(props: GridProps) {
  const { editor: e } = props;
  return (
    <Tabs
      className="se-editor__main se-editor__worksheet-tabs"
      selectedKey={e.sheet.id}
      onSelectionChange={(id) => e.attempt(() => e.activateSheet(String(id)))}
      keyboardActivation="manual"
    >
      <TabsContent id={e.sheet.id} className="se-editor__worksheet-panel">
        <div className="se-editor__grid" data-gridlines={e.showGridlines}>
          <Suspense fallback={<p role="status">Loading spreadsheet…</p>}>
            <EditorGrid {...props} />
          </Suspense>
        </div>
      </TabsContent>
      <SheetBar
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
      </SheetBar>
    </Tabs>
  );
}
