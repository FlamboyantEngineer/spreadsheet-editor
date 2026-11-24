'use client';

import { type Dispatch, type SetStateAction, useEffect, useLayoutEffect } from 'react';
import { gridWorksheetExtent } from '../model/extent.js';
import { renderContributors, renderReviewMarkers } from './cell-markers.js';
import { type CellPortal, renderCalculatedCells } from './cell-renderer.js';
import type { GridSurfaceContext, WorkbookGridProps } from './grid-context.js';
import { selectGridRange } from './grid-controller.js';
import { gridViewport, gridZoomScale, restoreGridScroll } from './grid-geometry.js';
import type {
  JSpreadsheetHostHandle,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';
import { refreshGridRowWindow } from './row-window.js';
import { patchWorksheetData, syncWorksheetData } from './worksheet-data.js';
import { syncWorksheetPresentation } from './worksheet-presentation.js';

type SyncContext = Omit<GridSurfaceContext, 'native' | 'resize'>;
const paintedProps = new WeakMap<JSpreadsheetHostHandle, WorkbookGridProps>();

/** Paints review markers, calculated cells and contributor colours from the given props. */
export function paint(
  api: JSpreadsheetHostHandle,
  current: WorkbookGridProps,
  setPortals: Dispatch<SetStateAction<CellPortal[]>>,
) {
  paintedProps.set(api, current);
  // Markers rewrite cell titles, which presentation tooltips share.
  const markersChanged = renderReviewMarkers(api, current.worksheet.id, {
    comments: current.comments,
    suggestions: current.suggestions,
  });
  const portalList = renderCalculatedCells(api, current.values, {
    display: current.display,
    raw: current.worksheet.data,
    unavailableText: current.unavailableText,
    formats: current.worksheet.formats,
    showFormulas: current.showFormulas,
    presentation: {
      sheetId: current.worksheet.id,
      draft: current.draft,
      readyCells: current.readyCells,
      publication: current.publication,
      cellPresentation: current.cellPresentation,
      renderCellContent: current.renderCellContent,
    },
    repaint: markersChanged,
  });
  // An unchanged empty list keeps its identity, so painting causes no render.
  setPortals((previous) =>
    previous.length === portalList.length &&
    previous.every((portal, index) => {
      const next = portalList[index];
      return (
        portal.key === next.key && portal.target === next.target && portal.content === next.content
      );
    })
      ? previous
      : portalList,
  );
  renderContributors(api, current.worksheet, {
    visible: current.showContributors ?? false,
    editing: current.pickingFormula ? current.selection : undefined,
  });
}

/** Paints the currently mounted geometry while a protected edit defers replacements. */
export function paintCurrentGrid(context: SyncContext, props: WorkbookGridProps) {
  const handle = context.handle.current;
  if (!handle) return;
  const protectedEdit = context.inCell.current || props.pickingFormula;
  const current =
    protectedEdit && context.applied.current !== props.worksheet ? paintedProps.get(handle) : props;
  if (current) paint(handle, current, context.setPortals);
}

/** Writes the worksheet's data into the engine, in place when only values changed. */
function syncData(
  context: SyncContext,
  sheet: JSpreadsheetWorksheetInstance,
  props: WorkbookGridProps,
) {
  const nextExtent = gridWorksheetExtent(props.worksheet, props.logicalExtent);
  const oldExtent = context.appliedExtent.current;
  const extentChanged = !oldExtent || nextExtent.some((value, axis) => value !== oldExtent[axis]);
  const applied = context.applied.current;
  // Changed values are written in place; rebuilding every row restyles every cell.
  if (
    extentChanged ||
    (applied?.data !== props.worksheet.data &&
      (!applied || !patchWorksheetData(sheet, props.worksheet, props.logicalExtent)))
  )
    syncWorksheetData(sheet, props.worksheet, props.logicalExtent);
  return nextExtent;
}

/** Shows the props' worksheet in the engine, keeping the scroll anchor across a zoom. */
function applyWorksheet(
  context: SyncContext,
  { sheet, api }: { sheet: JSpreadsheetWorksheetInstance; api: JSpreadsheetHostHandle },
  props: WorkbookGridProps,
) {
  const nextExtent = syncData(context, sheet, props);
  const viewport = gridViewport(sheet);
  const scroll = { left: viewport?.scrollLeft ?? 0, top: viewport?.scrollTop ?? 0 };
  const scale = gridZoomScale(props.zoom),
    previousScale = gridZoomScale(context.appliedZoom.current);
  const zoomOnly = context.applied.current === props.worksheet && scale !== previousScale;
  syncWorksheetPresentation(
    sheet,
    { previous: context.applied.current, next: props.worksheet },
    { readOnly: !!props.readOnly, zoom: props.zoom, logicalExtent: props.logicalExtent },
  );
  context.applied.current = props.worksheet;
  context.appliedExtent.current = nextExtent;
  context.appliedZoom.current = props.zoom ?? 100;
  if (scale !== previousScale)
    restoreGridScroll(viewport, scroll, { from: previousScale, to: scale });
  refreshGridRowWindow(sheet);
  paint(api, props, context.setPortals);
  if (!zoomOnly && !context.draggingSelection.current)
    selectGridRange(context, props.selection, false);
  context.scheduleReferences();
}

function syncEngine(context: SyncContext, props: WorkbookGridProps) {
  const sheet = context.worksheet(),
    api = context.handle.current;
  if (!api || !sheet) return;
  if (sheet.options) sheet.options.editable = !props.readOnly;
  if (props.readOnly && context.inCell.current) context.controller.cancelDraft?.();
  if (context.inCell.current || props.pickingFormula) return;
  context.suppress.current = true;
  try {
    applyWorksheet(context, { sheet, api }, props);
  } finally {
    context.suppress.current = false;
  }
}

/** Follows a selection the host changed, unless a pointer drag is selecting. */
function followSelection(context: SyncContext, props: WorkbookGridProps) {
  const current = context.visualSelection.current,
    next = props.selection;
  if (
    context.handle.current &&
    !context.draggingSelection.current &&
    (current.x1 !== next.x1 ||
      current.y1 !== next.y1 ||
      current.x2 !== next.x2 ||
      current.y2 !== next.y2)
  )
    context.controller.select(next);
}

/** Keeps the engine showing the props: worksheet, selection and painted presentation. */
export function useWorksheetSync(context: SyncContext, props: WorkbookGridProps, revision: number) {
  useLayoutEffect(() => {
    syncEngine(context, props);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- syncs only on these inputs; the rest are read through refs or are stable for the grid's life
  }, [
    props.worksheet,
    props.readOnly,
    props.pickingFormula,
    props.zoom,
    props.logicalExtent?.columns,
    props.logicalExtent?.rows,
    revision,
  ]);
  const { controller } = context;
  useLayoutEffect(() => {
    followSelection(context, props);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- follows the host's selection; the rest is read through refs
  }, [props.selection, controller]);
  useEffect(() => {
    paintCurrentGrid(context, props);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- repaints only when what it paints changes
  }, [
    props.values,
    props.display,
    props.worksheet,
    props.unavailableText,
    props.comments,
    props.suggestions,
    props.showFormulas,
    props.showContributors,
    // A draft is a new object on each render; only its cell and text matter.
    props.draft?.x,
    props.draft?.y,
    props.draft?.text,
    props.readyCells,
    props.publication,
    props.cellPresentation,
    props.renderCellContent,
    revision,
  ]);
}
