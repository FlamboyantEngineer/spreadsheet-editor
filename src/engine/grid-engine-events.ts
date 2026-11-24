'use client';

import { gridWorksheetExtent } from '../model/extent.js';
import { clearCalculatedPresentation } from './cell-renderer.js';
import type { GridSurfaceContext } from './grid-context.js';
import { gridViewport, placeGridCorner } from './grid-geometry.js';
import { paint, paintCurrentGrid } from './grid-sync.js';
import {
  focusJSpreadsheetWorksheet,
  releaseJSpreadsheetKeyboardOwner,
} from './jspreadsheet-measures.js';
import type {
  JSpreadsheetHostHandle,
  JSpreadsheetSelectionRange,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';
import { refreshGridRowWindow, revealGridRows, windowGridRows } from './row-window.js';
import { observeGridBorderCopies, syncSharedGridBorders } from './shared-grid-borders.js';
import { presentAttachedRows, syncWorksheetPresentation } from './worksheet-presentation.js';

type Context = GridSurfaceContext;

/** Windows the engine's rows, presenting each row the window attaches. */
function startRowWindow(context: Context, sheet: JSpreadsheetWorksheetInstance | null) {
  context.stopRowWindow.current?.();
  context.stopRowWindow.current = null;
  if (!sheet?.options) return;
  // The engine's own lazy mode stays off: its scrollbar spans only the rows it holds.
  (sheet.options as { lazyLoading?: boolean }).lazyLoading = false;
  context.stopRowWindow.current = windowGridRows(sheet, (rows) => {
    presentAttachedRows(sheet, rows);
    syncSharedGridBorders(sheet);
    const corner = Number(sheet.highlighted?.at(-1)?.element?.dataset.y);
    if (rows.includes(corner)) placeGridCorner(sheet);
    if (!context.suppress.current) paintCurrentGrid(context, context.latest.current);
  });
}

/** Presents the mounted engine's worksheet, and records what it shows. */
function presentMounted(context: Context, sheet: JSpreadsheetWorksheetInstance | null) {
  const current = context.latest.current;
  if (sheet)
    syncWorksheetPresentation(
      sheet,
      { previous: null, next: current.worksheet },
      { readOnly: !!current.readOnly, zoom: current.zoom, logicalExtent: current.logicalExtent },
    );
  if (sheet) refreshGridRowWindow(sheet);
  context.appliedZoom.current = current.zoom ?? 100;
  context.applied.current = current.worksheet;
  context.appliedExtent.current = gridWorksheetExtent(current.worksheet, current.logicalExtent);
}

/** A lazy grid mount must not interrupt a press or draft in the chrome. */
function focusAfterMount(context: Context) {
  if (!context.focusOnMount.current) return;
  context.focusRequest.current = requestAnimationFrame(() => {
    const active = document.activeElement;
    const surface = context.surface.current;
    const editor = surface?.closest('.se-editor');
    if (
      active === document.body ||
      (active && surface?.contains(active)) ||
      (active?.getAttribute('role') === 'tab' && active.closest('.se-editor') === editor)
    )
      context.controller.focus();
  });
}

export function engineMounted(context: Context, api: JSpreadsheetHostHandle) {
  context.handle.current = api;
  const sheet = context.worksheet();
  const surface = context.surface.current;
  const engine = surface?.querySelector<HTMLElement>('[data-jspreadsheet-host]');
  const viewport = sheet && gridViewport(sheet);
  if (engine) engine.tabIndex = -1;
  if (viewport) viewport.tabIndex = -1;
  if (sheet && surface?.contains(surface.ownerDocument.activeElement))
    focusJSpreadsheetWorksheet(sheet);
  else releaseJSpreadsheetKeyboardOwner(sheet);
  context.stopBorderCopies.current?.();
  context.stopBorderCopies.current = sheet ? observeGridBorderCopies(sheet) : null;
  startRowWindow(context, sheet);
  presentMounted(context, sheet);
  paint(api, context.latest.current, context.setPortals);
  context.controller.select(context.latest.current.selection);
  const { initialViewport } = context.latest.current;
  if (initialViewport) context.controller.restoreViewport?.(initialViewport);
  focusAfterMount(context);
  context.scheduleReferences();
}

export function engineDestroying(context: Context) {
  context.stopRowWindow.current?.();
  context.stopRowWindow.current = null;
  context.selectionPublication.dispose();
  releaseJSpreadsheetKeyboardOwner(context.worksheet());
  context.native.end();
  if (context.handle.current) clearCalculatedPresentation(context.handle.current);
}

/**
 * CE selects merge anchors while constructing or changing merges. Those
 * presentation callbacks must not replace the host's current selection.
 */
export function engineSelected(context: Context, range: JSpreadsheetSelectionRange) {
  if (!context.applied.current || context.suppress.current || context.selecting.current) return;
  // Keyboard moves can leave the attached rows; the engine measures the new cell
  // to scroll to it right after this callback.
  const sheet = context.worksheet(),
    cursor = Number(sheet?.selectedCell?.[3]);
  if (sheet && Number.isInteger(cursor)) revealGridRows(sheet, cursor);
  context.visualSelection.current = range;
  context.scheduleReferences();
  context.selectionPublication.update(range);
}

export function engineEditStarted(context: Context, { x, y }: { x: number; y: number }) {
  context.inCell.current = true;
  context.nativeCoords.current = { x, y };
  context.native.start(context.worksheet());
  context.latest.current.onEditStateChange?.(true);
}

export function engineEditEnded(context: Context) {
  context.inCell.current = false;
  context.nativeComposing.current = false;
  context.native.end();
  context.nativeCoords.current = null;
  context.latest.current.onEditStateChange?.(false);
  context.surface.current?.focus({ preventScroll: true });
  context.setRevision((value) => value + 1);
}
