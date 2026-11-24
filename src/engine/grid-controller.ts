'use client';

import { useEffect, useMemo } from 'react';
import type { GridController } from '../editor/contracts.js';
import type { CellSelection } from '../model/editing.js';
import { gridCellAt } from './grid-cell.js';
import type { GridContext } from './grid-context.js';
import { gridViewport, restoreGridScroll, scrollGridCellIntoView } from './grid-geometry.js';
import {
  focusJSpreadsheetWorksheet,
  releaseJSpreadsheetKeyboardOwner,
} from './jspreadsheet-measures.js';

type ControllerContext = Pick<
  GridContext,
  | 'worksheet'
  | 'selecting'
  | 'visualSelection'
  | 'scheduleReferences'
  | 'commentClick'
  | 'selectedHeader'
  | 'surface'
  | 'inCell'
>;

/** Selects a range in the engine, scrolling it into view when asked. */
export function selectGridRange(context: ControllerContext, range: CellSelection, reveal: boolean) {
  if (context.selecting.current) return;
  if (reveal) context.selectedHeader.current = null;
  context.visualSelection.current = range;
  context.selecting.current = true;
  try {
    const sheet = context.worksheet();
    if (sheet && gridCellAt(sheet, range.x1, range.y1) && gridCellAt(sheet, range.x2, range.y2)) {
      sheet.updateSelectionFromCoords?.(range.x1, range.y1, range.x2, range.y2);
      if (reveal) scrollGridCellIntoView(sheet, range.x1, range.y1);
    }
  } finally {
    context.selecting.current = false;
  }
  context.scheduleReferences();
}

function focusGrid(context: ControllerContext) {
  const sheet = context.worksheet();
  if (!sheet) return;
  focusJSpreadsheetWorksheet(sheet);
  context.surface.current?.focus({ preventScroll: true });
  context.scheduleReferences();
}

/** Closes an open cell editor, saving its draft or not. */
function closeDraft(context: ControllerContext, save: boolean) {
  const sheet = context.worksheet();
  const cell = (sheet as { edition?: [HTMLTableCellElement] } | null)?.edition?.[0];
  if (context.inCell.current && cell) sheet?.closeEditor?.(cell, save);
}

function releaseKeyboard(context: ControllerContext) {
  if (context.commentClick.current) clearTimeout(context.commentClick.current);
  context.commentClick.current = null;
  releaseJSpreadsheetKeyboardOwner(context.worksheet());
}

function viewportOf(context: ControllerContext) {
  const sheet = context.worksheet();
  const viewport = sheet && gridViewport(sheet);
  return { left: viewport?.scrollLeft ?? 0, top: viewport?.scrollTop ?? 0 };
}

/** One controller for the grid's life, announced to the host once. */
export function useGridController(
  context: ControllerContext,
  onController: (controller: GridController | null) => void,
) {
  const controller = useMemo<GridController & { focus(): void }>(
    () => ({
      select: (range) => selectGridRange(context, range, true),
      getAnchor: (range) =>
        gridCellAt(context.worksheet(), range.x1, range.y1)?.getBoundingClientRect() ?? null,
      releaseKeyboard: () => releaseKeyboard(context),
      focus: () => focusGrid(context),
      commitDraft: () => closeDraft(context, true),
      cancelDraft: () => closeDraft(context, false),
      getViewport: () => viewportOf(context),
      restoreViewport: (position) => {
        const sheet = context.worksheet();
        if (sheet) restoreGridScroll(gridViewport(sheet), position, { from: 1, to: 1 });
      },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- one controller for the grid's life; its methods read refs only
    [],
  );
  useEffect(() => {
    onController(controller);
    return () => onController(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps -- announced once per controller; the host's callback may change every render
  }, [controller]);
  return controller;
}
