'use client';

import { useMenuController } from '@flamboyantengineer/ui-components/context-menu';
import { useEffect, useEffectEvent, useState } from 'react';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
} from '../engine/jspreadsheet-types.js';
import type { CellSelection } from '../model/editing.js';
import { gridWorksheetExtent, worksheetExtent } from '../model/extent.js';
import { editorMenus } from './editor-menus.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type ContextKind = 'cell' | 'row' | 'column';

const outside = (value: number, start: number, end: number) => value < start || value > end;

/** The selection a context click makes: its row, column or cell unless already selected. */
function contextSelection(
  { selection, previewSheet, workbook }: SpreadsheetEditorController,
  { x, y }: JSpreadsheetCellCoords,
): CellSelection | null {
  if (x < 0 && y >= 0) {
    const [columns] = gridWorksheetExtent(previewSheet, worksheetExtent(previewSheet, workbook));
    return outside(y, selection.y1, selection.y2)
      ? {
          x1: 0,
          x2: columns - 1,
          y1: y,
          y2: y,
        }
      : null;
  }
  if (y < 0 && x >= 0) {
    const [, rows] = gridWorksheetExtent(previewSheet, worksheetExtent(previewSheet, workbook));
    return outside(x, selection.x1, selection.x2)
      ? {
          x1: x,
          x2: x,
          y1: 0,
          y2: rows - 1,
        }
      : null;
  }
  const outsideCell =
    outside(x, selection.x1, selection.x2) || outside(y, selection.y1, selection.y2);
  return x >= 0 && y >= 0 && outsideCell ? { x1: x, x2: x, y1: y, y2: y } : null;
}

/** The grid's context menus: a cell, row header or column header menu at the pointer or cell. */
export function useEditorContextMenu(e: SpreadsheetEditorController) {
  const context = useMenuController();
  const [contextKind, setContextKind] = useState<ContextKind>('cell');
  const [contextPoint, setContextPoint] = useState<{ x: number; y: number } | null>(null);
  const contextItems = () => {
    const menus = editorMenus(e);
    return contextKind === 'row'
      ? menus.rowContext
      : contextKind === 'column'
        ? menus.columnContext
        : menus.context;
  };
  function openContext(coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect) {
    const next = contextSelection(e, coords);
    e.attempt(() => {
      if (next && !e.select(next)) return;
      setContextKind(coords.x < 0 ? 'row' : coords.y < 0 ? 'column' : 'cell');
      e.grid.current?.releaseKeyboard();
      setContextPoint({ x: rect.left + 10, y: rect.bottom });
    });
  }
  function openCellMenu() {
    setContextKind('cell');
    const rect = e.grid.current?.getAnchor(e.selection);
    if (rect) {
      e.grid.current?.releaseKeyboard();
      setContextPoint({ x: rect.left, y: rect.bottom });
    }
  }
  const currentContextItems = useEffectEvent(contextItems);
  useEffect(() => {
    if (!contextPoint) return;
    // The menu opens after the render that selected the cell, so its items read that selection.
    context.open(contextPoint, { items: currentContextItems() });
    // eslint-disable-next-line react-hooks/set-state-in-effect -- the point is consumed once the menu has opened
    setContextPoint(null);
  }, [contextPoint, context, contextKind]);
  useEffect(() => {
    context.setItems(currentContextItems());
  }, [e.workbook, e.selection, e.mode, e.capabilities, context, contextKind]);
  return { context, openContext, openCellMenu };
}
