'use client';

import type {
  ClipboardEvent,
  CSSProperties,
  FocusEvent,
  KeyboardEvent,
  MouseEvent,
  PointerEvent,
  UIEvent,
} from 'react';
import { selectionAddress } from '../model/editing.js';
import { selectionLock } from '../model/presence.js';
import {
  type GridFrame,
  type GridSurfaceContext,
  pickCell,
  type WorkbookGridProps,
} from './grid-context.js';
import { gridViewport } from './grid-geometry.js';
import { deliveredHeader } from './header-resize.js';
import {
  focusJSpreadsheetWorksheet,
  releaseJSpreadsheetKeyboardOwner,
} from './jspreadsheet-measures.js';

type Surface = HTMLDivElement;
/** What the surface shows now: its outlines and the clip box they are drawn in. */
interface SurfaceView {
  props: WorkbookGridProps;
  frames: GridFrame[];
  clip: CSSProperties;
  setHoveredPeer: (key: string | null) => void;
}

const composing = (context: GridSurfaceContext, event: KeyboardEvent<Surface>) =>
  context.nativeComposing.current ||
  event.nativeEvent.isComposing ||
  event.nativeEvent.keyCode === 229;

const cellNavigation = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
  'Home',
  'End',
  'Enter',
  'Tab',
]);

function ownKeyboard(context: GridSurfaceContext, target: EventTarget | null) {
  const sheet = context.worksheet();
  if (!sheet) return false;
  const interactive =
    target instanceof Element &&
    target.closest('button, input, textarea, select, a[href], [contenteditable="true"]');
  if (interactive && !context.inCell.current) {
    releaseJSpreadsheetKeyboardOwner(sheet);
    return false;
  }
  focusJSpreadsheetWorksheet(sheet);
  return true;
}

function releaseOnBlur(context: GridSurfaceContext, event: FocusEvent<Surface>) {
  if (event.relatedTarget instanceof Node && event.currentTarget.contains(event.relatedTarget))
    return;
  context.controller.releaseKeyboard();
}

function keyDown(context: GridSurfaceContext, event: KeyboardEvent<Surface>) {
  if (!ownKeyboard(context, event.target)) return;
  const inCell = context.inCell.current;
  if (inCell && composing(context, event)) {
    // CE's document listener treats every Enter as a commit. Leave the
    // browser's input method in control without forwarding the key to CE.
    event.stopPropagation();
    return;
  }
  if (inCell) context.native.onKey(event);
  else if (cellNavigation.has(event.key)) {
    context.selectedHeader.current = null;
    if (event.key === 'Tab') {
      // A grid is one stop in the surrounding page's tab order. The vendor uses arrows to navigate.
      event.stopPropagation();
      return;
    }
  }
  if (!context.inCell.current && ['Delete', 'Backspace'].includes(event.key)) {
    event.preventDefault();
    event.stopPropagation();
    if (!context.latest.current.readOnly)
      context.latest.current.onDelete(context.selectedHeader.current);
    return;
  }
  if (context.resize.active && event.key === 'Escape') {
    event.preventDefault();
    event.stopPropagation();
    context.resize.cancel();
  }
}

/** A peer's edit lock blocks opening the cell's editor. */
function doubleClick(
  context: GridSurfaceContext,
  props: WorkbookGridProps,
  event: MouseEvent<Surface>,
) {
  if (context.commentClick.current) clearTimeout(context.commentClick.current);
  const cell = pickCell(context, event);
  if (!cell) return;
  const locked = selectionLock(
    { schemaVersion: 1, worksheets: [props.worksheet] },
    { sheetId: props.worksheet.id, selection: { x1: cell.x, x2: cell.x, y1: cell.y, y2: cell.y } },
    props.collaborators ?? [],
  );
  if (!locked) return;
  event.preventDefault();
  event.stopPropagation();
  props.onBlockedAction?.(`${locked.name} is editing this cell.`);
}

/** A single click on a marked cell opens its review, unless a double click follows. */
function click(context: GridSurfaceContext, event: MouseEvent<Surface>) {
  if (context.commentClick.current) clearTimeout(context.commentClick.current);
  if (context.inCell.current || context.picking.current || event.detail !== 1 || !event.isTrusted)
    return;
  const coords = pickCell(context, event),
    cell = coords && context.worksheet()?.getCellFromCoords?.(coords.x, coords.y);
  if (!coords || (!cell?.dataset.seComments && !cell?.dataset.seSuggestions)) return;
  context.commentClick.current = setTimeout(() => {
    const current = context.latest.current;
    if (
      !context.inCell.current &&
      current.selection.x1 === coords.x &&
      current.selection.y1 === coords.y
    )
      current.onCommentRequest(coords);
  }, 220);
}

function contextMenu(context: GridSurfaceContext, event: MouseEvent<Surface>) {
  const header = deliveredHeader(event, context.surface.current);
  if (header) {
    event.preventDefault();
    event.stopPropagation();
    context.latest.current.onContextMenu(
      header.axis === 'row' ? { x: -1, y: header.index } : { x: header.index, y: -1 },
      header.element.getBoundingClientRect(),
    );
    return;
  }
  // Own cell context menus before the vendor or browser handles the same event.
  // An active native text editor retains its text-editing context menu.
  if (
    event.target instanceof Element &&
    event.target.closest('input, textarea, [contenteditable="true"]')
  ) {
    event.stopPropagation();
    return;
  }
  const coords = pickCell(context, event);
  const cell = coords && context.worksheet()?.getCellFromCoords?.(coords.x, coords.y);
  if (!coords || !cell) return;
  event.preventDefault();
  event.stopPropagation();
  context.latest.current.onContextMenu(coords, cell.getBoundingClientRect());
}

function scroll(context: GridSurfaceContext, event: UIEvent<Surface>) {
  const sheet = context.worksheet();
  if (sheet && event.target === gridViewport(sheet)) context.scrolled.current = true;
  context.scheduleReferences();
}

/** Native text editing keeps its own clipboard. Outside it CE must not alter the clipboard block. */
function clipboard(context: GridSurfaceContext, event: ClipboardEvent<Surface>) {
  if (
    event.target instanceof Element &&
    event.target.closest('input, textarea, [contenteditable="true"]')
  )
    return;
  event.preventDefault();
  event.stopPropagation();
  const current = context.latest.current;
  if (event.type === 'paste') current.onClipboardPaste(event.clipboardData);
  else current.onClipboardCopy(event.clipboardData, event.type === 'cut');
}

/** Focuses the surface for header and cell presses outside formula picking and editing. */
function focusOnPress(context: GridSurfaceContext, event: PointerEvent<Surface>) {
  if (event.button !== 0 || context.latest.current.pickingFormula) return;
  const header = deliveredHeader(event, context.surface.current);
  context.selectedHeader.current = header?.axis ?? null;
  if (!context.inCell.current && (header || pickCell(context, event)))
    event.currentTarget.focus({ preventScroll: true });
}

function pointerDown(context: GridSurfaceContext, event: PointerEvent<Surface>) {
  if (context.resize.start(event)) return;
  if (event.button === 0 && !context.latest.current.pickingFormula && !context.inCell.current) {
    context.draggingSelection.current = true;
    context.selectionPublication.start();
  }
  focusOnPress(context, event);
  if (!context.latest.current.pickingFormula || event.button !== 0 || !event.isTrusted) return;
  const cell = pickCell(context, event);
  if (!cell) return;
  event.preventDefault();
  event.stopPropagation();
  context.picking.current = { x1: cell.x, y1: cell.y, x2: cell.x, y2: cell.y };
  event.currentTarget.setPointerCapture(event.pointerId);
}

/** The peer selection outline under a point of the clip box. */
function hoveredFrame(frames: GridFrame[], { x, y }: { x: number; y: number }) {
  return frames.find(
    (frame) =>
      frame.kind === 'selection' &&
      x >= Number(frame.style.left) &&
      x <= Number(frame.style.left) + Number(frame.style.width) &&
      y >= Number(frame.style.top) &&
      y <= Number(frame.style.top) + Number(frame.style.height),
  );
}

function pointerMove(context: GridSurfaceContext, view: SurfaceView, event: PointerEvent<Surface>) {
  if (context.resize.move(event)) return;
  const base = context.surface.current?.getBoundingClientRect();
  if (base) {
    const x = event.clientX - base.left - Number(view.clip.left ?? 0),
      y = event.clientY - base.top - Number(view.clip.top ?? 0);
    view.setHoveredPeer(hoveredFrame(view.frames, { x, y })?.key ?? null);
  }
  if (!context.picking.current) return;
  const cell = pickCell(context, event);
  if (cell) context.picking.current = { ...context.picking.current, x2: cell.x, y2: cell.y };
}

function endDrag(context: GridSurfaceContext) {
  context.draggingSelection.current = false;
  context.selectionPublication.finish();
}

function pointerUp(context: GridSurfaceContext, event: PointerEvent<Surface>) {
  endDrag(context);
  if (context.resize.finish(event)) return;
  const picked = context.picking.current;
  if (!picked) return;
  event.preventDefault();
  event.stopPropagation();
  if (context.inCell.current) context.native.insert(selectionAddress(picked));
  else context.latest.current.onFormulaReference?.(selectionAddress(picked));
  context.picking.current = null;
  event.currentTarget.releasePointerCapture(event.pointerId);
}

/** The surface's event handlers, over the shared grid context. */
export function useSurfaceHandlers(context: GridSurfaceContext, view: SurfaceView) {
  return {
    onKeyDownCapture: (event: KeyboardEvent<Surface>) => keyDown(context, event),
    onFocusCapture: (event: FocusEvent<Surface>) => ownKeyboard(context, event.target),
    onBlurCapture: (event: FocusEvent<Surface>) => releaseOnBlur(context, event),
    onDoubleClickCapture: (event: MouseEvent<Surface>) => doubleClick(context, view.props, event),
    onClick: (event: MouseEvent<Surface>) => click(context, event),
    onContextMenuCapture: (event: MouseEvent<Surface>) => contextMenu(context, event),
    onScrollCapture: (event: UIEvent<Surface>) => scroll(context, event),
    onInputCapture: (event: Parameters<GridSurfaceContext['native']['onInput']>[0]) => {
      if (context.inCell.current) context.native.onInput(event);
    },
    onCompositionStartCapture: () => {
      if (context.inCell.current) context.nativeComposing.current = true;
    },
    onCompositionEndCapture: () => {
      context.nativeComposing.current = false;
    },
    onCopyCapture: (event: ClipboardEvent<Surface>) => clipboard(context, event),
    onCutCapture: (event: ClipboardEvent<Surface>) => clipboard(context, event),
    onPasteCapture: (event: ClipboardEvent<Surface>) => clipboard(context, event),
    onPointerDownCapture: (event: PointerEvent<Surface>) => pointerDown(context, event),
    onPointerMove: (event: PointerEvent<Surface>) => pointerMove(context, view, event),
    onPointerUp: (event: PointerEvent<Surface>) => pointerUp(context, event),
    onPointerLeave: () => view.setHoveredPeer(null),
    onPointerCancel: () => {
      endDrag(context);
      context.resize.cancel();
      context.picking.current = null;
    },
    onLostPointerCapture: () => endDrag(context),
  };
}
