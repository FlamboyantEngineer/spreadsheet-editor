'use client';

import {
  type PointerEvent,
  type RefObject,
  useEffect,
  useEffectEvent,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { requireDefined } from '../model/require-defined.js';
import type { Worksheet } from '../model/workbook.js';
import { DEFAULT_COLUMN_WIDTH, DEFAULT_ROW_HEIGHT, gridZoomScale } from './grid-geometry.js';
import {
  applyJSpreadsheetColumnMeasure,
  applyJSpreadsheetRowMeasure,
  type JSpreadsheetColumnMeasureSnapshot,
  type JSpreadsheetRowMeasureSnapshot,
  readJSpreadsheetColumnMeasure,
  readJSpreadsheetRowMeasureSnapshot,
} from './jspreadsheet-measures.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';

interface Options {
  root: RefObject<HTMLDivElement | null>;
  worksheet: Worksheet;
  getWorksheet(): JSpreadsheetWorksheetInstance | null;
  canResize: boolean;
  zoom?: number;
  onChange(worksheet: Worksheet): void;
  onGeometryChange(): void;
  subscribeBlur?: (callback: () => void) => () => void;
}
interface Drag {
  axis: 'row' | 'column';
  pointerId: number;
  start: number;
  scale: number;
  delta: number;
  worksheet: JSpreadsheetWorksheetInstance;
  originals: {
    index: number;
    size: number;
    column?: JSpreadsheetColumnMeasureSnapshot;
    row?: JSpreadsheetRowMeasureSnapshot;
  }[];
}
type Header = { axis: 'row' | 'column'; index: number; element: HTMLElement };

const coordinate = (element: HTMLElement, axis: 'x' | 'y') =>
  element.hasAttribute(`data-${axis}`) ? Number(element.dataset[axis]) : null;

/** A header cell has a coordinate on its own axis and none on the other. */
function headerOf(element: HTMLElement): Header | null {
  const x = coordinate(element, 'x'),
    y = coordinate(element, 'y');
  if (element.tagName !== 'TD') return null;
  if (x !== null && x >= 0 && (y === null || y < 0)) return { axis: 'column', index: x, element };
  if (y !== null && y >= 0 && (x === null || x < 0)) return { axis: 'row', index: y, element };
  return null;
}

export function deliveredHeader(event: Pick<PointerEvent, 'target'>, root: HTMLElement | null) {
  let element = event.target instanceof HTMLElement ? event.target : null;
  if (!root || !element || !root.contains(element)) return null;
  while (element && element !== root) {
    const header = headerOf(element);
    if (header) return header;
    element = element.parentElement;
  }
  return null;
}

/** Puts the engine's measures back as they were before the drag. */
function restoreMeasures(current: Drag) {
  for (const original of current.originals) {
    if (original.column)
      applyJSpreadsheetColumnMeasure(current.worksheet, original.index, original.column);
    if (original.row) applyJSpreadsheetRowMeasure(current.worksheet, original.index, original.row);
  }
}

function releasePointer(root: HTMLDivElement | null, pointerId: number) {
  if (root?.hasPointerCapture(pointerId)) root.releasePointerCapture(pointerId);
}

/** The lines a drag resizes: the selected ones when it starts on one of them, else its own. */
function draggedLines(worksheet: JSpreadsheetWorksheetInstance, header: Header): number[] {
  const selected =
    header.axis === 'column'
      ? worksheet.getSelectedColumns?.(true)
      : worksheet.getSelectedRows?.(true);
  const indices = Array.isArray(selected)
    ? selected.map(Number).filter((index) => Number.isInteger(index) && index >= 0)
    : [];
  return indices.includes(header.index) ? indices : [header.index];
}

function originalsOf(
  options: Options,
  worksheet: JSpreadsheetWorksheetInstance,
  header: Header,
): Drag['originals'] {
  return draggedLines(worksheet, header).map((index) =>
    header.axis === 'column'
      ? {
          index,
          size: options.worksheet.columns?.[index]?.width ?? DEFAULT_COLUMN_WIDTH,
          column: readJSpreadsheetColumnMeasure(worksheet, index),
        }
      : {
          index,
          size: options.worksheet.rows?.[String(index)]?.height ?? DEFAULT_ROW_HEIGHT,
          row: readJSpreadsheetRowMeasureSnapshot(worksheet, index),
        },
  );
}

/** A drag that starts within six pixels of a header's trailing edge, else null. */
function beginDrag(event: PointerEvent<HTMLDivElement>, options: Options): Drag | null {
  const header = deliveredHeader(event, options.root.current),
    worksheet = options.getWorksheet();
  if (!header || !worksheet || event.button !== 0 || !options.canResize) return null;
  const rect = header.element.getBoundingClientRect(),
    edge = header.axis === 'column' ? rect.right - event.clientX : rect.bottom - event.clientY;
  if (edge < 0 || edge > 6) return null;
  event.preventDefault();
  event.stopPropagation();
  const scale = gridZoomScale(options.zoom);
  return {
    axis: header.axis,
    pointerId: event.pointerId,
    start: header.axis === 'column' ? event.clientX : event.clientY,
    scale,
    delta: 0,
    worksheet,
    originals: originalsOf(options, worksheet, header),
  };
}

const clampedSize = (size: number) => Math.max(1, Math.min(2000, size));

/** Shows the dragged sizes through the engine's measures only. */
function previewDrag(current: Drag, event: PointerEvent<HTMLDivElement>) {
  current.delta =
    ((current.axis === 'column' ? event.clientX : event.clientY) - current.start) / current.scale;
  for (const original of current.originals) {
    const size = String(clampedSize(original.size + current.delta) * current.scale);
    if (current.axis === 'column')
      applyJSpreadsheetColumnMeasure(current.worksheet, original.index, {
        attributeWidth: size,
        cssWidth: `${size}px`,
      });
    else
      applyJSpreadsheetRowMeasure(current.worksheet, original.index, {
        cssHeight: `${size}px`,
        attributeHeight: size,
      });
  }
}

/** The worksheet with the dragged lines at their new sizes. */
function resizedWorksheet(worksheet: Worksheet, current: Drag): Worksheet {
  const next = structuredClone(worksheet);
  for (const original of current.originals) {
    const size = clampedSize(original.size + current.delta);
    if (current.axis === 'column') {
      next.columns = [...(next.columns ?? [])];
      while (next.columns.length <= original.index) next.columns.push({});
      next.columns[original.index] = { ...next.columns[original.index], width: size };
    } else (next.rows ??= {})[String(original.index)] = { height: size };
  }
  return next;
}

/** Ends drags when the window blurs, resizing is turned off, the zoom changes or it unmounts. */
function useDragCancellation(
  options: Options,
  { active, drag }: { active: boolean; drag: RefObject<Drag | null> },
  cancel: (notify?: boolean) => void,
) {
  const cancelDrag = useEffectEvent(cancel);
  const { subscribeBlur, canResize, zoom } = options;
  useEffect(() => {
    if (!active) return;
    const callback = () => cancelDrag();
    if (subscribeBlur) return subscribeBlur(callback);
    window.addEventListener('blur', callback);
    return () => window.removeEventListener('blur', callback);
  }, [active, subscribeBlur]);
  useLayoutEffect(() => {
    // A drag ends as soon as resizing is turned off.
    if (!canResize) cancelDrag();
  }, [canResize]);
  useLayoutEffect(() => {
    if (drag.current && drag.current.scale !== gridZoomScale(zoom)) cancelDrag();
  }, [zoom, drag]);
  useEffect(() => () => cancelDrag(false), []);
}

/** Resize only the engine's measure channels while dragging, then publish one immutable change. */
export function useHeaderResize(options: Options) {
  const latest = useRef(options),
    drag = useRef<Drag | null>(null);
  const [active, setActive] = useState(false);
  useLayoutEffect(() => {
    latest.current = options;
  });
  function restore(current: Drag) {
    if (latest.current.getWorksheet() === current.worksheet) restoreMeasures(current);
  }
  function cancel(notify = true) {
    const current = drag.current;
    if (!current) return;
    restore(current);
    drag.current = null;
    releasePointer(latest.current.root.current, current.pointerId);
    if (notify) setActive(false);
    latest.current.onGeometryChange();
  }
  useDragCancellation(options, { active, drag }, cancel);
  function start(event: PointerEvent<HTMLDivElement>): boolean {
    const next = beginDrag(event, latest.current);
    if (!next) return false;
    const root = requireDefined(latest.current.root.current);
    drag.current = next;
    root.focus({ preventScroll: true });
    root.setPointerCapture(event.pointerId);
    setActive(true);
    return true;
  }
  function move(event: PointerEvent<HTMLDivElement>): boolean {
    const current = drag.current;
    if (!current) return false;
    event.preventDefault();
    event.stopPropagation();
    if (latest.current.getWorksheet() !== current.worksheet) {
      cancel();
      return true;
    }
    previewDrag(current, event);
    latest.current.onGeometryChange();
    return true;
  }
  function finish(event: PointerEvent<HTMLDivElement>): boolean {
    const current = drag.current;
    if (!current) return false;
    event.preventDefault();
    event.stopPropagation();
    const unchanged = Math.abs(current.delta) < 0.01;
    restore(current);
    drag.current = null;
    setActive(false);
    releasePointer(latest.current.root.current, current.pointerId);
    latest.current.onGeometryChange();
    const options = latest.current;
    if (unchanged || !options.canResize || options.getWorksheet() !== current.worksheet)
      return true;
    options.onChange(resizedWorksheet(options.worksheet, current));
    return true;
  }
  return { start, move, finish, cancel, active };
}
