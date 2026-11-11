'use client';

import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type PointerEvent,
  type RefObject,
} from 'react';
import type { Worksheet } from '../model/workbook.js';
import { DEFAULT_COLUMN_WIDTH, DEFAULT_ROW_HEIGHT, gridZoomScale } from './grid-geometry.js';
import {
  applyJSpreadsheetColumnMeasure,
  applyJSpreadsheetRowMeasure,
  readJSpreadsheetColumnMeasure,
  readJSpreadsheetRowMeasureSnapshot,
  type JSpreadsheetWorksheetInstance,
  type JSpreadsheetColumnMeasureSnapshot,
  type JSpreadsheetRowMeasureSnapshot,
} from './jspreadsheet-host.js';

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
export function deliveredHeader(event: Pick<PointerEvent, 'target'>, root: HTMLElement | null) {
  let element = event.target instanceof HTMLElement ? event.target : null;
  if (!root || !element || !root.contains(element)) return null;
  while (element && element !== root) {
    const x = element.hasAttribute('data-x') ? Number(element.dataset.x) : null,
      y = element.hasAttribute('data-y') ? Number(element.dataset.y) : null;
    if (element.tagName === 'TD') {
      if (x !== null && x >= 0 && (y === null || y < 0))
        return { axis: 'column' as const, index: x, element };
      if (y !== null && y >= 0 && (x === null || x < 0))
        return { axis: 'row' as const, index: y, element };
    }
    element = element.parentElement;
  }
  return null;
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
    if (latest.current.getWorksheet() !== current.worksheet) return;
    for (const original of current.originals) {
      if (original.column)
        applyJSpreadsheetColumnMeasure(
          current.worksheet,
          original.index,
          original.column.attributeWidth,
          original.column.cssWidth,
        );
      if (original.row)
        applyJSpreadsheetRowMeasure(
          current.worksheet,
          original.index,
          original.row.cssHeight,
          original.row.attributeHeight,
        );
    }
  }
  function cancel(notify = true) {
    const current = drag.current;
    if (!current) return;
    restore(current);
    drag.current = null;
    const root = latest.current.root.current;
    if (root?.hasPointerCapture(current.pointerId)) root.releasePointerCapture(current.pointerId);
    if (notify) setActive(false);
    latest.current.onGeometryChange();
  }
  useEffect(() => {
    if (!active) return;
    const callback = () => cancel();
    if (options.subscribeBlur) return options.subscribeBlur(callback);
    window.addEventListener('blur', callback);
    return () => window.removeEventListener('blur', callback);
  }, [active, options.subscribeBlur]);
  useLayoutEffect(() => {
    if (!options.canResize) cancel();
  }, [options.canResize]);
  useLayoutEffect(() => {
    if (drag.current && drag.current.scale !== gridZoomScale(options.zoom)) cancel();
  }, [options.zoom]);
  useEffect(() => () => cancel(false), []);
  function start(event: PointerEvent<HTMLDivElement>): boolean {
    const current = latest.current,
      header = deliveredHeader(event, current.root.current),
      worksheet = current.getWorksheet();
    if (!header || !worksheet || event.button !== 0 || !current.canResize) return false;
    const rect = header.element.getBoundingClientRect(),
      edge = header.axis === 'column' ? rect.right - event.clientX : rect.bottom - event.clientY;
    if (edge < 0 || edge > 6) return false;
    event.preventDefault();
    event.stopPropagation();
    const root = current.root.current!,
      scale = gridZoomScale(current.zoom);
    const selected =
      header.axis === 'column'
        ? worksheet.getSelectedColumns?.(true)
        : worksheet.getSelectedRows?.(true);
    const indices = Array.isArray(selected)
      ? selected.map(Number).filter((index) => Number.isInteger(index) && index >= 0)
      : [];
    const affected = indices.includes(header.index) ? indices : [header.index];
    const originals = affected.map((index) =>
      header.axis === 'column'
        ? {
            index,
            size: current.worksheet.columns?.[index]?.width ?? DEFAULT_COLUMN_WIDTH,
            column: readJSpreadsheetColumnMeasure(worksheet, index),
          }
        : {
            index,
            size: current.worksheet.rows?.[String(index)]?.height ?? DEFAULT_ROW_HEIGHT,
            row: readJSpreadsheetRowMeasureSnapshot(worksheet, index),
          },
    );
    drag.current = {
      axis: header.axis,
      pointerId: event.pointerId,
      start: header.axis === 'column' ? event.clientX : event.clientY,
      scale,
      delta: 0,
      worksheet,
      originals,
    };
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
    current.delta =
      ((current.axis === 'column' ? event.clientX : event.clientY) - current.start) / current.scale;
    for (const original of current.originals) {
      const size = String(
        Math.max(1, Math.min(2000, original.size + current.delta)) * current.scale,
      );
      if (current.axis === 'column')
        applyJSpreadsheetColumnMeasure(current.worksheet, original.index, size, `${size}px`);
      else applyJSpreadsheetRowMeasure(current.worksheet, original.index, `${size}px`, size);
    }
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
    const root = latest.current.root.current;
    if (root?.hasPointerCapture(current.pointerId)) root.releasePointerCapture(current.pointerId);
    latest.current.onGeometryChange();
    if (
      unchanged ||
      !latest.current.canResize ||
      latest.current.getWorksheet() !== current.worksheet
    )
      return true;
    const next = structuredClone(latest.current.worksheet);
    for (const original of current.originals) {
      const size = Math.max(1, Math.min(2000, original.size + current.delta));
      if (current.axis === 'column') {
        next.columns = [...(next.columns ?? [])];
        while (next.columns.length <= original.index) next.columns.push({});
        next.columns[original.index] = { ...next.columns[original.index], width: size };
      } else (next.rows ??= {})[String(original.index)] = { height: size };
    }
    latest.current.onChange(next);
    return true;
  }
  return { start, move, finish, cancel, active };
}
