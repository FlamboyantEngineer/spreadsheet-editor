import { jspreadsheet } from '@jspreadsheet-ce/react';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { invalidateGridRowHeights } from './row-window.js';

export interface JSpreadsheetColumnMeasureSnapshot {
  readonly attributeWidth: string;
  readonly cssWidth: string;
}

type JSpreadsheetWorksheetMeasureSource = Pick<JSpreadsheetWorksheetInstance, 'cols' | 'rows'>;

/**
 * Read the two column-width channels from jspreadsheet's own `<col>` element.
 * Callers receive strings only; vendor DOM never escapes this boundary.
 */
export function readJSpreadsheetColumnMeasure(
  worksheet: JSpreadsheetWorksheetMeasureSource,
  columnIndex: number,
): JSpreadsheetColumnMeasureSnapshot {
  const columnElement = worksheet.cols?.[columnIndex]?.colElement;
  return {
    attributeWidth: columnElement?.getAttribute('width') ?? '',
    cssWidth: columnElement?.style.width ?? '',
  };
}

/**
 * Keep jspreadsheet's attribute storage and fractional CSS render channel in
 * sync on the exact vendor-created `<col>` element. Passing `cssWidth` restores
 * a pair read by `readJSpreadsheetColumnMeasure`, including an empty CSS width.
 */
export function applyJSpreadsheetColumnMeasure(
  worksheet: JSpreadsheetWorksheetMeasureSource,
  columnIndex: number,
  measure: { attributeWidth: string; cssWidth?: string },
): boolean {
  const normalizedWidth = measure.attributeWidth,
    cssWidth = measure.cssWidth ?? `${normalizedWidth}px`;
  const columnElement = worksheet.cols?.[columnIndex]?.colElement;
  if (!(columnElement instanceof HTMLElement)) return false;
  // Unchanged measures are not rewritten: each write is a mutation the spill observer handles.
  if (columnElement.getAttribute('width') !== normalizedWidth)
    columnElement.setAttribute('width', normalizedWidth);
  if (columnElement.style.width !== cssWidth) columnElement.style.width = cssWidth;
  return true;
}

export interface JSpreadsheetRowMeasureSnapshot {
  readonly attributeHeight: string;
  readonly cssHeight: string;
}

/** Read both `<tr>` height channels; an absent attribute reads as empty. */
export function readJSpreadsheetRowMeasureSnapshot(
  worksheet: JSpreadsheetWorksheetMeasureSource,
  rowIndex: number,
): JSpreadsheetRowMeasureSnapshot {
  const rowElement = worksheet.rows?.[rowIndex]?.element;
  return {
    attributeHeight: rowElement?.getAttribute('height') ?? '',
    cssHeight: rowElement?.style.height ?? '',
  };
}

/**
 * Keep jspreadsheet's attribute storage and fractional CSS render channel in
 * sync on the exact vendor-created `<tr>` element. Passing `attributeHeight`
 * restores a pair read by `readJSpreadsheetRowMeasureSnapshot`; the browser and
 * jspreadsheet's `setHeight` both treat an empty attribute as absent.
 */
export function applyJSpreadsheetRowMeasure(
  worksheet: JSpreadsheetWorksheetMeasureSource,
  rowIndex: number,
  measure: { cssHeight: string; attributeHeight?: string },
): boolean {
  const { cssHeight } = measure,
    attributeHeight = measure.attributeHeight ?? cssHeight;
  const rowElement = worksheet.rows?.[rowIndex]?.element;
  if (!(rowElement instanceof HTMLElement)) return false;
  if (rowElement.style.height !== cssHeight) {
    rowElement.style.height = cssHeight;
    invalidateGridRowHeights(worksheet);
  }
  if (rowElement.getAttribute('height') !== attributeHeight)
    rowElement.setAttribute('height', attributeHeight);
  return true;
}

/**
 * Detaches the worksheet jspreadsheet tracks as `jspreadsheet.current`. The
 * vendor's document keydown and paste listeners act on that worksheet whatever
 * the event target, so keys typed in an app dialog would edit, clear, or delete
 * grid content. Returns null when no worksheet was tracked, otherwise a function
 * that reattaches the worksheet unless the vendor tracks another one by then or
 * the worksheet was destroyed.
 */
export function releaseJSpreadsheetKeyboardOwner(
  expectedOwner?: JSpreadsheetWorksheetInstance | null,
): (() => void) | null {
  const vendor = jspreadsheet as unknown as { current?: { element?: unknown } | null };
  const owner = vendor.current;
  if (!owner) return null;
  // Late teardown of one mounted grid must not release a newer sheet or editor.
  if (expectedOwner !== undefined && owner !== expectedOwner) return null;
  vendor.current = null;
  return () => {
    const element = owner.element;
    if (vendor.current == null && (!(element instanceof Element) || element.isConnected))
      vendor.current = owner;
  };
}

/** Restore keyboard ownership to the exact mounted worksheet requested by React chrome. */
export function focusJSpreadsheetWorksheet(worksheet: JSpreadsheetWorksheetInstance): void {
  const element = worksheet.element ?? worksheet.root ?? worksheet.table;
  if (element instanceof Element && !element.isConnected) return;
  const vendor = jspreadsheet as unknown as { current?: JSpreadsheetWorksheetInstance | null };
  vendor.current = worksheet;
}
