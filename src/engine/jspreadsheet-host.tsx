'use client';

/*
 * Allowlisted jspreadsheet engine host (#1677, #1687). The official
 * @jspreadsheet-ce/react Spreadsheet wrapper owns the engine root and instance;
 * this host preserves data-jspreadsheet-host on that root and bridges typed
 * callbacks through createBridgedConfig. React chrome consumes request events
 * outside the engine-owned DOM. See REACT_UI_ENGINE_HOST_ALLOWLIST.
 *
 * The wrapper has no teardown, so unmount calls onBeforeDestroy and
 * teardownJspreadsheetInstance. setData updates the worksheet in place when
 * available; otherwise a key change remounts the wrapper with the new dataset.
 */

import {
  createElement,
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { ComponentType, CSSProperties, ReactElement, Ref } from 'react';
// The vendor-provided React wrapper for jspreadsheet-ce. `<Spreadsheet>` owns
// the engine root <div> and mounts `jspreadsheet(div, options)` into it,
// exposing the live instance on the `ref` it receives. The package also
// re-exports the raw `jspreadsheet` namespace, which we use only for engine
// teardown (`jspreadsheet.destroy`) — the host never constructs the engine
// itself, so it does NOT import `jspreadsheet-ce` directly.
import { Spreadsheet, jspreadsheet } from '@jspreadsheet-ce/react';
import type JSpreadsheetCore from 'jspreadsheet-ce';
import {
  resolveSpreadsheetRowSpillWidths,
  type SpreadsheetSpillCell,
} from './spreadsheet-cell-spill.js';

// --- Public types ------------------------------------------------------------

/**
 * Coordinates of a single cell in the active worksheet. `x` is the column
 * index and `y` is the row index. This mirrors the jspreadsheet convention.
 */
export interface JSpreadsheetCellCoords {
  readonly x: number;
  readonly y: number;
}

/**
 * Selection range in the active worksheet, normalized so that `x1 <= x2` and
 * `y1 <= y2`.
 */
export interface JSpreadsheetSelectionRange {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

/**
 * Geometry of an anchor element jspreadsheet supplied for a request event.
 * Phase 9B uses this rect to position the React overlay it renders through a
 * portal. The rect is in viewport coordinates (the same convention as
 * `Element.getBoundingClientRect()`).
 */
export interface JSpreadsheetAnchorRect {
  readonly top: number;
  readonly left: number;
  readonly right: number;
  readonly bottom: number;
  readonly width: number;
  readonly height: number;
}

/**
 * Generic value a single cell may hold. The grid is text-first; numbers and
 * booleans surface as their stringified form in most code paths.
 */
export type JSpreadsheetCellValue = JSpreadsheetCore.CellValue | null;

/**
 * Vendor callback shape. jspreadsheet-ce callback arity and argument
 * types vary by callback and plugin, so the boundary preserves the caller's
 * callback while normalizing only the dispatcher values the editor owns.
 */
type JSpreadsheetVendorCallback = {
  bivarianceHack(...args: unknown[]): unknown;
}['bivarianceHack'];

type JSpreadsheetContextMenuHost = Partial<HTMLDivElement> & {
  contextmenu?: {
    close?: (focus?: boolean) => void;
  };
};

type JSpreadsheetStyleMap = Record<string, string | string[]>;

/**
 * Configuration passed straight through to the vendor `<Spreadsheet>` wrapper.
 *
 * This is a the editor adapter type rather than the exact upstream
 * `SpreadsheetOptions`: the runtime still passes data-only configs,
 * workbook configs, and plugin-specific callback fields through this boundary.
 * The strongly typed surface here is limited to the callback names this host
 * intercepts; everything else stays a pass-through engine/plugin option.
 */
export type JSpreadsheetConfig = Record<string, unknown> & {
  contextMenu?: JSpreadsheetVendorCallback;
  data?: ReadonlyArray<ReadonlyArray<JSpreadsheetCellValue>>;
  onchange?: JSpreadsheetVendorCallback;
  oncomments?: JSpreadsheetVendorCallback;
  oncreateworksheet?: JSpreadsheetVendorCallback;
  oneditionend?: JSpreadsheetVendorCallback;
  oneditionstart?: JSpreadsheetVendorCallback;
  onselection?: JSpreadsheetVendorCallback;
};

/**
 * Worksheet handle as consumed by the editor. The upstream declarations are
 * intentionally narrower than the live wrapper shape in a few places
 * (`setValue`, `setData`, test stubs), so this adapter captures the methods
 * this host and the runtime actually call.
 */
export type JSpreadsheetWorksheetInstance = Partial<
  Omit<
    JSpreadsheetCore.WorksheetInstance,
    | 'closeEditor'
    | 'deleteColumn'
    | 'deleteRow'
    | 'getCellFromCoords'
    | 'getComments'
    | 'getConfig'
    | 'getData'
    | 'getHeight'
    | 'getMerge'
    | 'getValue'
    | 'getValueFromCoords'
    | 'getWidth'
    | 'insertColumn'
    | 'insertRow'
    | 'openWorksheet'
    | 'getSelectedColumns'
    | 'getSelectedRows'
    | 'getSelection'
    | 'parent'
    | 'removeMerge'
    | 'setComments'
    | 'setData'
    | 'setHeight'
    | 'setMerge'
    | 'setReadOnly'
    | 'setStyle'
    | 'setValue'
    | 'setValueFromCoords'
    | 'setWidth'
    | 'updateSelectionFromCoords'
  >
> & {
  closeEditor?: (cell?: HTMLTableCellElement, save?: boolean) => void;
  deleteColumn?: (columnIndex?: number, count?: number) => boolean | void;
  deleteRow?: (rowIndex?: number, count?: number) => boolean | void;
  destroy?: () => void;
  element?: HTMLElement;
  el?: HTMLElement;
  getCellFromCoords?: (x: number, y: number) => HTMLTableCellElement | HTMLElement | null;
  getComments?: (cell?: string) => Record<string, string> | string;
  getConfig?: () => Record<string, unknown>;
  getData?: (
    highlighted?: boolean,
    processed?: boolean,
    delimiter?: string,
    asJson?: boolean,
  ) => JSpreadsheetCellValue[][];
  getHeight?: (rowIndex?: number) => number | string | string[] | null;
  getMerge?: (
    cellName?: string,
  ) => Record<string, [number, number]> | [number, number] | null | undefined;
  getSelectedColumns?: (visibleOnly?: boolean) => number[];
  getSelectedRows?: (visibleOnly?: boolean) => number[];
  getSelection?: () => unknown;
  getValue?: (
    x: number,
    y: number,
    processedValue?: boolean,
  ) => JSpreadsheetCellValue | HTMLElement;
  getValueFromCoords?: (
    x: number,
    y: number,
    processedValue?: boolean,
  ) => JSpreadsheetCellValue | HTMLElement;
  getWidth?: (columnIndex?: number) => number | string | Array<number | string> | null;
  insertColumn?: (count?: number, columnIndex?: number, insertBefore?: boolean) => boolean | void;
  insertRow?: (count?: number, rowIndex?: number, insertBefore?: boolean) => boolean | void;
  openWorksheet?: (...args: unknown[]) => unknown;
  parent?: JSpreadsheetSpreadsheetInstance | null;
  refresh?: () => void;
  removeMerge?: (address: string) => void;
  root?: HTMLElement;
  setComments?: (comments: Record<string, string>) => void;
  setData?: (data?: ReadonlyArray<ReadonlyArray<JSpreadsheetCellValue>>) => void;
  setHeight?: (rowIndex: number, height: number) => void;
  setMerge?: (cellName?: string, colspan?: number, rowspan?: number) => null | void;
  setReadOnly?: (cell: string | HTMLTableCellElement, state: boolean) => void;
  setStyle?: {
    (style: JSpreadsheetStyleMap, property?: null, value?: null, force?: boolean): void;
    (cell: string, property: string, value: string, force?: boolean): void;
  };
  setValue?: (
    cell:
      | string
      | HTMLTableCellElement
      | [number, number]
      | Array<
          | string
          | HTMLTableCellElement
          | [number, number]
          | { x: number; y: number; value?: JSpreadsheetCellValue }
        >,
    value?: JSpreadsheetCellValue,
    force?: boolean,
  ) => void;
  setValueFromCoords?: (
    x: number,
    y: number,
    value: JSpreadsheetCellValue,
    force?: boolean,
  ) => void;
  setWidth?: (columnIndex: number | number[], width: number | number[]) => void;
  updateSelectionFromCoords?: (x1: number, y1: number, x2: number, y2: number) => void;
};

/**
 * Spreadsheet/root handle as consumed by the editor. The vendor wrapper can
 * expose either the root spreadsheet instance or worksheet instances depending
 * on build/test shape, so consumers use `resolveSpreadsheetWorksheets` before
 * making worksheet-level calls.
 */
export type JSpreadsheetSpreadsheetInstance = Partial<
  Omit<JSpreadsheetCore.SpreadsheetInstance, 'contextMenu' | 'worksheets'>
> & {
  contextMenu?: JSpreadsheetContextMenuHost | null;
  destroy?: () => void;
  element?: HTMLElement;
  el?: HTMLElement;
  root?: HTMLElement;
  worksheets?: JSpreadsheetWorksheetInstance[];
};
export type JSpreadsheetInstance =
  JSpreadsheetSpreadsheetInstance | JSpreadsheetWorksheetInstance | JSpreadsheetWorksheetInstance[];

/**
 * Imperative API exposed via `ref` for surrounding chrome / commands. These
 * methods are the only sanctioned way for Phase 9B to drive the engine from
 * outside the host. Anything else should be expressed as new props/events.
 */
export interface JSpreadsheetHostHandle {
  /** Write `value` to the cell at `(row, col)` using the active worksheet. */
  setCellValue(row: number, col: number, value: JSpreadsheetCellValue): void;
  /** Read the current value at `(row, col)` from the active worksheet. */
  getCellValue(row: number, col: number): JSpreadsheetCellValue;
  /**
   * Replace the active worksheet dataset without re-instantiating the engine.
   * Falls back to a full re-mount if the engine cannot mutate the dataset in
   * place.
   */
  setData(data: ReadonlyArray<ReadonlyArray<JSpreadsheetCellValue>>): void;
  /** Return the current selection range, or `null` if nothing is selected. */
  getCurrentSelection(): JSpreadsheetSelectionRange | null;
  /** Ask jspreadsheet to refresh the rendered grid. */
  refresh(): void;
  /** Return the raw jspreadsheet instance for advanced commands. */
  getInstance(): JSpreadsheetInstance | null;
}

export interface JSpreadsheetColumnMeasureSnapshot {
  readonly attributeWidth: string;
  readonly cssWidth: string;
}

export type JSpreadsheetWorksheetMeasureSource = Pick<
  JSpreadsheetWorksheetInstance,
  'cols' | 'rows'
>;

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
  normalizedWidth: string,
  cssWidth = `${normalizedWidth}px`,
): boolean {
  const columnElement = worksheet.cols?.[columnIndex]?.colElement;
  if (!(columnElement instanceof HTMLElement)) return false;
  columnElement.setAttribute('width', normalizedWidth);
  columnElement.style.width = cssWidth;
  return true;
}

/** Read the rendered height from jspreadsheet's own `<tr>` element. */
export function readJSpreadsheetRowMeasure(
  worksheet: JSpreadsheetWorksheetMeasureSource,
  rowIndex: number,
): string {
  return worksheet.rows?.[rowIndex]?.element?.style.height ?? '';
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
  cssHeight: string,
  attributeHeight = cssHeight,
): boolean {
  const rowElement = worksheet.rows?.[rowIndex]?.element;
  if (!(rowElement instanceof HTMLElement)) return false;
  rowElement.style.height = cssHeight;
  rowElement.setAttribute('height', attributeHeight);
  return true;
}

/** Axis, index, and pointer origin of jspreadsheet's in-progress header-border drag. */
export interface JSpreadsheetResizeDescriptor {
  readonly axis: 'column' | 'row';
  readonly index: number;
  readonly pointerOrigin: number;
  /** On-screen size of the dragged header when the drag began. */
  readonly renderedSize: number;
}

// jspreadsheet keeps its header-drag state on the worksheet object it tracks as
// `jspreadsheet.current`; these fields are vendor internals read only here.
type JSpreadsheetResizeSource = {
  headers?: ArrayLike<Element | undefined>;
  records?: ArrayLike<ArrayLike<{ element?: unknown } | undefined> | undefined>;
  resizing?: {
    column?: unknown;
    element?: unknown;
    height?: unknown;
    mousePosition?: unknown;
    row?: unknown;
    width?: unknown;
  } | null;
  rows?: ArrayLike<{ element?: unknown } | undefined>;
};

/** Reads jspreadsheet's header-border drag state, or null when no drag is active. */
export function readJSpreadsheetResizeDescriptor(
  worksheet: object,
): JSpreadsheetResizeDescriptor | null {
  const resizing = (worksheet as JSpreadsheetResizeSource).resizing;
  const pointerOrigin = Number(resizing?.mousePosition);
  if (!resizing || !Number.isFinite(pointerOrigin)) return null;
  const axis = resizing.column != null ? 'column' : 'row';
  const index = Number(axis === 'column' ? resizing.column : resizing.row);
  const renderedSize = Number(axis === 'column' ? resizing.width : resizing.height);
  return Number.isInteger(index) && index >= 0
    ? { axis, index, pointerOrigin, renderedSize: Number.isFinite(renderedSize) ? renderedSize : 0 }
    : null;
}

/**
 * Ends jspreadsheet's header-border drag without its own mouseup commit: clears
 * the vendor drag state and the `resizing` classes it put on the dragged header,
 * column cells, or row.
 */
export function clearJSpreadsheetResizeDescriptor(worksheet: object): void {
  const source = worksheet as JSpreadsheetResizeSource;
  const descriptor = readJSpreadsheetResizeDescriptor(source);
  const rowElement = source.resizing?.element;
  source.resizing = null;
  if (!descriptor) return;
  const dragged: unknown[] =
    descriptor.axis === 'column'
      ? [
          source.headers?.[descriptor.index],
          ...Array.from(source.records ?? [], (row) => row?.[descriptor.index]?.element),
        ]
      : [
          rowElement,
          (source.rows?.[descriptor.index]?.element as Element | undefined)?.children?.[0],
        ];
  for (const element of dragged) {
    if (element instanceof Element) element.classList.remove('resizing');
  }
}

/**
 * Runs an engine write with jspreadsheet's native `confirm()` answered by
 * `answer`, and reports whether the engine asked. The engine asks before an
 * insert, delete or move that cuts through merged cells (then drops every merge
 * on the sheet) and returns without changing anything when told no, so a write
 * run with `false` probes whether the question arises (PT-MD-2.1).
 */
export function runWithJSpreadsheetConfirmAnswer<T>(
  answer: boolean,
  write: () => T,
): { result: T; asked: boolean } {
  const host = globalThis as { confirm?: ((message?: string) => boolean) | undefined };
  const nativeConfirm = host.confirm;
  let asked = false;
  host.confirm = () => {
    asked = true;
    return answer;
  };
  try {
    const result = write();
    return { result, asked };
  } finally {
    host.confirm = nativeConfirm;
  }
}

/**
 * The whole rows or columns selected through their headers on the worksheet
 * jspreadsheet's keyboard handler acts on, or null. Delete and Backspace remove
 * those lines, behind a native `confirm()` the runtime replaces.
 */
export function readJSpreadsheetHeaderSelection(): {
  axis: 'row' | 'column';
  worksheet: object;
} | null {
  const owner = (
    jspreadsheet as unknown as {
      current?: {
        selectedCell?: unknown;
        selectedRow?: unknown;
        selectedHeader?: unknown;
        edition?: unknown;
        options?: { editable?: unknown; allowDeleteRow?: unknown; allowDeleteColumn?: unknown };
      } | null;
    }
  ).current;
  // The engine's own Delete handling runs under exactly these conditions.
  if (!owner || owner.edition || !owner.selectedCell || owner.options?.editable === false)
    return null;
  if (owner.selectedRow != null)
    return owner.options?.allowDeleteRow === false ? null : { axis: 'row', worksheet: owner };
  if (owner.selectedHeader != null)
    return owner.options?.allowDeleteColumn === false ? null : { axis: 'column', worksheet: owner };
  return null;
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

/**
 * Public prop contract of `JSpreadsheetHost`. Every prop below is OPTIONAL
 * except `config` so Phase 9B can adopt the host incrementally.
 */
export interface JSpreadsheetHostProps {
  /**
   * jspreadsheet configuration object (data, columns, worksheets, options).
   * Passed directly to `jspreadsheet(root, config)`. The host clones nothing;
   * the caller owns immutability.
   */
  config: JSpreadsheetConfig;

  /**
   * Optional className on the engine root `<div>`. Use sparingly — the host
   * exists to OWN engine DOM, not be styled by chrome.
   */
  className?: string | undefined;

  /** Optional inline style on the engine root `<div>`. */
  style?: CSSProperties | undefined;

  /** Optional id on the engine root `<div>` (defaults to none). */
  id?: string | undefined;

  /**
   * Fired once after the jspreadsheet instance is created. The chrome (Phase
   * 9B) typically needs the instance for direct commands such as inserting
   * rows or applying formats.
   */
  onMounted?:
    ((instance: JSpreadsheetInstance, handle: JSpreadsheetHostHandle) => void) | undefined;

  /** Fired once just before the instance is destroyed (unmount or remount). */
  onBeforeDestroy?: ((instance: JSpreadsheetInstance) => void) | undefined;

  /**
   * Fired when a cell value is committed by the engine. The host bridges the
   * jspreadsheet `onchange(worksheet, cellEl, x, y, newValue, oldValue)`
   * callback into this typed shape.
   */
  onChange?:
    | ((
        row: number,
        col: number,
        oldValue: JSpreadsheetCellValue,
        newValue: JSpreadsheetCellValue,
      ) => void)
    | undefined;

  /** Fired when the selection range changes (mouse, keyboard, programmatic). */
  onSelectionChange?: ((coords: JSpreadsheetSelectionRange) => void) | undefined;

  /** Fired when the user begins editing a cell. */
  onCellEditStart?: ((row: number, col: number) => void) | undefined;

  /** Fired when the user commits or cancels a cell edit. */
  onCellEditEnd?:
    ((row: number, col: number, committedValue: JSpreadsheetCellValue) => void) | undefined;

  /**
   * REQUEST event: the engine wants formula autocomplete to be displayed.
   * Phase 9B will render the autocomplete dropdown as a React portal outside
   * the engine DOM, anchored at `anchorRect`.
   */
  onFormulaAutocompleteRequest?:
    ((query: string, anchorRect: JSpreadsheetAnchorRect) => void) | undefined;

  /**
   * REQUEST event: the user opened the cell context menu. Phase 9B will
   * render the React context menu portal anchored at `anchorRect`.
   */
  onContextMenuRequest?:
    ((coords: JSpreadsheetCellCoords, anchorRect: JSpreadsheetAnchorRect) => void) | undefined;

  /**
   * REQUEST event: the user requested the comment popover for a cell. Phase
   * 9B will render the React comment popover portal anchored at `anchorRect`.
   */
  onCellCommentRequest?:
    ((coords: JSpreadsheetCellCoords, anchorRect: JSpreadsheetAnchorRect) => void) | undefined;
}

// --- Engine config bridge ----------------------------------------------------

/**
 * Options forwarded to `createBridgedConfig`. It wires the jspreadsheet-level
 * callbacks (`onchange`, `onselection`, `oneditionstart`, `oneditionend`,
 * `contextMenu`, `oncomments`) to the host's typed dispatchers.
 */
export interface JspreadsheetBridgeOptions {
  readonly config: JSpreadsheetConfig;
  readonly dispatchChange: (
    row: number,
    col: number,
    oldValue: JSpreadsheetCellValue,
    newValue: JSpreadsheetCellValue,
  ) => void;
  readonly dispatchSelectionChange: (range: JSpreadsheetSelectionRange) => void;
  readonly dispatchEditStart: (row: number, col: number) => void;
  readonly dispatchEditEnd: (
    row: number,
    col: number,
    committedValue: JSpreadsheetCellValue,
  ) => void;
  readonly dispatchFormulaAutocompleteRequest: (
    query: string,
    rect: JSpreadsheetAnchorRect,
  ) => void;
  readonly dispatchContextMenuRequest: (
    coords: JSpreadsheetCellCoords,
    rect: JSpreadsheetAnchorRect,
  ) => void;
  readonly dispatchCellCommentRequest: (
    coords: JSpreadsheetCellCoords,
    rect: JSpreadsheetAnchorRect,
  ) => void;
}

function rectFromElement(target: unknown): JSpreadsheetAnchorRect {
  if (target && typeof target === 'object' && 'getBoundingClientRect' in target) {
    const fn = (target as { getBoundingClientRect?: () => DOMRect }).getBoundingClientRect;
    if (typeof fn === 'function') {
      const rect = fn.call(target);
      return {
        top: rect.top,
        left: rect.left,
        right: rect.right,
        bottom: rect.bottom,
        width: rect.width,
        height: rect.height,
      };
    }
  }
  return { top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 };
}

function safeCoord(value: unknown): number {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function readActiveWorksheet(
  instance: JSpreadsheetInstance | null,
): JSpreadsheetWorksheetInstance | null {
  if (!instance) return null;
  if (Array.isArray(instance)) {
    return instance[0] || null;
  }
  if ('worksheets' in instance && Array.isArray(instance.worksheets)) {
    const worksheets = instance.worksheets;
    try {
      const index =
        typeof instance.getWorksheetActive === 'function' ? instance.getWorksheetActive() : 0;
      return worksheets[index] || worksheets[0] || null;
    } catch {
      // ignore
    }
  }
  return instance as JSpreadsheetWorksheetInstance;
}

/**
 * Resolve the engine root `<div>` from a jspreadsheet instance. The vendor
 * wrapper does not hand us the element it created, so we derive it from the
 * instance: jspreadsheet-ce returns `WorksheetInstance[]`, and each worksheet's
 * `parent` is the `SpreadsheetInstance` whose `el`/`element` is the root div.
 * Single-instance and `root`-bearing shapes (e.g. the test stub) are covered as
 * fallbacks. Used to stamp the `data-jspreadsheet-host` marker and to target
 * teardown at the correct element.
 */
function resolveEngineRootElement(instance: JSpreadsheetInstance | null): HTMLElement | null {
  if (!instance) return null;
  const first = Array.isArray(instance) ? instance[0] : instance;
  if (!first || typeof first !== 'object') return null;
  const parent = 'parent' in first ? first.parent : null;
  if (parent && (parent.el || parent.element)) {
    return parent.el ?? parent.element ?? null;
  }
  if ('el' in first && first.el) {
    return first.el;
  }
  if ('element' in first && first.element) {
    return first.element;
  }
  if ('root' in first && first.root instanceof HTMLElement) {
    return first.root;
  }
  return null;
}

/** INV-049 / PT-RM-9.4: report vendor callback failures while allowing typed dispatchers to run. */
function reportVendorCallbackError(callbackName: string, error: unknown): void {
  console.error(`[JSpreadsheetHost] vendor ${callbackName} callback threw`, error);
}

/**
 * Build the merged jspreadsheet config: the caller's `config` plus the engine
 * callback shims that translate into the typed dispatcher functions. Existing
 * callback handlers in `config` are preserved and run BEFORE the host
 * dispatchers so caller-supplied behavior wins. This bridged config is handed
 * to the vendor `<Spreadsheet>` wrapper as props — it is the host's single
 * engine-config builder.
 */
export function createBridgedConfig(options: JspreadsheetBridgeOptions): Record<string, unknown> {
  const { config } = options;

  const userOnChange = config.onchange;
  const userOnSelection = config.onselection;
  const userOnEditionStart = config.oneditionstart;
  const userOnEditionEnd = config.oneditionend;
  const userContextMenu = config.contextMenu;
  const userOnComments = config.oncomments;
  const userOnCreateWorksheet = config.oncreateworksheet;

  const mergedConfig: Record<string, unknown> = {
    ...config,
    onchange: (
      worksheet: JSpreadsheetWorksheetInstance,
      cellEl: HTMLTableCellElement,
      x: unknown,
      y: unknown,
      newValue: unknown,
      oldValue: unknown,
    ) => {
      if (typeof userOnChange === 'function') {
        try {
          (userOnChange as JSpreadsheetVendorCallback)(worksheet, cellEl, x, y, newValue, oldValue);
        } catch (error) {
          reportVendorCallbackError('onchange', error);
        }
      }
      options.dispatchChange(
        safeCoord(y),
        safeCoord(x),
        oldValue as JSpreadsheetCellValue,
        newValue as JSpreadsheetCellValue,
      );
    },
    onselection: (instance: unknown, x1: unknown, y1: unknown, x2: unknown, y2: unknown) => {
      if (typeof userOnSelection === 'function') {
        try {
          (userOnSelection as JSpreadsheetVendorCallback)(instance, x1, y1, x2, y2);
        } catch (error) {
          reportVendorCallbackError('onselection', error);
        }
      }
      options.dispatchSelectionChange({
        x1: safeCoord(x1),
        y1: safeCoord(y1),
        x2: safeCoord(x2),
        y2: safeCoord(y2),
      });
    },
    oneditionstart: (
      worksheet: JSpreadsheetWorksheetInstance,
      cellEl: HTMLTableCellElement,
      x: unknown,
      y: unknown,
    ) => {
      if (typeof userOnEditionStart === 'function') {
        try {
          (userOnEditionStart as JSpreadsheetVendorCallback)(worksheet, cellEl, x, y);
        } catch (error) {
          reportVendorCallbackError('oneditionstart', error);
        }
      }
      options.dispatchEditStart(safeCoord(y), safeCoord(x));
    },
    oneditionend: (
      worksheet: JSpreadsheetWorksheetInstance,
      cellEl: HTMLTableCellElement,
      x: unknown,
      y: unknown,
      value: unknown,
    ) => {
      if (typeof userOnEditionEnd === 'function') {
        try {
          (userOnEditionEnd as JSpreadsheetVendorCallback)(worksheet, cellEl, x, y, value);
        } catch (error) {
          reportVendorCallbackError('oneditionend', error);
        }
      }
      options.dispatchEditEnd(safeCoord(y), safeCoord(x), value as JSpreadsheetCellValue);
    },
    contextMenu: (
      instance: unknown,
      colIndex: unknown,
      rowIndex: unknown,
      event: unknown,
      items: unknown,
      role: unknown,
      x: unknown,
      y: unknown,
    ) => {
      const rect = rectFromElement(
        event && typeof event === 'object' && 'target' in event
          ? (event as { target: unknown }).target
          : null,
      );
      options.dispatchContextMenuRequest(
        { x: safeCoord(colIndex ?? x), y: safeCoord(rowIndex ?? y) },
        rect,
      );
      if (typeof userContextMenu === 'function') {
        try {
          return (userContextMenu as JSpreadsheetVendorCallback)(
            instance,
            colIndex,
            rowIndex,
            event,
            items,
            role,
            x,
            y,
          );
        } catch (error) {
          reportVendorCallbackError('contextMenu', error);
          return items;
        }
      }
      return items;
    },
    oncomments: (
      worksheet: JSpreadsheetWorksheetInstance,
      cellEl: HTMLTableCellElement,
      x: unknown,
      y: unknown,
    ) => {
      if (typeof userOnComments === 'function') {
        try {
          (userOnComments as JSpreadsheetVendorCallback)(worksheet, cellEl, x, y);
        } catch (error) {
          reportVendorCallbackError('oncomments', error);
        }
      }
      options.dispatchCellCommentRequest(
        { x: safeCoord(x), y: safeCoord(y) },
        rectFromElement(cellEl),
      );
    },
    oncreateworksheet: (...args: unknown[]) => {
      try {
        if (typeof userOnCreateWorksheet === 'function') {
          return (userOnCreateWorksheet as JSpreadsheetVendorCallback)(...args);
        }
      } catch (error) {
        reportVendorCallbackError('oncreateworksheet', error);
      }
    },
  };

  return mergedConfig;
}

/**
 * Destroy a jspreadsheet instance and clear the host children. The vendor
 * `<Spreadsheet>` wrapper performs no teardown, so this host calls
 * `jspreadsheet.destroy(root, true)` (the namespace is re-exported by
 * `@jspreadsheet-ce/react`) and then clears any residual engine DOM.
 */
function teardownJspreadsheetInstance(
  host: HTMLElement,
  instance: JSpreadsheetInstance | null,
): void {
  const jss = jspreadsheet as unknown as { destroy?: (root: unknown, recursive: boolean) => void };
  try {
    if (typeof jss.destroy === 'function') {
      jss.destroy(host, true);
    } else if (
      instance &&
      typeof (instance as unknown as { destroy?: () => void }).destroy === 'function'
    ) {
      (instance as unknown as { destroy: () => void }).destroy();
    }
  } catch {
    // jspreadsheet-ce destroy behavior differs across builds; fall through to
    // the manual child-clear below.
  }
  try {
    host.replaceChildren();
  } catch {
    while (host.firstChild) {
      host.removeChild(host.firstChild);
    }
  }
}

// --- Cell spill (SE-CR-2.2a) -------------------------------------------------

/**
 * How far a text cell may paint past its right edge, in whole pixels; a
 * generated rule per width sets the cell's clip margin. It is presentation
 * only: the vendor style map, and so the saved workbook, never sees it.
 */
const JSPREADSHEET_CELL_SPILL_ATTRIBUTE = 'data-spreadsheet-spill';

function listJSpreadsheetWorksheets(
  instance: JSpreadsheetInstance | null,
): JSpreadsheetWorksheetInstance[] {
  if (!instance) return [];
  if (Array.isArray(instance)) return instance;
  if ('worksheets' in instance && Array.isArray(instance.worksheets)) return instance.worksheets;
  return [instance as JSpreadsheetWorksheetInstance];
}

function readJSpreadsheetSpillColumnWidth(colElement: HTMLElement | undefined): number {
  if (!colElement || colElement.style.display === 'none') return 0;
  const width = Number.parseFloat(colElement.style.width || colElement.getAttribute('width') || '');
  return Number.isFinite(width) ? width : 0;
}

/**
 * Classes that leave an empty cell open to spill: a read-only grid's marker and
 * a comment, whose corner is a background image. The adapter stylesheet takes
 * the same cells out of the positioned layer.
 */
const JSPREADSHEET_SPILL_TRANSPARENT_CLASSES = new Set(['readonly', 'jss_comments']);

function describeJSpreadsheetSpillCell(
  cell: HTMLTableCellElement | undefined,
): SpreadsheetSpillCell {
  if (!cell || cell.style.display === 'none') {
    return { rendered: false, empty: true, decorated: false, spillable: false };
  }
  return {
    rendered: true,
    empty: !cell.firstChild,
    decorated: Array.from(cell.classList).some(
      (name) => !JSPREADSHEET_SPILL_TRANSPARENT_CLASSES.has(name),
    ),
    spillable:
      !cell.firstElementChild &&
      cell.colSpan === 1 &&
      cell.rowSpan === 1 &&
      !cell.classList.contains('editor') &&
      !cell.classList.contains('jss_freezed'),
  };
}

/**
 * One stylesheet rule per spill width, shared by a document's grids, sets the
 * clip margin: typed `attr()` could read the width from the attribute itself,
 * but only Chromium supports it.
 */
const jspreadsheetSpillRuleSheets = new WeakMap<
  Document,
  { sheet: CSSStyleSheet; widths: Set<number> }
>();

function ensureJSpreadsheetSpillRule(ownerDocument: Document, width: number): void {
  let rules = jspreadsheetSpillRuleSheets.get(ownerDocument);
  if (!rules) {
    const StyleSheet = ownerDocument.defaultView?.CSSStyleSheet;
    if (!StyleSheet || !Array.isArray(ownerDocument.adoptedStyleSheets)) return;
    rules = { sheet: new StyleSheet(), widths: new Set<number>() };
    ownerDocument.adoptedStyleSheets = [...ownerDocument.adoptedStyleSheets, rules.sheet];
    jspreadsheetSpillRuleSheets.set(ownerDocument, rules);
  }
  if (rules.widths.has(width)) return;
  rules.widths.add(width);
  rules.sheet.insertRule(
    `[data-jspreadsheet-host] .jss_worksheet > tbody > tr > td[${JSPREADSHEET_CELL_SPILL_ATTRIBUTE}="${width}"] { overflow-clip-margin: ${width}px; }`,
    rules.sheet.cssRules.length,
  );
}

/**
 * Write one cell's spill width, in whole pixels, on jspreadsheet's own `<td>`,
 * read from the worksheet's records; a zero width removes it.
 */
function applyJSpreadsheetCellSpill(
  worksheet: Pick<JSpreadsheetWorksheetInstance, 'records'>,
  rowIndex: number,
  columnIndex: number,
  spillWidth: number,
): void {
  const cellElement = worksheet.records?.[rowIndex]?.[columnIndex]?.element;
  if (!(cellElement instanceof HTMLElement)) return;
  const width = Math.round(spillWidth);
  const spill = width > 0 ? String(width) : null;
  if (width > 0) ensureJSpreadsheetSpillRule(cellElement.ownerDocument, width);
  if (cellElement.getAttribute(JSPREADSHEET_CELL_SPILL_ATTRIBUTE) === spill) return;
  if (spill === null) cellElement.removeAttribute(JSPREADSHEET_CELL_SPILL_ATTRIBUTE);
  else cellElement.setAttribute(JSPREADSHEET_CELL_SPILL_ATTRIBUTE, spill);
}

/** Rewrites the spill width of each cell in the given rows, or in every row. */
function refreshJSpreadsheetWorksheetSpill(
  worksheet: JSpreadsheetWorksheetInstance,
  rowIndexes?: Iterable<number>,
): void {
  const records = worksheet.records;
  if (!records) return;
  const columnWidths = (worksheet.cols ?? []).map((column) =>
    readJSpreadsheetSpillColumnWidth(column?.colElement),
  );
  for (const rowIndex of rowIndexes ?? records.keys()) {
    const row = records[rowIndex];
    if (!row) continue;
    const spillWidths = resolveSpreadsheetRowSpillWidths(
      row.map((record) => describeJSpreadsheetSpillCell(record?.element)),
      columnWidths,
    );
    spillWidths.forEach((spillWidth, columnIndex) => {
      applyJSpreadsheetCellSpill(worksheet, rowIndex, columnIndex, spillWidth);
    });
  }
}

/**
 * Keeps each text cell's spill width current as the engine renders values,
 * edits, decorations and column widths: a changed cell refreshes its own row,
 * and a changed column, row list or worksheet list refreshes every row.
 */
function observeJSpreadsheetCellSpill(
  rootEl: HTMLElement,
  instance: JSpreadsheetInstance,
): () => void {
  const MutationObserverClass = rootEl.ownerDocument.defaultView?.MutationObserver;
  if (!MutationObserverClass) return () => {};
  const rowOwners = new WeakMap<
    Node,
    { worksheet: JSpreadsheetWorksheetInstance; rowIndex: number }
  >();
  let worksheetCount = 0;
  const refreshAll = (): void => {
    const worksheets = listJSpreadsheetWorksheets(instance);
    worksheetCount = worksheets.length;
    worksheets.forEach((worksheet) => {
      worksheet.rows?.forEach((row, rowIndex) => {
        if (row?.element) rowOwners.set(row.element, { worksheet, rowIndex });
      });
      refreshJSpreadsheetWorksheetSpill(worksheet);
    });
  };
  const observer = new MutationObserverClass((mutations) => {
    if (listJSpreadsheetWorksheets(instance).length !== worksheetCount) {
      refreshAll();
      return;
    }
    const changedRows = new Map<JSpreadsheetWorksheetInstance, Set<number>>();
    for (const mutation of mutations) {
      // Walk from a text node to its cell and row, or stop at a changed column or row list.
      let node: Node | null = mutation.target;
      for (let depth = 0; node && depth < 3; depth += 1, node = node.parentNode) {
        const owner = rowOwners.get(node);
        if (owner) {
          const rows = changedRows.get(owner.worksheet) ?? new Set<number>();
          rows.add(owner.rowIndex);
          changedRows.set(owner.worksheet, rows);
          break;
        }
        if (node.nodeName === 'COL' || node.nodeName === 'COLGROUP' || node.nodeName === 'TBODY') {
          refreshAll();
          return;
        }
      }
    }
    changedRows.forEach((rowIndexes, worksheet) =>
      refreshJSpreadsheetWorksheetSpill(worksheet, rowIndexes),
    );
  });
  observer.observe(rootEl, {
    subtree: true,
    childList: true,
    characterData: true,
    attributes: true,
    attributeFilter: ['class', 'style', 'width', 'colspan', 'rowspan'],
  });
  refreshAll();
  return () => observer.disconnect();
}

// --- React component ---------------------------------------------------------

/**
 * Mutable cells owned by the imperative vendor boundary rather than React's
 * render data. The jspreadsheet React wrapper requires an object ref whose
 * `current` property it assigns itself; encapsulating that contract here keeps
 * the engine's mutable instance out of React render-time refs. The latest-props
 * cell likewise feeds callbacks raised by the engine after a React commit.
 */
class JSpreadsheetEngineSlot {
  #instance: JSpreadsheetInstance | null = null;

  get current(): JSpreadsheetInstance | null {
    return this.#instance;
  }

  set current(instance: JSpreadsheetInstance | null) {
    this.#instance = instance;
  }
}

class JSpreadsheetHostPropsSlot {
  #props: JSpreadsheetHostProps;

  constructor(props: JSpreadsheetHostProps) {
    this.#props = props;
  }

  read(): JSpreadsheetHostProps {
    return this.#props;
  }

  publish(props: JSpreadsheetHostProps): void {
    this.#props = props;
  }
}

export function JSpreadsheetHost(
  props: JSpreadsheetHostProps & { ref?: Ref<JSpreadsheetHostHandle> },
): ReactElement {
  const { ref } = props;
  // `engineRef` is handed to the vendor <Spreadsheet> wrapper, which sets its
  // `.current` to the live jspreadsheet instance at mount time. `instanceRef`
  // is our canonical copy, and `engineRootElRef` is the wrapper's engine root
  // <div> (where we stamp the data-jspreadsheet-host marker and which teardown
  // targets).
  const [engineSlot] = useState(() => new JSpreadsheetEngineSlot());
  const instanceRef = useRef<JSpreadsheetInstance | null>(null);
  const engineRootElRef = useRef<HTMLElement | null>(null);
  const [propsSlot] = useState(() => new JSpreadsheetHostPropsSlot(props));
  // The config the engine is (re)mounted with. It intentionally captures the
  // first config rather than following later prop changes; the setData fallback
  // replaces this state and increments `nonce` to force a fresh <Spreadsheet>
  // mount with the replacement dataset.
  const [mountState, setMountState] = useState(() => ({
    config: props.config,
    nonce: 0,
  }));

  // Keep the latest props in a ref so the engine callback bridge always sees
  // the most recent handlers without remounting jspreadsheet on every render.
  useLayoutEffect(() => {
    propsSlot.publish(props);
  });

  // Stable dispatchers — they read through propsRef so prop changes do not
  // tear down the engine.
  const dispatchChange = useCallback(
    (
      row: number,
      col: number,
      oldValue: JSpreadsheetCellValue,
      newValue: JSpreadsheetCellValue,
    ) => {
      propsSlot.read().onChange?.(row, col, oldValue, newValue);
    },
    [propsSlot],
  );
  const dispatchSelectionChange = useCallback(
    (range: JSpreadsheetSelectionRange) => {
      propsSlot.read().onSelectionChange?.(range);
    },
    [propsSlot],
  );
  const dispatchEditStart = useCallback(
    (row: number, col: number) => {
      propsSlot.read().onCellEditStart?.(row, col);
    },
    [propsSlot],
  );
  const dispatchEditEnd = useCallback(
    (row: number, col: number, committedValue: JSpreadsheetCellValue) => {
      propsSlot.read().onCellEditEnd?.(row, col, committedValue);
    },
    [propsSlot],
  );
  const dispatchFormulaAutocompleteRequest = useCallback(
    (query: string, rect: JSpreadsheetAnchorRect) => {
      propsSlot.read().onFormulaAutocompleteRequest?.(query, rect);
    },
    [propsSlot],
  );
  const dispatchContextMenuRequest = useCallback(
    (coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect) => {
      propsSlot.read().onContextMenuRequest?.(coords, rect);
    },
    [propsSlot],
  );
  const dispatchCellCommentRequest = useCallback(
    (coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect) => {
      propsSlot.read().onCellCommentRequest?.(coords, rect);
    },
    [propsSlot],
  );

  // Build the bridged config for the current mount. The vendor <Spreadsheet>
  // wrapper consumes it once per mount, so the
  // engine is created from `mountState.config`; later `config` prop changes
  // do not remount it (the baked-in dispatchers are stable and read through
  // `propsRef`, so the latest handlers still fire). It is rebuilt only when
  // the setData fallback replaces `mountState.config`.
  const bridgedConfig = useMemo(
    () =>
      createBridgedConfig({
        config: mountState.config,
        dispatchChange,
        dispatchSelectionChange,
        dispatchEditStart,
        dispatchEditEnd,
        dispatchFormulaAutocompleteRequest,
        dispatchContextMenuRequest,
        dispatchCellCommentRequest,
      }),
    [
      mountState.config,
      dispatchChange,
      dispatchSelectionChange,
      dispatchEditStart,
      dispatchEditEnd,
      dispatchFormulaAutocompleteRequest,
      dispatchContextMenuRequest,
      dispatchCellCommentRequest,
    ],
  );

  const imperativeHandle = useMemo<JSpreadsheetHostHandle>(
    () => ({
      setCellValue(row: number, col: number, value: JSpreadsheetCellValue): void {
        const worksheet = readActiveWorksheet(instanceRef.current);
        if (!worksheet) return;
        if (typeof worksheet.setValueFromCoords === 'function') {
          worksheet.setValueFromCoords(col, row, value, false);
        } else if (typeof worksheet.setValue === 'function') {
          worksheet.setValue([[col, row]], value);
        }
      },
      getCellValue(row: number, col: number): JSpreadsheetCellValue {
        const worksheet = readActiveWorksheet(instanceRef.current);
        if (!worksheet) return null;
        if (typeof worksheet.getValueFromCoords === 'function') {
          return worksheet.getValueFromCoords(col, row) as JSpreadsheetCellValue;
        }
        if (typeof worksheet.getValue === 'function') {
          return worksheet.getValue(col, row) as JSpreadsheetCellValue;
        }
        return null;
      },
      setData(data: ReadonlyArray<ReadonlyArray<JSpreadsheetCellValue>>): void {
        const worksheet = readActiveWorksheet(instanceRef.current);
        if (worksheet && typeof worksheet.setData === 'function') {
          worksheet.setData(data);
          return;
        }
        // Fallback for engines without an in-place setData: rebuild the mount
        // config with the new dataset and force a React remount of the vendor
        // <Spreadsheet> (a `key` bump). The mount effect's cleanup tears the old
        // engine down and its re-run adopts the freshly mounted one.
        setMountState((current) => ({
          config: { ...current.config, data },
          nonce: current.nonce + 1,
        }));
      },
      getCurrentSelection(): JSpreadsheetSelectionRange | null {
        const worksheet = readActiveWorksheet(instanceRef.current);
        if (
          !worksheet ||
          typeof worksheet.getSelectedColumns !== 'function' ||
          typeof worksheet.getSelectedRows !== 'function'
        ) {
          if (worksheet && typeof worksheet.getSelection === 'function') {
            const selection = worksheet.getSelection();
            if (Array.isArray(selection) && selection.length === 4) {
              const [x1, y1, x2, y2] = selection.map((value: unknown) => safeCoord(value));
              return { x1, y1, x2, y2 };
            }
          }
          return null;
        }
        const columns = worksheet.getSelectedColumns(true) as number[] | undefined;
        const rows = worksheet.getSelectedRows(true) as number[] | undefined;
        if (
          !Array.isArray(columns) ||
          !Array.isArray(rows) ||
          columns.length === 0 ||
          rows.length === 0
        ) {
          return null;
        }
        return {
          x1: Math.min(...columns),
          x2: Math.max(...columns),
          y1: Math.min(...rows),
          y2: Math.max(...rows),
        };
      },
      refresh(): void {
        const worksheet = readActiveWorksheet(instanceRef.current);
        if (typeof worksheet?.refresh === 'function') worksheet.refresh();
      },
      getInstance(): JSpreadsheetInstance | null {
        return instanceRef.current;
      },
    }),
    [],
  );

  useImperativeHandle(ref, () => imperativeHandle, [imperativeHandle]);

  useEffect(() => {
    // The vendor <Spreadsheet> wrapper mounts the engine in its OWN mount-time
    // effect — a CHILD effect, so it runs before this one — and stores the live
    // instance on `engineRef`. We adopt it here, stamp the engine-host marker
    // (and any id/className/style) onto the wrapper's root <div> so the
    // established DOM contract holds (`[data-jspreadsheet-host] > .jtabs-content`
    // and the `:scope > [data-jspreadsheet-host]` worksheet/collaboration
    // lookups), then fire onMounted. The wrapper performs no teardown of its
    // own, so this host owns destroy.
    const instance = engineSlot.current;
    instanceRef.current = instance;

    const rootEl = resolveEngineRootElement(instance);
    engineRootElRef.current = rootEl;
    if (rootEl) {
      rootEl.setAttribute('data-jspreadsheet-host', 'true');
      const { id, className, style } = propsSlot.read();
      if (id) rootEl.id = id;
      if (className) {
        // classList.add (not assignment): jspreadsheet adds its own classes
        // (.jtabs, etc.) to this div and we must not clobber them.
        for (const cls of className.split(/\s+/)) {
          if (cls) rootEl.classList.add(cls);
        }
      }
      if (style) {
        Object.assign(rootEl.style, style);
      }
    }

    // SE-CR-2.2a: keep the adopted engine's text-cell spill widths current.
    const stopCellSpill =
      rootEl && instance ? observeJSpreadsheetCellSpill(rootEl, instance) : null;

    if (instance) {
      propsSlot.read().onMounted?.(instance, imperativeHandle);
    }

    return () => {
      stopCellSpill?.();
      const currentInstance = instanceRef.current;
      if (currentInstance) {
        propsSlot.read().onBeforeDestroy?.(currentInstance);
      }
      const rootForTeardown = engineRootElRef.current;
      if (rootForTeardown) {
        teardownJspreadsheetInstance(rootForTeardown, currentInstance);
      }
      instanceRef.current = null;
      engineSlot.current = null;
      engineRootElRef.current = null;
    };
    // Runs once on mount and again only on a setData-driven remount
    // (`remountNonce` bump + <Spreadsheet> `key` change): the cleanup tears the
    // old engine down and the re-run adopts the freshly mounted one. The
    // dispatch callbacks are stable (useCallback with []).
  }, [engineSlot, imperativeHandle, mountState.nonce, propsSlot]);

  // Render the vendor-provided React wrapper. It owns the engine root <div> and
  // mounts `jspreadsheet(div, options)` into it; `engineRef` receives the live
  // instance. We deliberately render it WITHOUT an extra wrapping element so the
  // engine root stays a direct child of the React mount container — the
  // worksheet/collaboration DOM lookups rely on `:scope > [data-jspreadsheet-host]`.
  // The data-jspreadsheet-host marker and any id/className/style are stamped
  // onto that div in the mount effect above (the wrapper only forwards `id`).
  // `key={remountNonce}` lets the setData fallback force a fresh engine mount.
  return createElement(Spreadsheet as unknown as ComponentType<Record<string, unknown>>, {
    ...bridgedConfig,
    ref: engineSlot,
    key: mountState.nonce,
  });
}

/**
 * JSpreadsheetHost — the engine boundary React component for the spreadsheet
 * editor. See file header for the carveout policy.
 */
export default JSpreadsheetHost;
