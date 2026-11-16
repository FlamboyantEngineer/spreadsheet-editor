import type JSpreadsheetCore from 'jspreadsheet-ce';

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
 * The editor positions the React overlay it renders through a portal with this
 * rect. The rect is in viewport coordinates (the same convention as
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
export type JSpreadsheetVendorCallback = {
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
  /** The vendor's optional fourth `force` argument is never passed. */
  setValueFromCoords?: (x: number, y: number, value: JSpreadsheetCellValue) => void;
  setWidth?: (columnIndex: number | number[], width: number | number[]) => void;
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
 * methods are the only way for the editor to drive the engine from outside the
 * host. Anything else should be expressed as new props/events.
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
