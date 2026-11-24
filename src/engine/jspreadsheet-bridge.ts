import { jspreadsheet } from '@jspreadsheet-ce/react';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
  JSpreadsheetCellValue,
  JSpreadsheetConfig,
  JSpreadsheetInstance,
  JSpreadsheetSelectionRange,
  JSpreadsheetSpreadsheetInstance,
  JSpreadsheetVendorCallback,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';

/**
 * Options forwarded to `createBridgedConfig`. It wires the jspreadsheet-level
 * callbacks (`onchange`, `onselection`, `oneditionstart`, `oneditionend`,
 * `contextMenu`, `oncomments`) to the host's typed dispatchers.
 */
export interface JspreadsheetBridgeOptions {
  readonly config: JSpreadsheetConfig;
  readonly dispatchSelectionChange: (range: JSpreadsheetSelectionRange) => void;
  readonly dispatchEditStart: (row: number, col: number) => void;
  readonly dispatchEditEnd: (
    row: number,
    col: number,
    committedValue: JSpreadsheetCellValue,
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

/** An instance's own root element: its `el`, its `element`, or a test stub's `root`. */
function ownRootElement(
  first: JSpreadsheetSpreadsheetInstance | JSpreadsheetWorksheetInstance,
): HTMLElement | null {
  if ('el' in first && first.el) return first.el;
  if ('element' in first && first.element) return first.element;
  if ('root' in first && first.root instanceof HTMLElement) return first.root;
  return null;
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
export function resolveEngineRootElement(
  instance: JSpreadsheetInstance | null,
): HTMLElement | null {
  if (!instance) return null;
  const first = Array.isArray(instance) ? instance[0] : instance;
  if (!first || typeof first !== 'object') return null;
  const parent = 'parent' in first ? first.parent : null;
  if (parent && (parent.el || parent.element)) return parent.el ?? parent.element ?? null;
  return ownRootElement(first);
}

/** Report vendor callback failures while allowing typed dispatchers to run. */
function reportVendorCallbackError(callbackName: string, error: unknown): void {
  console.error(`[JSpreadsheetHost] vendor ${callbackName} callback threw`, error);
}

/** The vendor's arguments, cut or padded to the callback's own parameter count. */
const fixedArguments = (args: readonly unknown[], arity: number) =>
  Array.from({ length: arity }, (_value, index) => args[index]);

/**
 * The caller's own handler for a vendor callback, called with the vendor's arguments as a
 * callback of `arity` parameters; a throw is reported, so the host's dispatcher still runs.
 */
function callerCallback(config: JSpreadsheetConfig, name: string, arity: number) {
  const handler = config[name];
  return (args: readonly unknown[]): void => {
    if (typeof handler !== 'function') return;
    try {
      (handler as JSpreadsheetVendorCallback)(...fixedArguments(args, arity));
    } catch (error) {
      reportVendorCallbackError(name, error);
    }
  };
}

/** A context menu request dispatches first; the caller's handler then decides the items. */
function contextMenuCallback(options: JspreadsheetBridgeOptions) {
  const handler = options.config.contextMenu;
  return (...args: unknown[]): unknown => {
    const [, colIndex, rowIndex, event, items, , x, y] = fixedArguments(args, 8);
    const target =
      event && typeof event === 'object' && 'target' in event
        ? (event as { target: unknown }).target
        : null;
    options.dispatchContextMenuRequest(
      { x: safeCoord(colIndex ?? x), y: safeCoord(rowIndex ?? y) },
      rectFromElement(target),
    );
    if (typeof handler !== 'function') return items;
    try {
      return (handler as JSpreadsheetVendorCallback)(...fixedArguments(args, 8));
    } catch (error) {
      reportVendorCallbackError('contextMenu', error);
      return items;
    }
  };
}

/** Worksheet creation passes every vendor argument through to the caller's handler. */
function createWorksheetCallback(config: JSpreadsheetConfig) {
  const handler = config.oncreateworksheet;
  return (...args: unknown[]): unknown => {
    try {
      if (typeof handler === 'function') return (handler as JSpreadsheetVendorCallback)(...args);
    } catch (error) {
      reportVendorCallbackError('oncreateworksheet', error);
    }
    return undefined;
  };
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
  // Vendor arguments: (worksheet, cell, x, y, ...) for cell callbacks; (instance, x1, y1, x2, y2).
  const change = callerCallback(config, 'onchange', 6),
    selection = callerCallback(config, 'onselection', 5),
    editionStart = callerCallback(config, 'oneditionstart', 4),
    editionEnd = callerCallback(config, 'oneditionend', 5),
    comments = callerCallback(config, 'oncomments', 4);
  return {
    ...config,
    onchange: (...args: unknown[]) => change(args),
    onselection: (...args: unknown[]) => {
      selection(args);
      const [, x1, y1, x2, y2] = args;
      options.dispatchSelectionChange({
        x1: safeCoord(x1),
        y1: safeCoord(y1),
        x2: safeCoord(x2),
        y2: safeCoord(y2),
      });
    },
    oneditionstart: (...args: unknown[]) => {
      editionStart(args);
      options.dispatchEditStart(safeCoord(args[3]), safeCoord(args[2]));
    },
    oneditionend: (...args: unknown[]) => {
      editionEnd(args);
      const [, , x, y, value] = args;
      options.dispatchEditEnd(safeCoord(y), safeCoord(x), value as JSpreadsheetCellValue);
    },
    contextMenu: contextMenuCallback(options),
    oncomments: (...args: unknown[]) => {
      comments(args);
      const [, cell, x, y] = args;
      options.dispatchCellCommentRequest(
        { x: safeCoord(x), y: safeCoord(y) },
        rectFromElement(cell),
      );
    },
    oncreateworksheet: createWorksheetCallback(config),
  };
}

/**
 * Destroy a jspreadsheet instance and clear the host children. The vendor
 * `<Spreadsheet>` wrapper performs no teardown, so this host calls
 * `jspreadsheet.destroy(root, true)` (the namespace is re-exported by
 * `@jspreadsheet-ce/react`) and then clears any residual engine DOM.
 */
export function teardownJspreadsheetInstance(
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
