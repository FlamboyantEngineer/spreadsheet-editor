'use client';

/*
 * The one place that mounts the Jspreadsheet engine (only `src/engine/` may touch
 * engine-owned DOM). The official @jspreadsheet-ce/react Spreadsheet wrapper owns
 * the engine root and instance; this host preserves data-jspreadsheet-host on that
 * root and bridges typed callbacks through createBridgedConfig. React chrome
 * consumes request events outside the engine-owned DOM.
 *
 * The wrapper has no teardown, so unmount calls onBeforeDestroy and
 * teardownJspreadsheetInstance. Worksheet changes reach the mounted engine
 * through the grid synchronization code.
 */

// The vendor-provided React wrapper for jspreadsheet-ce. `<Spreadsheet>` owns
// the engine root <div> and mounts `jspreadsheet(div, options)` into it,
// exposing the live instance on the `ref` it receives. The package also
// re-exports the raw `jspreadsheet` namespace, which we use only for engine
// teardown (`jspreadsheet.destroy`) — the host never constructs the engine
// itself, so it does NOT import `jspreadsheet-ce` directly.
import { Spreadsheet } from '@jspreadsheet-ce/react';
import {
  type ComponentType,
  createElement,
  type ReactElement,
  type Ref,
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  createBridgedConfig,
  type JspreadsheetBridgeOptions,
  resolveEngineRootElement,
  teardownJspreadsheetInstance,
} from './jspreadsheet-bridge.js';
import { observeJSpreadsheetCellSpill } from './jspreadsheet-spill.js';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
  JSpreadsheetCellValue,
  JSpreadsheetConfig,
  JSpreadsheetHostHandle,
  JSpreadsheetInstance,
  JSpreadsheetSelectionRange,
} from './jspreadsheet-types.js';

/**
 * Prop contract of `JSpreadsheetHost`. Every prop below is optional except
 * `config`.
 */
interface JSpreadsheetHostProps {
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

  /**
   * Fired once after the jspreadsheet instance is created. The chrome typically
   * needs the instance for direct commands such as inserting rows or applying
   * formats.
   */
  onMounted?:
    ((instance: JSpreadsheetInstance, handle: JSpreadsheetHostHandle) => void) | undefined;

  /** Fired once just before the instance is destroyed (unmount or remount). */
  onBeforeDestroy?: ((instance: JSpreadsheetInstance) => void) | undefined;

  /** Fired when the selection range changes (mouse, keyboard, programmatic). */
  onSelectionChange?: ((coords: JSpreadsheetSelectionRange) => void) | undefined;

  /** Fired when the user begins editing a cell. */
  onCellEditStart?: ((row: number, col: number) => void) | undefined;

  /** Fired when the user commits or cancels a cell edit. */
  onCellEditEnd?:
    ((row: number, col: number, committedValue: JSpreadsheetCellValue) => void) | undefined;

  /**
   * REQUEST event: the user opened the cell context menu. The editor renders
   * the React context menu portal anchored at `anchorRect`.
   */
  onContextMenuRequest?:
    ((coords: JSpreadsheetCellCoords, anchorRect: JSpreadsheetAnchorRect) => void) | undefined;

  /**
   * REQUEST event: the user requested the comment popover for a cell. The
   * editor renders the React comment popover portal anchored at `anchorRect`.
   */
  onCellCommentRequest?:
    ((coords: JSpreadsheetCellCoords, anchorRect: JSpreadsheetAnchorRect) => void) | undefined;
}

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

type Dispatchers = Omit<JspreadsheetBridgeOptions, 'config'>;

/**
 * Stable dispatchers — they read through the props slot so prop changes do not
 * tear down the engine.
 */
function useHostDispatchers(propsSlot: JSpreadsheetHostPropsSlot): Dispatchers {
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
  return useMemo(
    () => ({
      dispatchSelectionChange,
      dispatchEditStart,
      dispatchEditEnd,
      dispatchContextMenuRequest,
      dispatchCellCommentRequest,
    }),
    [
      dispatchSelectionChange,
      dispatchEditStart,
      dispatchEditEnd,
      dispatchContextMenuRequest,
      dispatchCellCommentRequest,
    ],
  );
}

/**
 * Stamp the engine-host marker and class name onto the wrapper's
 * root <div> so the established DOM contract holds
 * (`[data-jspreadsheet-host] > .jtabs-content` and the
 * `:scope > [data-jspreadsheet-host]` worksheet/collaboration lookups).
 */
function stampEngineRoot(rootEl: HTMLElement, className?: string) {
  rootEl.setAttribute('data-jspreadsheet-host', 'true');
  // classList.add (not assignment): jspreadsheet adds its own classes
  // (.jtabs, etc.) to this div and we must not clobber them.
  for (const cls of className?.split(/\s+/) ?? []) if (cls) rootEl.classList.add(cls);
}

export function JSpreadsheetHost(
  props: JSpreadsheetHostProps & { ref?: Ref<JSpreadsheetHostHandle> },
): ReactElement {
  const { ref } = props;
  // `engineSlot` is handed to the vendor <Spreadsheet> wrapper, which sets its
  // `.current` to the live jspreadsheet instance at mount time.
  const [engineSlot] = useState(() => new JSpreadsheetEngineSlot());
  const instanceRef = useRef<JSpreadsheetInstance | null>(null);
  const [propsSlot] = useState(() => new JSpreadsheetHostPropsSlot(props));
  const [config] = useState(() => props.config);

  // Keep the latest props in a slot so the engine callback bridge always sees
  // the most recent handlers without remounting jspreadsheet on every render.
  useLayoutEffect(() => {
    propsSlot.publish(props);
  });
  const dispatchers = useHostDispatchers(propsSlot);

  const bridgedConfig = useMemo(
    () => createBridgedConfig({ config, ...dispatchers }),
    [config, dispatchers],
  );

  const imperativeHandle = useMemo<JSpreadsheetHostHandle>(
    () => ({ getInstance: () => instanceRef.current }),
    [],
  );

  useImperativeHandle(ref, () => imperativeHandle, [imperativeHandle]);

  useEffect(() => {
    // The vendor <Spreadsheet> wrapper mounts the engine in its OWN mount-time
    // effect — a CHILD effect, so it runs before this one — and stores the live
    // instance on `engineSlot`. We adopt it here, stamp the engine root, then
    // fire onMounted. The wrapper performs no teardown of its own, so this host
    // owns destroy.
    const instance = engineSlot.current;
    instanceRef.current = instance;
    const rootEl = resolveEngineRootElement(instance);
    if (rootEl) stampEngineRoot(rootEl, propsSlot.read().className);
    // Keep the adopted engine's text-cell spill widths current.
    const stopCellSpill =
      rootEl && instance ? observeJSpreadsheetCellSpill(rootEl, instance) : null;
    if (instance) propsSlot.read().onMounted?.(instance, imperativeHandle);
    return () => {
      stopCellSpill?.();
      const currentInstance = instanceRef.current;
      if (currentInstance) propsSlot.read().onBeforeDestroy?.(currentInstance);
      if (rootEl) teardownJspreadsheetInstance(rootEl, currentInstance);
      instanceRef.current = null;
      engineSlot.current = null;
    };
  }, [engineSlot, imperativeHandle, propsSlot]);

  // Render the vendor-provided React wrapper. It owns the engine root <div> and
  // mounts `jspreadsheet(div, options)` into it; `engineSlot` receives the live
  // instance. We deliberately render it WITHOUT an extra wrapping element so the
  // engine root stays a direct child of the React mount container — the
  // worksheet/collaboration DOM lookups rely on `:scope > [data-jspreadsheet-host]`.
  return createElement(Spreadsheet as unknown as ComponentType<Record<string, unknown>>, {
    ...bridgedConfig,
    ref: engineSlot,
  });
}
