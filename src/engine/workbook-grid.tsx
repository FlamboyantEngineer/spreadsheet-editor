'use client';

import { type CSSProperties, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { type CellPortal, clearCalculatedPresentation } from './cell-renderer.js';
import { useGridConfig } from './grid-config.js';
import {
  type GridContext,
  type GridFrame,
  type GridRefs,
  gridWorksheet,
  useGridRefs,
  type WorkbookGridProps,
} from './grid-context.js';
import { useGridController } from './grid-controller.js';
import {
  engineDestroying,
  engineEditEnded,
  engineEditStarted,
  engineMounted,
  engineSelected,
} from './grid-engine-events.js';
import { useReferenceFrames } from './grid-frames.js';
import { syncFrozenColumns } from './grid-geometry.js';
import { useSurfaceHandlers } from './grid-surface.js';
import { useWorksheetSync } from './grid-sync.js';
import { useHeaderResize } from './header-resize.js';
import { JSpreadsheetHost } from './jspreadsheet-host.js';
import type { JSpreadsheetHostHandle } from './jspreadsheet-types.js';
import { useNativeFormula } from './native-formula.js';
import { createSelectionPublication } from './selection-publication.js';
import { useLatest } from './use-latest.js';

/** Stops the grid's observers and releases its painted presentation on unmount. */
function useGridTeardown(refs: GridRefs) {
  useEffect(
    () => () => {
      if (refs.commentClick.current) clearTimeout(refs.commentClick.current);
      refs.stopBorderCopies.current?.();
      refs.stopRowWindow.current?.();
      if (refs.handle.current) clearCalculatedPresentation(refs.handle.current);
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps -- tears down once, when the grid unmounts
    [],
  );
}

/** Ends a selection drag when the window blurs or any pointer is released. */
function useSelectionRelease(context: GridContext) {
  const { selectionPublication } = context;
  useEffect(() => {
    const finish = () => {
      context.draggingSelection.current = false;
      selectionPublication.finish();
      context.scheduleReferences();
    };
    const stop = context.latest.current.subscribeWindowBlur?.(finish);
    if (!stop) window.addEventListener('blur', finish);
    window.addEventListener('pointerup', finish);
    window.addEventListener('pointercancel', finish);
    return () => {
      stop?.();
      window.removeEventListener('blur', finish);
      window.removeEventListener('pointerup', finish);
      window.removeEventListener('pointercancel', finish);
      selectionPublication.dispose();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- window listeners follow the publication; scheduleReferences and latest read refs
  }, [selectionPublication]);
}

/** Header resizing over the engine's measures; a resize moves the outlines with it. */
function useGridResize(context: GridContext, props: WorkbookGridProps) {
  return useHeaderResize({
    root: context.surface,
    worksheet: props.worksheet,
    getWorksheet: context.worksheet,
    canResize: !props.readOnly && !props.reviewMode,
    zoom: props.zoom,
    onChange: props.onChange,
    onGeometryChange: () => {
      const sheet = context.worksheet();
      if (sheet) syncFrozenColumns(sheet);
      context.scheduleReferences();
    },
    subscribeBlur: props.subscribeWindowBlur,
  });
}

function frameClass(frame: GridFrame) {
  if (frame.local) return 'se-editor__selection-outline';
  return frame.kind ? 'se-editor__presence-outline' : 'se-editor__reference-outline';
}

/** The outlines over the grid, naming the peer whose selection the pointer is over. */
function gridOutlines(
  frames: GridFrame[],
  { clip, hoveredPeer }: { clip: CSSProperties; hoveredPeer: string | null },
) {
  if (frames.length === 0) return false;
  return (
    <div className="se-editor__reference-frame" aria-hidden="true" style={clip}>
      {frames.map((frame) => (
        <div
          className={frameClass(frame)}
          data-kind={frame.kind}
          data-color={frame.colorIndex}
          key={frame.key}
          style={frame.style}
        >
          {frame.name && frame.key === hoveredPeer && <span>{frame.name}</span>}
        </div>
      ))}
    </div>
  );
}

export function WorkbookGrid(props: WorkbookGridProps) {
  const latest = useLatest(props);
  const surfaceRef = useRef<HTMLDivElement | null>(null);
  const handle = useRef<JSpreadsheetHostHandle | null>(null);
  const refs = useGridRefs(props, { surface: surfaceRef, handle });
  useGridTeardown(refs);
  const worksheet = () => gridWorksheet(refs.handle.current);
  const native = useNativeFormula({
    onDraft: (value) => {
      const coords = refs.nativeCoords.current;
      latest.current.onEditDraft?.(value);
      latest.current.onNativeDraft?.(value === null || !coords ? null : { ...coords, text: value });
    },
    root: surfaceRef,
    functions: props.formulaFunctions,
    includeBuiltins: props.includeBuiltinFormulaFunctions,
  });
  const [portals, setPortals] = useState<CellPortal[]>([]);
  const [hoveredPeer, setHoveredPeer] = useState<string | null>(null);
  const [selectionPublication] = useState(() =>
    createSelectionPublication((range) => latest.current.onSelectionChange(range)),
  );
  const [revision, setRevision] = useState(0);
  const { frames, clip, scheduleReferences } = useReferenceFrames(
    { ...refs, latest, worksheet },
    { ...props, revision },
  );
  const context: GridContext = {
    ...refs,
    latest,
    worksheet,
    scheduleReferences,
    setRevision,
    setPortals,
    selectionPublication,
  };
  useSelectionRelease(context);
  const resize = useGridResize(context, props);
  const controller = useGridController(context, props.onController);
  const config = useGridConfig(context, props);
  const surface = { ...context, native, resize, controller };
  useWorksheetSync(surface, props, revision);
  const handlers = useSurfaceHandlers(surface, { props, frames, clip, setHoveredPeer });
  return (
    <div
      ref={surfaceRef}
      tabIndex={-1}
      className="se-editor__grid-surface"
      style={{ position: 'relative', minWidth: 0 }}
      {...handlers}
    >
      <JSpreadsheetHost
        ref={handle}
        className="se-editor__engine"
        config={config}
        onMounted={(_instance, api) => engineMounted(surface, api)}
        onBeforeDestroy={() => engineDestroying(surface)}
        onSelectionChange={(range) => engineSelected(surface, range)}
        onContextMenuRequest={props.onContextMenu}
        onCellCommentRequest={props.onCommentRequest}
        onCellEditStart={(y, x) => engineEditStarted(surface, { x, y })}
        onCellEditEnd={() => engineEditEnded(surface)}
      />
      {native.popup}
      {native.overlay}
      {portals.map((portal) => createPortal(portal.content, portal.target, portal.key))}
      {gridOutlines(frames, { clip, hoveredPeer })}
    </div>
  );
}
