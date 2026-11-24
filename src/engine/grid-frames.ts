'use client';

import { type CSSProperties, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { flushSync } from 'react-dom';
import type { CellSelection } from '../model/editing.js';
import {
  FORMULA_REFERENCE_HIGHLIGHT_COLORS,
  type FormulaReferenceHighlight,
} from '../model/formula-reference-parser.js';
import type { Worksheet } from '../model/workbook.js';
import { collaboratorColorIndex } from './collaborator-color.js';
import type { GridContext, GridFrame, WorkbookGridProps } from './grid-context.js';
import {
  clipGridSelectionHandle,
  gridBodyBounds,
  gridViewport,
  syncFrozenColumnOffsets,
} from './grid-geometry.js';
import { gridSelectionGeometry, measureGridSelection } from './grid-selection.js';

type FrameBase = Omit<
  GridContext,
  'scheduleReferences' | 'setRevision' | 'setPortals' | 'selectionPublication'
>;
type SelectionStyles = (range: CellSelection) => CSSProperties[];
type Body = ReturnType<typeof gridBodyBounds>;

/** The clip box over the visible cell body, in the surface's coordinates. */
function clipStyle(root: HTMLDivElement, body: Body): CSSProperties {
  const base = root.getBoundingClientRect();
  return {
    position: 'absolute',
    pointerEvents: 'none',
    overflow: 'hidden',
    boxSizing: 'border-box',
    zIndex: 5,
    left: body.left - base.left - root.clientLeft,
    top: body.top - base.top - root.clientTop,
    width: body.right - body.left,
    height: body.bottom - body.top,
  };
}

/** A reference's cells; a whole row or column reaches the end of the sheet's extent. */
function referenceRange(
  ref: FormulaReferenceHighlight,
  extent: { columns: number; rows: number },
): CellSelection {
  const r = ref.target;
  const x1 = 'startColumn' in r ? r.startColumn : 0,
    y1 = 'startRow' in r ? r.startRow : 0;
  const x2 = 'endColumn' in r ? r.endColumn : extent.columns - 1;
  const y2 = 'endRow' in r ? r.endRow : extent.rows - 1;
  return { x1, y1, x2, y2 };
}

function referenceColor(colorIndex: number): string {
  const index = colorIndex % FORMULA_REFERENCE_HIGHLIGHT_COLORS.length;
  return `var(--se-reference-${index}, ${FORMULA_REFERENCE_HIGHLIGHT_COLORS[index]})`;
}

/** One frame per distinct formula reference, in its colour. */
function referenceFrames(
  current: WorkbookGridProps,
  selectionStyles: SelectionStyles,
  extent: { columns: number; rows: number },
) {
  const frames: GridFrame[] = [];
  const seen = new Set<string>();
  for (const ref of current.formulaReferences ?? []) {
    if (seen.has(ref.normalizedKey)) continue;
    seen.add(ref.normalizedKey);
    selectionStyles(referenceRange(ref, extent)).forEach((style, index) =>
      frames.push({
        key: `${ref.normalizedKey}:${index}`,
        style: { ...style, borderColor: referenceColor(ref.colorIndex) },
      }),
    );
  }
  return frames;
}

/** Each peer's selection and edit lock on this sheet. */
function peerFrames(current: WorkbookGridProps, selectionStyles: SelectionStyles) {
  const frames: GridFrame[] = [];
  for (const person of current.collaborators ?? []) {
    if (person.sheetId !== current.worksheet.id) continue;
    for (const kind of ['selection', 'lock'] as const) {
      const range = kind === 'selection' ? person.selection : person.editLock;
      if (!range) continue;
      selectionStyles(range).forEach((style, index) =>
        frames.push({
          key: `${person.id}:${kind}:${index}`,
          kind,
          name: person.name,
          colorIndex: collaboratorColorIndex(person.id),
          style,
        }),
      );
    }
  }
  return frames;
}

/** Selection outlines for a range, in the clip box's coordinates. */
function selectionStylesOf(
  sheet: NonNullable<ReturnType<FrameBase['worksheet']>>,
  { worksheet, viewport, body }: { worksheet: Worksheet; viewport: HTMLElement; body: Body },
): SelectionStyles {
  const offset = { left: body.left - body.origin.left, top: body.top - body.origin.top };
  const geometry = gridSelectionGeometry(sheet, body, worksheet.freezeColumns ?? 0);
  return (range) =>
    measureGridSelection(sheet, { sheet: worksheet, selection: range, geometry }, viewport).map(
      (style) => ({
        ...style,
        left: Number(style.left ?? 0) - offset.left,
        top: Number(style.top ?? 0) - offset.top,
      }),
    );
}

/** Measures every outline the grid draws, and the clip box they are drawn in. */
function measureFrames(
  base: FrameBase,
  {
    setFrames,
    setClip,
  }: {
    setFrames: (frames: GridFrame[]) => void;
    setClip: (clip: CSSProperties) => void;
  },
) {
  base.latest.current.onViewportChange?.();
  const current = base.latest.current,
    root = base.surface.current,
    sheet = base.worksheet();
  if (!root || !sheet) {
    setFrames([]);
    return;
  }
  const viewport = gridViewport(sheet) ?? root;
  if (base.scrolled.current) {
    syncFrozenColumnOffsets(sheet);
    base.scrolled.current = false;
  }
  const body = gridBodyBounds(sheet, viewport);
  const selectionStyles = selectionStylesOf(sheet, {
    worksheet: current.worksheet,
    viewport,
    body,
  });
  setClip(clipStyle(root, body));
  const local = selectionStyles(base.visualSelection.current).map((style, index): GridFrame => ({
    key: `local:${index}`,
    local: true,
    style,
  }));
  setFrames([
    ...local,
    ...referenceFrames(current, selectionStyles, {
      columns: sheet.cols?.length ?? 1,
      rows: sheet.rows?.length ?? 1,
    }),
    ...peerFrames(current, selectionStyles),
  ]);
  clipGridSelectionHandle(sheet, body);
}

/** The outlines over the grid, measured once per frame after anything that moves them. */
export function useReferenceFrames(
  base: FrameBase,
  inputs: Pick<WorkbookGridProps, 'formulaReferences' | 'collaborators' | 'zoom'> & {
    revision: number;
  },
) {
  const [frames, setFrames] = useState<GridFrame[]>([]);
  const [clip, setClip] = useState<CSSProperties>({});
  const frameRequest = useRef(0);
  function scheduleReferences() {
    if (frameRequest.current) return;
    frameRequest.current = requestAnimationFrame(() => {
      frameRequest.current = 0;
      // The engine highlight already changed. Commit its matching outline before
      // this paint, without waiting for parent selection/summary publication.
      flushSync(() => measureFrames(base, { setFrames, setClip }));
    });
  }
  const { formulaReferences, collaborators, zoom, revision } = inputs;
  useLayoutEffect(() => {
    scheduleReferences();
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the frame is scheduled through refs when these inputs change
  }, [formulaReferences, collaborators, zoom, revision]);
  useEffect(() => {
    scheduleReferences();
    const observer = new ResizeObserver(scheduleReferences);
    if (base.surface.current) observer.observe(base.surface.current);
    return () => {
      observer.disconnect();
      cancelAnimationFrame(frameRequest.current);
      cancelAnimationFrame(base.focusRequest.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- observes once for the grid's life; scheduleReferences reads refs only
  }, []);
  return { frames, clip, scheduleReferences };
}
