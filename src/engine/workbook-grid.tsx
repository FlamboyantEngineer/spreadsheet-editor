'use client';

import { type CSSProperties, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { syncWorksheetPresentation } from './worksheet-presentation.js';
import { observeGridBorderCopies } from './shared-grid-borders.js';
import {
  gridViewport,
  gridZoomScale,
  measureGridSelection,
  restoreGridScroll,
  scrollGridCellIntoView,
  syncFrozenColumns,
} from './grid-geometry.js';
import { useNativeFormula } from './native-formula.js';
import { useHeaderResize, deliveredHeader } from './header-resize.js';
import type { CellValue } from '../calculation/values.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetComment, SpreadsheetSuggestion } from '../model/review.js';
import {
  FORMULA_REFERENCE_HIGHLIGHT_COLORS,
  type FormulaReferenceHighlight,
} from '../model/formula-reference-parser.js';
import { selectionAddress } from '../model/editing.js';
import type { GridController } from '../editor/contracts.js';
import { writeCells, parseDelimited, type CellSelection } from '../model/editing.js';
import {
  JSpreadsheetHost,
  focusJSpreadsheetWorksheet,
  releaseJSpreadsheetKeyboardOwner,
  type JSpreadsheetHostHandle,
  type JSpreadsheetSelectionRange,
  type JSpreadsheetWorksheetInstance,
  type JSpreadsheetAnchorRect,
  type JSpreadsheetCellCoords,
} from './jspreadsheet-host.js';
import {
  collaboratorSheet,
  collaboratorColorIndex,
  selectionLock,
  type SpreadsheetCollaborator,
} from '../model/presence.js';
import { renderCalculatedCells, renderReviewMarkers, renderContributors } from './cell-renderer.js';

export interface WorkbookGridProps {
  worksheet: Worksheet;
  values?: CellValue[][];
  display?: DisplayOptions;
  readOnly?: boolean;
  /** Grid-only zoom percentage. Chrome and the scroll viewport keep their size. */
  zoom?: number;
  unavailableText?: string;
  selection: CellSelection;
  reviewMode?: boolean;
  onReviewChange?(values: CellInput[][], x: number, y: number): void;
  onPasteCells?(values: CellInput[][], x: number, y: number): void;
  onBlockedAction?(message: string): void;
  showFormulas?: boolean;
  comments?: SpreadsheetComment[];
  suggestions?: SpreadsheetSuggestion[];
  collaborators?: readonly SpreadsheetCollaborator[];
  showContributors?: boolean;
  worksheetIndex?: number;
  onEditStateChange?(editing: boolean): void;
  formulaReferences?: FormulaReferenceHighlight[];
  pickingFormula?: boolean;
  onFormulaReference?(range: string): void;
  onHeaderDelete?(axis: 'row' | 'column'): void;
  onEditDraft?(value: string | null): void;
  onController(controller: GridController | null): void;
  onContextMenu(coords: JSpreadsheetCellCoords, rect: JSpreadsheetAnchorRect): void;
  subscribeWindowBlur?: (callback: () => void) => () => void;
  onViewportChange?(): void;
  onCommentRequest(coords: JSpreadsheetCellCoords): void;
  onChange(worksheet: Worksheet): void;
  onSelectionChange(range: JSpreadsheetSelectionRange): void;
}

export default function WorkbookGrid(props: WorkbookGridProps) {
  const handle = useRef<JSpreadsheetHostHandle | null>(null);
  const stopBorderCopies = useRef<(() => void) | null>(null);
  const surface = useRef<HTMLDivElement | null>(null);
  const picking = useRef<CellSelection | null>(null);
  const inCell = useRef(false);
  const commentClick = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (commentClick.current) clearTimeout(commentClick.current);
      stopBorderCopies.current?.();
    },
    [],
  );
  const selectedHeader = useRef<'row' | 'column' | null>(null);
  const applied = useRef<Worksheet | null>(null);
  const appliedZoom = useRef(100);
  const native = useNativeFormula((value) => latest.current.onEditDraft?.(value), surface);
  const [frames, setFrames] = useState<
    {
      key: string;
      style: CSSProperties;
      local?: boolean;
      kind?: 'selection' | 'lock';
      name?: string;
      colorIndex?: number;
    }[]
  >([]);
  const [hoveredPeer, setHoveredPeer] = useState<string | null>(null);
  const [clip, setClip] = useState<CSSProperties>({});
  const frameRequest = useRef(0);
  const latest = useRef(props),
    suppress = useRef(false),
    selecting = useRef(false);
  const [revision, setRevision] = useState(0);
  useLayoutEffect(() => {
    latest.current = props;
  });
  function worksheet(): JSpreadsheetWorksheetInstance | null {
    const instance = handle.current?.getInstance();
    if (!instance) return null;
    return Array.isArray(instance)
      ? (instance[0] ?? null)
      : 'worksheets' in instance
        ? (instance.worksheets?.[0] ?? null)
        : (instance as JSpreadsheetWorksheetInstance);
  }
  function pickCell(event: { target: EventTarget | null; clientX: number; clientY: number }) {
    const root = surface.current;
    let target = event.target instanceof Element ? event.target : null;
    if (picking.current)
      target = root?.ownerDocument.elementFromPoint(event.clientX, event.clientY) ?? target;
    if (!root || !target || !root.contains(target)) return null;
    while (target && target !== root) {
      if (
        target instanceof HTMLElement &&
        target.dataset.x !== undefined &&
        target.dataset.y !== undefined
      ) {
        const x = Number(target.dataset.x),
          y = Number(target.dataset.y);
        if (Number.isInteger(x) && Number.isInteger(y) && x >= 0 && y >= 0) return { x, y };
      }
      target = target.parentElement;
    }
    return null;
  }
  function measureReferences() {
    latest.current.onViewportChange?.();
    const current = latest.current,
      root = surface.current,
      sheet = worksheet();
    if (!root || !sheet) {
      setFrames([]);
      return;
    }
    syncFrozenColumns(sheet);
    const viewport = gridViewport(sheet) ?? root;
    const base = root.getBoundingClientRect(),
      bounds = viewport.getBoundingClientRect();
    setClip({
      position: 'absolute',
      pointerEvents: 'none',
      overflow: 'hidden',
      boxSizing: 'border-box',
      zIndex: 5,
      left: bounds.left - base.left + viewport.clientLeft - root.clientLeft,
      top: bounds.top - base.top + viewport.clientTop - root.clientTop,
      width: viewport.clientWidth,
      height: viewport.clientHeight,
    });
    const measured: typeof frames = measureGridSelection(
      sheet,
      current.worksheet,
      current.selection,
      viewport,
    ).map((style, index) => ({ key: `local:${index}`, local: true, style }));
    const seen = new Set<string>();
    for (const ref of current.formulaReferences ?? []) {
      if (seen.has(ref.normalizedKey)) continue;
      seen.add(ref.normalizedKey);
      const r = ref.target;
      const x1 = 'startColumn' in r ? r.startColumn : 0,
        y1 = 'startRow' in r ? r.startRow : 0;
      const x2 =
        'endColumn' in r
          ? r.endColumn
          : Math.max(
              (current.worksheet.minDimensions?.[0] ?? 26) - 1,
              ...current.worksheet.data.map((row) => row.length - 1),
            );
      const y2 =
        'endRow' in r
          ? r.endRow
          : Math.max(
              (current.worksheet.minDimensions?.[1] ?? 100) - 1,
              current.worksheet.data.length - 1,
            );
      measureGridSelection(sheet, current.worksheet, { x1, y1, x2, y2 }, viewport).forEach(
        (style, index) =>
          measured.push({
            key: `${ref.normalizedKey}:${index}`,
            style: {
              ...style,
              borderColor: `var(--se-reference-${ref.colorIndex % FORMULA_REFERENCE_HIGHLIGHT_COLORS.length}, ${FORMULA_REFERENCE_HIGHLIGHT_COLORS[ref.colorIndex % FORMULA_REFERENCE_HIGHLIGHT_COLORS.length]})`,
            },
          }),
      );
    }
    for (const person of current.collaborators ?? []) {
      const target = collaboratorSheet(
        { schemaVersion: 1, worksheets: [current.worksheet] },
        person,
      );
      if (
        !target &&
        (person.sheetId || person.sheetName || person.sheetIndex !== current.worksheetIndex)
      )
        continue;
      for (const kind of ['selection', 'lock'] as const) {
        const range = kind === 'selection' ? person.selection : person.editLock;
        if (!range) continue;
        measureGridSelection(sheet, current.worksheet, range, viewport).forEach((style, index) =>
          measured.push({
            key: `${person.id}:${kind}:${index}`,
            kind,
            name: person.name,
            colorIndex: collaboratorColorIndex(person.id),
            style,
          }),
        );
      }
    }
    setFrames(measured);
  }
  function scheduleReferences() {
    cancelAnimationFrame(frameRequest.current);
    frameRequest.current = requestAnimationFrame(measureReferences);
  }
  useEffect(() => {
    scheduleReferences();
    const observer = new ResizeObserver(scheduleReferences);
    if (surface.current) observer.observe(surface.current);
    return () => {
      observer.disconnect();
      cancelAnimationFrame(frameRequest.current);
    };
  }, [props.formulaReferences, props.collaborators, props.selection, props.zoom, revision]);
  const resize = useHeaderResize({
    root: surface,
    worksheet: props.worksheet,
    getWorksheet: worksheet,
    canResize: !props.readOnly && !props.reviewMode,
    zoom: props.zoom,
    onChange: props.onChange,
    onGeometryChange: scheduleReferences,
    subscribeBlur: props.subscribeWindowBlur,
  });
  const controller = useMemo<GridController & { focus(): void }>(
    () => ({
      select(range) {
        if (selecting.current) return;
        selecting.current = true;
        try {
          const sheet = worksheet();
          sheet?.updateSelectionFromCoords?.(range.x1, range.y1, range.x2, range.y2);
          if (sheet) scrollGridCellIntoView(sheet, range.x1, range.y1);
        } finally {
          selecting.current = false;
        }
      },
      getAnchor(range) {
        return (
          worksheet()?.getCellFromCoords?.(range.x1, range.y1)?.getBoundingClientRect() ?? null
        );
      },
      releaseKeyboard() {
        if (commentClick.current) clearTimeout(commentClick.current);
        commentClick.current = null;
        releaseJSpreadsheetKeyboardOwner();
      },
      focus() {
        const sheet = worksheet();
        if (!sheet) return;
        focusJSpreadsheetWorksheet(sheet);
        surface.current?.focus({ preventScroll: true });
        scheduleReferences();
      },
    }),
    [],
  );
  useEffect(() => {
    props.onController(controller);
    return () => props.onController(null);
  }, [controller]);
  const config = useMemo(() => {
    function capture(value: unknown, allowReview = false) {
      if (
        suppress.current ||
        latest.current.readOnly ||
        (latest.current.reviewMode && !allowReview)
      )
        return;
      const worksheet = value as JSpreadsheetWorksheetInstance,
        data = worksheet?.getData?.();
      if (!data) return;
      const current = latest.current.worksheet;
      const normalized = data.map((row) =>
        row.map((cell) => (cell === undefined ? null : (cell as CellInput))),
      );
      const merge = worksheet.getMerge?.();
      const next: Worksheet = {
        ...current,
        data: normalized.map((row) => [...row]),
        ...(merge && !Array.isArray(merge) ? { mergeCells: structuredClone(merge) } : {}),
      };
      // Native data edits must never capture zoomed presentation measures as authored sizes.
      if (JSON.stringify(next) !== JSON.stringify(current)) latest.current.onChange(next);
    }
    function propose(values: CellInput[][], x: number, y: number) {
      if (latest.current.onReviewChange) {
        latest.current.onReviewChange(values, x, y);
        return;
      }
      const current = latest.current.worksheet,
        workbook = createWorkbook();
      workbook.worksheets = [structuredClone(current)];
      latest.current.onChange(
        writeCells(workbook, current.id, Number(x), Number(y), values).worksheets[0],
      );
    }
    return {
      about: false,
      toolbar: false,
      tabs: false,
      allowExport: false,
      parseFormulas: false,
      contextMenu: () => false,
      onbeforechange: (
        _worksheet: unknown,
        _cell: unknown,
        x: number,
        y: number,
        value: CellInput,
      ) => {
        if (latest.current.readOnly)
          return latest.current.worksheet.data[Number(y)]?.[Number(x)] ?? '';
        const locked = selectionLock(
          { schemaVersion: 1, worksheets: [latest.current.worksheet] },
          latest.current.worksheet.id,
          { x1: Number(x), x2: Number(x), y1: Number(y), y2: Number(y) },
          latest.current.collaborators ?? [],
        );
        if (locked && !suppress.current) {
          latest.current.onBlockedAction?.(`${locked.name} is editing this cell.`);
          return latest.current.worksheet.data[Number(y)]?.[Number(x)] ?? '';
        }

        if (!latest.current.reviewMode || suppress.current) return undefined;
        propose([[value]], Number(x), Number(y));
        return latest.current.worksheet.data[Number(y)]?.[Number(x)] ?? '';
      },
      onbeforepaste: (_worksheet: unknown, data: unknown, x: number, y: number) => {
        if (latest.current.readOnly) return false;
        if (!latest.current.reviewMode && !latest.current.onPasteCells) return undefined;
        const values =
          typeof data === 'string'
            ? parseDelimited(data)
            : Array.isArray(data)
              ? data.map((row: unknown) =>
                  Array.isArray(row)
                    ? row.map((cell: unknown) =>
                        cell && typeof cell === 'object' && 'value' in cell
                          ? (cell.value as CellInput)
                          : (cell as CellInput),
                      )
                    : [],
                )
              : [];
        if (latest.current.onPasteCells) latest.current.onPasteCells(values, Number(x), Number(y));
        else propose(values, x, y);
        return false;
      },
      worksheets: [
        {
          worksheetName: props.worksheet.name,
          data: props.worksheet.data.map((row) => [...row]),
          columns: props.worksheet.columns?.map((column) => ({ ...column, type: 'text' })),
          mergeCells: structuredClone(props.worksheet.mergeCells ?? {}),
          style: { ...props.worksheet.style },
          minDimensions: props.worksheet.minDimensions ?? [26, 100],
          columnResize: false,
          rowResize: false,
          rows: structuredClone(props.worksheet.rows ?? {}),
          freezeColumns: props.worksheet.freezeColumns ?? 0,
          editable: !props.readOnly,
          // React commands move every affected formula and review anchor together.
          allowInsertRow: false,
          allowDeleteRow: false,
          allowInsertColumn: false,
          allowDeleteColumn: false,
          tableOverflow: true,
          tableWidth: '100%',
          tableHeight: '400px',
        },
      ],
      onafterchanges: (worksheet: unknown) => capture(worksheet),
      onresizecolumn: (worksheet: unknown) => capture(worksheet, true),
      onresizerow: (worksheet: unknown) => capture(worksheet, true),
      // Canonical React history includes cell data, formatting and review state.
      onundo: () => false,
      onredo: () => false,
    };
  }, []);
  useLayoutEffect(() => {
    const sheet = worksheet();
    if (!handle.current || !sheet) return;
    if (sheet.options) sheet.options.editable = !props.readOnly;
    if (props.readOnly && inCell.current) sheet.closeEditor?.(undefined, false);
    if (inCell.current || props.pickingFormula) return;
    suppress.current = true;
    try {
      if (applied.current?.data !== props.worksheet.data)
        handle.current.setData(props.worksheet.data.map((row) => [...row]));
      const viewport = gridViewport(sheet);
      const scroll = { left: viewport?.scrollLeft ?? 0, top: viewport?.scrollTop ?? 0 };
      const scale = gridZoomScale(props.zoom),
        previousScale = gridZoomScale(appliedZoom.current);
      const zoomOnly = applied.current === props.worksheet && scale !== previousScale;
      syncWorksheetPresentation(
        sheet,
        applied.current,
        props.worksheet,
        !!props.readOnly,
        props.zoom,
      );
      applied.current = props.worksheet;
      appliedZoom.current = props.zoom ?? 100;
      paint(handle.current, props);
      if (scale !== previousScale) restoreGridScroll(viewport, scroll, previousScale, scale);
      if (!zoomOnly) controller.select(props.selection);
      scheduleReferences();
    } finally {
      suppress.current = false;
    }
  }, [props.worksheet, props.readOnly, props.pickingFormula, props.zoom, revision]);
  useLayoutEffect(() => {
    if (handle.current) controller.select(props.selection);
  }, [props.selection, controller]);
  function paint(api: JSpreadsheetHostHandle, current: WorkbookGridProps) {
    renderCalculatedCells(
      api,
      current.values,
      current.display,
      current.worksheet.data,
      current.unavailableText,
      current.worksheet.formats,
      current.showFormulas,
    );
    renderReviewMarkers(api, current.worksheet.id, current.comments, current.suggestions);
    renderContributors(
      api,
      current.worksheet,
      current.showContributors ?? false,
      current.pickingFormula ? current.selection : undefined,
    );
  }
  useEffect(() => {
    if (handle.current) paint(handle.current, props);
  }, [
    props.values,
    props.display,
    props.worksheet,
    props.unavailableText,
    props.comments,
    props.suggestions,
    props.showFormulas,
    props.showContributors,
    revision,
  ]);
  return (
    <div
      ref={surface}
      tabIndex={-1}
      onKeyDownCapture={(event) => {
        if (inCell.current) native.onKey(event);
        if (
          !inCell.current &&
          selectedHeader.current &&
          ['Delete', 'Backspace'].includes(event.key)
        ) {
          event.preventDefault();
          event.stopPropagation();
          latest.current.onHeaderDelete?.(selectedHeader.current);
          return;
        }
        if (resize.active && event.key === 'Escape') {
          event.preventDefault();
          event.stopPropagation();
          resize.cancel();
        }
      }}
      onDoubleClickCapture={(event) => {
        if (commentClick.current) clearTimeout(commentClick.current);
        const cell = pickCell(event);
        if (!cell) return;
        const locked = selectionLock(
          { schemaVersion: 1, worksheets: [props.worksheet] },
          props.worksheet.id,
          { x1: cell.x, x2: cell.x, y1: cell.y, y2: cell.y },
          props.collaborators ?? [],
        );
        if (locked) {
          event.preventDefault();
          event.stopPropagation();
          props.onBlockedAction?.(`${locked.name} is editing this cell.`);
        }
      }}
      className="se-editor__grid-surface"
      style={{ position: 'relative', minWidth: 0 }}
      onClick={(event) => {
        if (commentClick.current) clearTimeout(commentClick.current);
        if (inCell.current || picking.current || event.detail !== 1 || !event.isTrusted) return;
        const coords = pickCell(event),
          cell = coords && worksheet()?.getCellFromCoords?.(coords.x, coords.y);
        if (!coords || (!cell?.dataset.seComments && !cell?.dataset.seSuggestions)) return;
        commentClick.current = setTimeout(() => {
          const current = latest.current;
          if (
            !inCell.current &&
            current.selection.x1 === coords.x &&
            current.selection.y1 === coords.y
          )
            current.onCommentRequest(coords);
        }, 220);
      }}
      onContextMenuCapture={(event) => {
        const header = deliveredHeader(event, surface.current);
        if (!header) return;
        event.preventDefault();
        event.stopPropagation();
        latest.current.onContextMenu(
          header.axis === 'row' ? { x: -1, y: header.index } : { x: header.index, y: -1 },
          header.element.getBoundingClientRect(),
        );
      }}
      onScrollCapture={scheduleReferences}
      onInputCapture={(event) => {
        if (inCell.current) native.onInput(event);
      }}
      onPointerDownCapture={(event) => {
        if (resize.start(event)) return;
        const header = deliveredHeader(event, surface.current);
        selectedHeader.current = header?.axis ?? null;
        if (header) event.currentTarget.focus({ preventScroll: true });
        if (!latest.current.pickingFormula && !inCell.current && pickCell(event))
          event.currentTarget.focus({ preventScroll: true });
        if (!latest.current.pickingFormula || event.button !== 0 || !event.isTrusted) return;
        const cell = pickCell(event);
        if (!cell) return;
        event.preventDefault();
        event.stopPropagation();
        picking.current = { x1: cell.x, y1: cell.y, x2: cell.x, y2: cell.y };
        event.currentTarget.setPointerCapture(event.pointerId);
      }}
      onPointerMove={(event) => {
        if (resize.move(event)) return;
        const base = surface.current?.getBoundingClientRect();
        if (base) {
          const x = event.clientX - base.left - Number(clip.left ?? 0),
            y = event.clientY - base.top - Number(clip.top ?? 0);
          const hovered = frames.find(
            (frame) =>
              frame.kind === 'selection' &&
              x >= Number(frame.style.left) &&
              x <= Number(frame.style.left) + Number(frame.style.width) &&
              y >= Number(frame.style.top) &&
              y <= Number(frame.style.top) + Number(frame.style.height),
          );
          setHoveredPeer(hovered?.key ?? null);
        }
        if (!picking.current) return;
        const cell = pickCell(event);
        if (cell) picking.current = { ...picking.current, x2: cell.x, y2: cell.y };
      }}
      onPointerUp={(event) => {
        if (resize.finish(event)) return;
        if (!picking.current) return;
        event.preventDefault();
        event.stopPropagation();
        if (inCell.current) native.insert(selectionAddress(picking.current));
        else latest.current.onFormulaReference?.(selectionAddress(picking.current));
        picking.current = null;
        event.currentTarget.releasePointerCapture(event.pointerId);
      }}
      onPointerLeave={() => setHoveredPeer(null)}
      onPointerCancel={() => {
        resize.cancel();
        picking.current = null;
      }}
    >
      <JSpreadsheetHost
        ref={handle}
        className="se-editor__engine"
        config={config}
        onMounted={(_instance, api) => {
          handle.current = api;
          const sheet = worksheet();
          stopBorderCopies.current?.();
          stopBorderCopies.current = sheet ? observeGridBorderCopies(sheet) : null;
          if (sheet)
            syncWorksheetPresentation(
              sheet,
              null,
              latest.current.worksheet,
              !!latest.current.readOnly,
              latest.current.zoom,
            );
          appliedZoom.current = latest.current.zoom ?? 100;
          applied.current = latest.current.worksheet;
          paint(api, latest.current);
          controller.select(latest.current.selection);
          scheduleReferences();
        }}
        onSelectionChange={(range) => {
          // CE selects merge anchors while constructing or changing merges. Those
          // presentation callbacks must not replace the host's current selection.
          if (applied.current && !suppress.current && !selecting.current)
            props.onSelectionChange(range);
        }}
        onContextMenuRequest={props.onContextMenu}
        onCellCommentRequest={props.onCommentRequest}
        onCellEditStart={() => {
          inCell.current = true;
          native.start(worksheet());
          latest.current.onEditStateChange?.(true);
        }}
        onCellEditEnd={() => {
          inCell.current = false;
          native.end();
          latest.current.onEditStateChange?.(false);
          latest.current.onEditDraft?.(null);
          surface.current?.focus({ preventScroll: true });
          setRevision((value) => value + 1);
        }}
      />
      {native.popup}
      {native.overlay}
      {frames.length > 0 && (
        <div className="se-editor__reference-frame" aria-hidden="true" style={clip}>
          {frames.map((frame) => (
            <div
              className={
                frame.local
                  ? 'se-editor__selection-outline'
                  : frame.kind
                    ? 'se-editor__presence-outline'
                    : 'se-editor__reference-outline'
              }
              data-kind={frame.kind}
              data-color={frame.colorIndex}
              key={frame.key}
              style={frame.style}
            >
              {frame.name && frame.key === hoveredPeer && <span>{frame.name}</span>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
