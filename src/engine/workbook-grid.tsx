'use client';

import { type CSSProperties, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { createPortal, flushSync } from 'react-dom';
import type {
  CellPresentationSnapshot,
  FormulaFunctionDescriptor,
  RenderCellContent,
} from '../editor/public-host.js';
import { syncWorksheetPresentation } from './worksheet-presentation.js';
import {
  captureWorksheetCellChanges,
  gridWorksheetExtent,
  syncWorksheetData,
} from './worksheet-data.js';
import { createSelectionPublication } from './selection-publication.js';
import { observeGridBorderCopies } from './shared-grid-borders.js';
import {
  gridViewport,
  gridBodyBounds,
  clipGridSelectionHandle,
  gridZoomScale,
  measureGridSelection,
  restoreGridScroll,
  scrollGridCellIntoView,
  syncFrozenColumns,
  syncFrozenColumnOffsets,
} from './grid-geometry.js';
import { useNativeFormula } from './native-formula.js';
import { useHeaderResize, deliveredHeader } from './header-resize.js';
import type { CellValue } from '../calculation/values.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import type { WorkbookIdFactory, WorksheetExtent } from '../model/extent.js';
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
import {
  renderCalculatedCells,
  renderReviewMarkers,
  renderContributors,
  clearCalculatedPresentation,
  type CellPortal,
  type GridDraft,
} from './cell-renderer.js';

export interface WorkbookGridProps {
  worksheet: Worksheet;
  logicalExtent?: WorksheetExtent;
  idFactory?: WorkbookIdFactory;
  values?: CellValue[][];
  draft?: GridDraft;
  readyCells?: ReadonlySet<string>;
  publication?: { inputKey: string; revision: number };
  cellPresentation?: CellPresentationSnapshot;
  renderCellContent?: RenderCellContent;
  formulaFunctions?: readonly FormulaFunctionDescriptor[];
  includeBuiltinFormulaFunctions?: boolean;
  initialViewport?: { left: number; top: number };
  focusOnMount?: boolean;
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
  onError?(error: unknown): void;
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
  onNativeDraft?(draft: GridDraft | null): void;
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
  const focusOnMount = useRef(props.focusOnMount);
  const stopBorderCopies = useRef<(() => void) | null>(null);
  const surface = useRef<HTMLDivElement | null>(null);
  const picking = useRef<CellSelection | null>(null);
  const inCell = useRef(false);
  const nativeComposing = useRef(false);
  const nativeCoords = useRef<{ x: number; y: number } | null>(null);
  const handledNativeChange = useRef(false);
  const commentClick = useRef<ReturnType<typeof setTimeout> | null>(null);
  useEffect(
    () => () => {
      if (commentClick.current) clearTimeout(commentClick.current);
      stopBorderCopies.current?.();
      if (handle.current) clearCalculatedPresentation(handle.current);
    },
    [],
  );
  const selectedHeader = useRef<'row' | 'column' | null>(null);
  const applied = useRef<Worksheet | null>(null);
  const appliedExtent = useRef<[number, number] | null>(null);
  const appliedZoom = useRef(100);
  const native = useNativeFormula(
    (value) => {
      latest.current.onEditDraft?.(value);
      latest.current.onNativeDraft?.(
        value === null || !nativeCoords.current ? null : { ...nativeCoords.current, text: value },
      );
    },
    surface,
    props.formulaFunctions,
    props.includeBuiltinFormulaFunctions,
  );
  const [portals, setPortals] = useState<CellPortal[]>([]);
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
  const focusRequest = useRef(0);
  const latest = useRef(props),
    suppress = useRef(false),
    selecting = useRef(false);
  const visualSelection = useRef(props.selection);
  const draggingSelection = useRef(false);
  const scrolled = useRef(false);
  const selectionPublication = useMemo(
    () => createSelectionPublication((range) => latest.current.onSelectionChange(range)),
    [],
  );
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
    const viewport = gridViewport(sheet) ?? root;
    if (scrolled.current) {
      syncFrozenColumnOffsets(sheet);
      scrolled.current = false;
    }
    const base = root.getBoundingClientRect(),
      body = gridBodyBounds(sheet, viewport);
    const offset = { left: body.left - body.origin.left, top: body.top - body.origin.top };
    const selectionStyles = (range: CellSelection) =>
      measureGridSelection(sheet, current.worksheet, range, viewport).map((style) => ({
        ...style,
        left: Number(style.left ?? 0) - offset.left,
        top: Number(style.top ?? 0) - offset.top,
      }));
    setClip({
      position: 'absolute',
      pointerEvents: 'none',
      overflow: 'hidden',
      boxSizing: 'border-box',
      zIndex: 5,
      left: body.left - base.left - root.clientLeft,
      top: body.top - base.top - root.clientTop,
      width: body.right - body.left,
      height: body.bottom - body.top,
    });
    const measured: typeof frames = selectionStyles(visualSelection.current).map(
      (style, index) => ({ key: `local:${index}`, local: true, style }),
    );
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
      selectionStyles({ x1, y1, x2, y2 }).forEach((style, index) =>
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
        selectionStyles(range).forEach((style, index) =>
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
    clipGridSelectionHandle(sheet, body);
  }
  function scheduleReferences() {
    if (frameRequest.current) return;
    frameRequest.current = requestAnimationFrame(() => {
      frameRequest.current = 0;
      // The engine highlight already changed. Commit its matching outline before
      // this paint, without waiting for parent selection/summary publication.
      flushSync(measureReferences);
    });
  }
  useLayoutEffect(() => {
    scheduleReferences();
  }, [props.formulaReferences, props.collaborators, props.zoom, revision]);
  useEffect(() => {
    scheduleReferences();
    const observer = new ResizeObserver(scheduleReferences);
    if (surface.current) observer.observe(surface.current);
    return () => {
      observer.disconnect();
      cancelAnimationFrame(frameRequest.current);
      cancelAnimationFrame(focusRequest.current);
    };
  }, []);
  useEffect(() => {
    const finish = () => {
      draggingSelection.current = false;
      selectionPublication.finish();
      scheduleReferences();
    };
    const stop = latest.current.subscribeWindowBlur?.(finish);
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
  }, [selectionPublication]);
  const resize = useHeaderResize({
    root: surface,
    worksheet: props.worksheet,
    getWorksheet: worksheet,
    canResize: !props.readOnly && !props.reviewMode,
    zoom: props.zoom,
    onChange: props.onChange,
    onGeometryChange: () => {
      const sheet = worksheet();
      if (sheet) syncFrozenColumns(sheet);
      scheduleReferences();
    },
    subscribeBlur: props.subscribeWindowBlur,
  });
  function selectGridRange(range: CellSelection, reveal: boolean) {
    if (selecting.current) return;
    visualSelection.current = range;
    selecting.current = true;
    try {
      const sheet = worksheet();
      sheet?.updateSelectionFromCoords?.(range.x1, range.y1, range.x2, range.y2);
      if (sheet && reveal) scrollGridCellIntoView(sheet, range.x1, range.y1);
    } finally {
      selecting.current = false;
    }
    scheduleReferences();
  }
  const controller = useMemo<GridController & { focus(): void }>(
    () => ({
      select(range) {
        selectGridRange(range, true);
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
      commitDraft() {
        const sheet = worksheet();
        const cell = (sheet as { edition?: [HTMLTableCellElement] } | null)?.edition?.[0];
        if (inCell.current && cell) sheet?.closeEditor?.(cell, true);
      },
      cancelDraft() {
        const sheet = worksheet();
        const cell = (sheet as { edition?: [HTMLTableCellElement] } | null)?.edition?.[0];
        if (inCell.current && cell) sheet?.closeEditor?.(cell, false);
      },
      getViewport() {
        const sheet = worksheet();
        const viewport = sheet && gridViewport(sheet);
        return { left: viewport?.scrollLeft ?? 0, top: viewport?.scrollTop ?? 0 };
      },
      restoreViewport(position) {
        const sheet = worksheet();
        if (sheet) restoreGridScroll(gridViewport(sheet), position, 1, 1);
      },
    }),
    [],
  );
  useEffect(() => {
    props.onController(controller);
    return () => props.onController(null);
  }, [controller]);
  const config = useMemo(() => {
    function capture(
      value: unknown,
      allowReview = false,
      changes?: readonly { x: number; y: number; value: CellInput | undefined }[],
    ) {
      if (handledNativeChange.current) {
        handledNativeChange.current = false;
        return;
      }
      if (
        suppress.current ||
        latest.current.readOnly ||
        (latest.current.reviewMode && !allowReview && !latest.current.onReviewChange)
      )
        return;
      const worksheet = value as JSpreadsheetWorksheetInstance;
      const current = latest.current.worksheet;
      try {
        const data = changes ? undefined : worksheet?.getData?.();
        if (!changes && !data) return;
        const merge = changes ? undefined : worksheet.getMerge?.();
        const next: Worksheet = {
          ...(changes
            ? captureWorksheetCellChanges(current, changes, latest.current.idFactory)
            : {
                ...current,
                data: data!.map((row) =>
                  row.map((cell) => (cell === undefined ? null : (cell as CellInput))),
                ),
              }),
          ...(merge && !Array.isArray(merge) ? { mergeCells: structuredClone(merge) } : {}),
        };
        // Native data edits must never capture zoomed presentation measures as authored sizes.
        if (JSON.stringify(next) !== JSON.stringify(current)) latest.current.onChange(next);
      } catch (error) {
        // A batch has already changed CE's presentation. Reconcile it with the
        // accepted workbook after host preflight refuses the complete command.
        if (latest.current.onError) latest.current.onError(error);
        else
          latest.current.onBlockedAction?.(
            error instanceof Error ? error.message : 'The cell changes could not be applied.',
          );
        applied.current = null;
        setRevision((value) => value + 1);
      }
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
        writeCells(workbook, current.id, Number(x), Number(y), values, latest.current.idFactory)
          .worksheets[0],
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

        if (suppress.current) return undefined;
        const authored = latest.current.worksheet.data[Number(y)]?.[Number(x)];
        if (inCell.current && String(authored ?? '') === String(value ?? '')) {
          // Text inputs stringify typed primitives. Finishing an unchanged edit
          // must preserve its logical/number type and any canonical input value.
          handledNativeChange.current = true;
          return authored ?? '';
        }
        if (!inCell.current || (!latest.current.reviewMode && !latest.current.onReviewChange))
          return undefined;
        // Host preflight runs before CE mutates the cell or discards its editor.
        // A synchronous rejection therefore retains the exact active native draft.
        propose([[value]], Number(x), Number(y));
        handledNativeChange.current = true;
        return latest.current.reviewMode
          ? (latest.current.worksheet.data[Number(y)]?.[Number(x)] ?? '')
          : undefined;
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
          minDimensions: gridWorksheetExtent(props.worksheet, props.logicalExtent),
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
      onafterchanges: (
        worksheet: unknown,
        changes: { x: number; y: number; value: CellInput | undefined }[],
      ) => capture(worksheet, false, changes),
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
    if (props.readOnly && inCell.current) controller.cancelDraft?.();
    if (inCell.current || props.pickingFormula) return;
    suppress.current = true;
    try {
      const nextExtent = gridWorksheetExtent(props.worksheet, props.logicalExtent);
      const oldExtent = appliedExtent.current;
      if (
        applied.current?.data !== props.worksheet.data ||
        !oldExtent ||
        nextExtent.some((value, axis) => value !== oldExtent[axis])
      )
        syncWorksheetData(sheet, props.worksheet, props.logicalExtent);
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
        props.logicalExtent,
      );
      applied.current = props.worksheet;
      appliedExtent.current = nextExtent;
      appliedZoom.current = props.zoom ?? 100;
      paint(handle.current, props);
      if (scale !== previousScale) restoreGridScroll(viewport, scroll, previousScale, scale);
      if (!zoomOnly && !draggingSelection.current) selectGridRange(props.selection, false);
      scheduleReferences();
    } finally {
      suppress.current = false;
    }
  }, [
    props.worksheet,
    props.readOnly,
    props.pickingFormula,
    props.zoom,
    props.logicalExtent?.columns,
    props.logicalExtent?.rows,
    revision,
  ]);
  useLayoutEffect(() => {
    const current = visualSelection.current,
      next = props.selection;
    if (
      handle.current &&
      !draggingSelection.current &&
      (current.x1 !== next.x1 ||
        current.y1 !== next.y1 ||
        current.x2 !== next.x2 ||
        current.y2 !== next.y2)
    )
      controller.select(next);
  }, [props.selection, controller]);
  function paint(api: JSpreadsheetHostHandle, current: WorkbookGridProps) {
    renderReviewMarkers(api, current.worksheet.id, current.comments, current.suggestions);
    setPortals(
      renderCalculatedCells(
        api,
        current.values,
        current.display,
        current.worksheet.data,
        current.unavailableText,
        current.worksheet.formats,
        current.showFormulas,
        {
          sheetId: current.worksheet.id,
          draft: current.draft,
          readyCells: current.readyCells,
          publication: current.publication,
          cellPresentation: current.cellPresentation,
          renderCellContent: current.renderCellContent,
        },
      ),
    );
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
    props.draft,
    props.readyCells,
    props.publication,
    props.cellPresentation,
    props.renderCellContent,
    revision,
  ]);
  return (
    <div
      ref={surface}
      tabIndex={-1}
      onKeyDownCapture={(event) => {
        if (
          inCell.current &&
          (nativeComposing.current ||
            event.nativeEvent.isComposing ||
            event.nativeEvent.keyCode === 229)
        ) {
          // CE's document listener treats every Enter as a commit. Leave the
          // browser's input method in control without forwarding the key to CE.
          event.stopPropagation();
          return;
        }
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
        if (header) {
          event.preventDefault();
          event.stopPropagation();
          latest.current.onContextMenu(
            header.axis === 'row' ? { x: -1, y: header.index } : { x: header.index, y: -1 },
            header.element.getBoundingClientRect(),
          );
          return;
        }
        // Own cell context menus before the vendor or browser handles the same event.
        // An active native text editor retains its text-editing context menu.
        if (
          event.target instanceof Element &&
          event.target.closest('input, textarea, [contenteditable="true"]')
        ) {
          event.stopPropagation();
          return;
        }
        const coords = pickCell(event);
        const cell = coords && worksheet()?.getCellFromCoords?.(coords.x, coords.y);
        if (!coords || !cell) return;
        event.preventDefault();
        event.stopPropagation();
        latest.current.onContextMenu(coords, cell.getBoundingClientRect());
      }}
      onScrollCapture={(event) => {
        const sheet = worksheet();
        if (sheet && event.target === gridViewport(sheet)) scrolled.current = true;
        scheduleReferences();
      }}
      onInputCapture={(event) => {
        if (inCell.current) native.onInput(event);
      }}
      onCompositionStartCapture={() => {
        if (inCell.current) nativeComposing.current = true;
      }}
      onCompositionEndCapture={() => {
        nativeComposing.current = false;
      }}
      onPasteCapture={(event) => {
        if (inCell.current || (!latest.current.onPasteCells && !latest.current.reviewMode)) return;
        const values = parseDelimited(event.clipboardData.getData('text/plain'));
        event.preventDefault();
        event.stopPropagation();
        if (latest.current.readOnly) return;
        const { x1, y1 } = latest.current.selection;
        // Capture the original clipboard block before CE expands it over the
        // selected rectangle. Both native and menu paths share the host command.
        if (latest.current.onPasteCells) latest.current.onPasteCells(values, x1, y1);
        else latest.current.onReviewChange?.(values, x1, y1);
      }}
      onPointerDownCapture={(event) => {
        if (resize.start(event)) return;
        if (event.button === 0 && !latest.current.pickingFormula && !inCell.current) {
          draggingSelection.current = true;
          selectionPublication.start();
        }
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
        draggingSelection.current = false;
        selectionPublication.finish();
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
        draggingSelection.current = false;
        selectionPublication.finish();
        resize.cancel();
        picking.current = null;
      }}
      onLostPointerCapture={() => {
        draggingSelection.current = false;
        selectionPublication.finish();
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
              latest.current.logicalExtent,
            );
          appliedZoom.current = latest.current.zoom ?? 100;
          applied.current = latest.current.worksheet;
          appliedExtent.current = gridWorksheetExtent(
            latest.current.worksheet,
            latest.current.logicalExtent,
          );
          paint(api, latest.current);
          controller.select(latest.current.selection);
          if (latest.current.initialViewport)
            controller.restoreViewport?.(latest.current.initialViewport);
          if (focusOnMount.current) {
            // A lazy grid mount must not interrupt a press or draft in the chrome.
            focusRequest.current = requestAnimationFrame(() => {
              const active = document.activeElement;
              const editor = surface.current?.closest('.se-editor');
              if (
                active === document.body ||
                (active && surface.current?.contains(active)) ||
                (active?.getAttribute('role') === 'tab' && active.closest('.se-editor') === editor)
              )
                controller.focus();
            });
          }
          scheduleReferences();
        }}
        onBeforeDestroy={() => {
          selectionPublication.dispose();
          releaseJSpreadsheetKeyboardOwner(worksheet());
          native.end();
          if (handle.current) clearCalculatedPresentation(handle.current);
        }}
        onSelectionChange={(range) => {
          // CE selects merge anchors while constructing or changing merges. Those
          // presentation callbacks must not replace the host's current selection.
          if (applied.current && !suppress.current && !selecting.current) {
            visualSelection.current = range;
            scheduleReferences();
            selectionPublication.update(range);
          }
        }}
        onContextMenuRequest={props.onContextMenu}
        onCellCommentRequest={props.onCommentRequest}
        onCellEditStart={(y, x) => {
          inCell.current = true;
          nativeCoords.current = { x, y };
          native.start(worksheet());
          latest.current.onEditStateChange?.(true);
        }}
        onCellEditEnd={() => {
          inCell.current = false;
          nativeComposing.current = false;
          native.end();
          nativeCoords.current = null;
          latest.current.onEditStateChange?.(false);
          surface.current?.focus({ preventScroll: true });
          setRevision((value) => value + 1);
        }}
      />
      {native.popup}
      {native.overlay}
      {portals.map((portal) => createPortal(portal.content, portal.target, portal.key))}
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
