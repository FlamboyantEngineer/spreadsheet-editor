'use client';

import {
  type CSSProperties,
  type Dispatch,
  type RefObject,
  type SetStateAction,
  useRef,
} from 'react';
import type { DisplayOptions } from '../calculation/format.js';
import type { CellValue } from '../calculation/values.js';
import type { GridController } from '../editor/contracts.js';
import type {
  CellPresentationSnapshot,
  FormulaFunctionDescriptor,
  RenderCellContent,
} from '../editor/public-host.js';
import type { CellSelection } from '../model/editing.js';
import type { WorkbookIdFactory, WorksheetExtent } from '../model/extent.js';
import type { FormulaReferenceHighlight } from '../model/formula-reference-parser.js';
import type { SpreadsheetCollaborator } from '../model/presence.js';
import type { SpreadsheetComment, SpreadsheetSuggestion } from '../model/review.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import type { CellPortal, GridDraft } from './cell-renderer.js';
import type { useHeaderResize } from './header-resize.js';
import type {
  JSpreadsheetAnchorRect,
  JSpreadsheetCellCoords,
  JSpreadsheetHostHandle,
  JSpreadsheetSelectionRange,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';
import type { useNativeFormula } from './native-formula.js';
import type { createSelectionPublication } from './selection-publication.js';

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

/** One outline drawn over the grid: the local selection, a formula reference or a peer's range. */
export interface GridFrame {
  key: string;
  style: CSSProperties;
  local?: boolean;
  kind?: 'selection' | 'lock';
  name?: string;
  colorIndex?: number;
}

/** The grid's mutable engine-side state, read by engine callbacks, effects and handlers. */
export interface GridRefs {
  handle: RefObject<JSpreadsheetHostHandle | null>;
  focusOnMount: RefObject<boolean | undefined>;
  stopBorderCopies: RefObject<(() => void) | null>;
  stopRowWindow: RefObject<(() => void) | null>;
  surface: RefObject<HTMLDivElement | null>;
  /** The formula reference being dragged out, while picking one. */
  picking: RefObject<CellSelection | null>;
  /** Whether a cell's native editor is open. */
  inCell: RefObject<boolean>;
  nativeComposing: RefObject<boolean>;
  nativeCoords: RefObject<{ x: number; y: number } | null>;
  handledNativeChange: RefObject<boolean>;
  commentClick: RefObject<ReturnType<typeof setTimeout> | null>;
  selectedHeader: RefObject<'row' | 'column' | null>;
  /** The worksheet, extent and zoom the engine last showed. */
  applied: RefObject<Worksheet | null>;
  appliedExtent: RefObject<[number, number] | null>;
  appliedZoom: RefObject<number>;
  focusRequest: RefObject<number>;
  /** Engine callbacks raised by the grid's own presentation work are ignored. */
  suppress: RefObject<boolean>;
  selecting: RefObject<boolean>;
  visualSelection: RefObject<CellSelection>;
  draggingSelection: RefObject<boolean>;
  scrolled: RefObject<boolean>;
}

/** The refs the component renders with, so React sees them as refs. */
type RenderedRefs = Pick<GridRefs, 'handle' | 'surface'>;

export function useGridRefs(props: WorkbookGridProps, rendered: RenderedRefs): GridRefs {
  return {
    ...rendered,
    focusOnMount: useRef(props.focusOnMount),
    stopBorderCopies: useRef<(() => void) | null>(null),
    stopRowWindow: useRef<(() => void) | null>(null),
    picking: useRef<CellSelection | null>(null),
    inCell: useRef(false),
    nativeComposing: useRef(false),
    nativeCoords: useRef<{ x: number; y: number } | null>(null),
    handledNativeChange: useRef(false),
    commentClick: useRef<ReturnType<typeof setTimeout> | null>(null),
    selectedHeader: useRef<'row' | 'column' | null>(null),
    applied: useRef<Worksheet | null>(null),
    appliedExtent: useRef<[number, number] | null>(null),
    appliedZoom: useRef(100),
    focusRequest: useRef(0),
    suppress: useRef(false),
    selecting: useRef(false),
    visualSelection: useRef(props.selection),
    draggingSelection: useRef(false),
    scrolled: useRef(false),
  };
}

/** Everything the grid's engine callbacks, effects and handlers share. */
export interface GridContext extends GridRefs {
  latest: RefObject<WorkbookGridProps>;
  /** The engine's worksheet, while it is mounted. */
  worksheet(): JSpreadsheetWorksheetInstance | null;
  scheduleReferences(): void;
  setRevision: Dispatch<SetStateAction<number>>;
  setPortals: Dispatch<SetStateAction<CellPortal[]>>;
  selectionPublication: ReturnType<typeof createSelectionPublication>;
}
/** The grid context once the native editor, header resizing and controller exist. */
export interface GridSurfaceContext extends GridContext {
  native: ReturnType<typeof useNativeFormula>;
  resize: ReturnType<typeof useHeaderResize>;
  controller: GridController & { focus(): void };
}

/** The engine's first worksheet. */
export function gridWorksheet(
  handle: JSpreadsheetHostHandle | null,
): JSpreadsheetWorksheetInstance | null {
  const instance = handle?.getInstance();
  if (!instance) return null;
  return Array.isArray(instance)
    ? (instance[0] ?? null)
    : 'worksheets' in instance
      ? (instance.worksheets?.[0] ?? null)
      : (instance as JSpreadsheetWorksheetInstance);
}

type PointerTarget = { target: EventTarget | null; clientX: number; clientY: number };

/** The grid coordinates an element carries, when they are a cell's. */
function cellCoordinates(target: Element): { x: number; y: number } | null {
  if (!(target instanceof HTMLElement)) return null;
  if (target.dataset.x === undefined || target.dataset.y === undefined) return null;
  const x = Number(target.dataset.x),
    y = Number(target.dataset.y);
  return Number.isInteger(x) && Number.isInteger(y) && x >= 0 && y >= 0 ? { x, y } : null;
}

/** The cell under an event; while picking a reference, the cell under the pointer. */
export function pickCell({ surface, picking }: GridRefs, event: PointerTarget) {
  const root = surface.current;
  let target = event.target instanceof Element ? event.target : null;
  if (picking.current)
    target = root?.ownerDocument.elementFromPoint(event.clientX, event.clientY) ?? target;
  if (!root || !target || !root.contains(target)) return null;
  while (target && target !== root) {
    const cell = cellCoordinates(target);
    if (cell) return cell;
    target = target.parentElement;
  }
  return null;
}
