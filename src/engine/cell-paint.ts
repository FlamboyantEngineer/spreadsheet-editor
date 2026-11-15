import type { ReactNode } from 'react';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { DisplayOptions } from '../calculation/format.js';
import type { CellValue } from '../calculation/values.js';
import type {
  CellPresentation,
  CellPresentationSnapshot,
  RenderCellContent,
} from '../editor/public-host.js';
import type { CellFormat } from '../model/review.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { CellInput } from '../model/workbook.js';

export interface GridDraft {
  x: number;
  y: number;
  text: string;
}
export interface CellPortal {
  key: string;
  target: HTMLElement;
  content: ReactNode;
}
export interface GridCellPresentation {
  sheetId: string;
  publication?: { inputKey: string; revision: number };
  readyCells?: ReadonlySet<string>;
  cellPresentation?: CellPresentationSnapshot;
  renderCellContent?: RenderCellContent;
  draft?: GridDraft;
}
/** What one paint of a grid draws from, shared by every cell it paints. */
export interface PaintFrame {
  /** A new token means every cell is painted again. */
  settings: object;
  values: CellValue[][] | undefined;
  raw: CellInput[][];
  display: DisplayOptions | undefined;
  unavailableText: string;
  formats: Record<string, CellFormat>;
  /** Whether any cell has an authored format, so cells without one skip the lookup. */
  hasFormats: boolean;
  showFormulas: boolean;
  presentation: GridCellPresentation | undefined;
  /** Host presentation for the accepted publication, by calculation cell key. */
  hints: Readonly<Record<string, CellPresentation>> | undefined;
  /** The start of every cell key of the sheet, as calculationCellKey writes it. */
  keyPrefix: string;
  /** The host content painted so far, in paint order. */
  portals: CellPortal[];
}
/** What a cell was last painted from, so an unchanged cell is not painted again. */
interface PaintedCell {
  settings: object;
  value: CellValue;
  raw: CellInput | undefined;
  ready: boolean;
  drafting: boolean;
  hint: CellPresentation | undefined;
  format: CellFormat | undefined;
  /** The cell's first child after painting; the engine replaces it when it rewrites the cell. */
  node: ChildNode | null;
  content: ReactNode | undefined;
}
/** One cell of a paint: where it is and what it shows. */
interface CellState {
  x: number;
  y: number;
  key: string;
  input: CellInput | undefined;
  value: CellValue;
  drafting: boolean;
  ready: boolean;
  hint: CellPresentation | undefined;
  format: CellFormat | undefined;
}

const portalTargets = new WeakMap<HTMLElement, HTMLElement>();
const presentationTitles = new WeakMap<HTMLElement, string>();
const paintedCells = new WeakMap<HTMLElement, PaintedCell>();

const formula = (input: unknown): input is string =>
  typeof input === 'string' && input.trimStart().startsWith('=');

export function clearCellPresentation(cell: HTMLElement) {
  if (cell.dataset.sePresentation !== undefined) {
    if (cell.title === presentationTitles.get(cell)) cell.removeAttribute('title');
    presentationTitles.delete(cell);
    cell.removeAttribute('aria-description');
    delete cell.dataset.sePresentation;
    delete cell.dataset.seDecoration;
    delete cell.dataset.sePresentationStatus;
  }
}

/** Forgets what a cell was painted from, so its next paint draws it again. */
export function forgetCell(cell: HTMLElement) {
  paintedCells.delete(cell);
}

/** Removes everything this renderer painted on a cell; authored styles remain untouched. */
export function releaseCell(cell: HTMLElement) {
  clearCellPresentation(cell);
  paintedCells.delete(cell);
  const target = portalTargets.get(cell);
  if (target) target.remove();
  portalTargets.delete(cell);
}

/** A draft shows its own text, without calculated presentation. */
export function paintDraftCell(cell: HTMLElement, text: string) {
  if (cell.textContent !== text) cell.textContent = text;
  cell.dataset.approximate = 'false';
  clearCellPresentation(cell);
  portalTargets.delete(cell);
  paintedCells.delete(cell);
}

/** Calculated values compare by content: a new calculation result holds new objects. */
function sameCellValue(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false;
  const keys = Object.keys(a) as (keyof typeof a)[];
  if (keys.length !== Object.keys(b).length) return false;
  return keys.every((key) => {
    const left: unknown = a[key],
      right: unknown = (b as typeof a)[key];
    return (
      left === right ||
      (typeof left === 'object' &&
        typeof right === 'object' &&
        JSON.stringify(left) === JSON.stringify(right))
    );
  });
}

/** Without calculated values, a formula waits and a number shows as its input. */
function inputValue(input: CellInput | undefined, unavailableText: string): CellValue {
  if (formula(input)) return unavailableText;
  return typeof input === 'number' ? String(input) : (input as CellValue);
}

function cellState(frame: PaintFrame, { x, y }: { x: number; y: number }): CellState {
  const input = frame.raw[y]?.[x];
  const value = frame.values ? frame.values[y][x] : inputValue(input, frame.unavailableText);
  // Equal to calculationCellKey(sheetId, x, y), without a JSON.stringify per cell.
  const key = `${frame.keyPrefix}${x},${y}]`;
  const draft = frame.presentation?.draft;
  const drafting = draft?.x === x && draft.y === y;
  const readyCells = frame.presentation?.readyCells;
  const ready = !drafting && (!readyCells || readyCells.has(key));
  const hint = ready ? frame.hints?.[key] : undefined;
  const format = frame.hasFormats ? frame.formats[formatSpreadsheetAddress(x, y)] : undefined;
  return { x, y, key, input, value, drafting, ready, hint, format };
}

/** Whether a cell still shows what it was painted from, in the same engine node. */
function unchanged(
  painted: PaintedCell,
  state: CellState,
  { settings, node }: { settings: object; node: ChildNode | null },
) {
  return (
    painted.settings === settings &&
    !state.drafting &&
    !painted.drafting &&
    painted.ready === state.ready &&
    painted.raw === state.input &&
    painted.hint === state.hint &&
    painted.format === state.format &&
    painted.node === node &&
    sameCellValue(painted.value, state.value)
  );
}

function cellText(frame: PaintFrame, state: CellState, calculatedText: string): string {
  const { input } = state;
  if (state.drafting) return frame.presentation?.draft?.text ?? '';
  if (!state.ready && input !== null && input !== undefined && input !== '')
    return frame.unavailableText;
  if (frame.showFormulas && formula(input)) return input;
  return state.hint?.text ?? calculatedText;
}

function applyHint(cell: HTMLElement, hint: CellPresentation | undefined) {
  if (!hint) return;
  cell.dataset.sePresentation = '';
  if (hint.tooltip) {
    cell.title = hint.tooltip;
    presentationTitles.set(cell, hint.tooltip);
  }
  if (hint.ariaDescription) cell.setAttribute('aria-description', hint.ariaDescription);
  if (hint.decoration) cell.dataset.seDecoration = hint.decoration;
  if (hint.status) cell.dataset.sePresentationStatus = hint.status;
}

/** The host's content for a ready cell of the accepted publication. */
function hostContent(frame: PaintFrame, state: CellState, calculatedText: string): ReactNode {
  const presentation = frame.presentation,
    accepted = presentation?.publication;
  if (!accepted || !state.ready || !presentation?.renderCellContent) return undefined;
  return presentation.renderCellContent({
    cellKey: state.key,
    sheetId: presentation.sheetId,
    x: state.x,
    y: state.y,
    inputKey: accepted.inputKey,
    revision: accepted.revision,
    rawInput: state.input ?? null,
    value: state.value,
    presentation: state.hint,
    calculatedText,
  });
}

/** Shows host content through a portal target, or else the cell's text. */
function placeContent(
  frame: PaintFrame,
  cell: HTMLElement,
  { key, content, text }: { key: string; content: ReactNode; text: string },
): boolean {
  if (content !== undefined && !frame.showFormulas) {
    let target = portalTargets.get(cell);
    if (!target || target.parentElement !== cell) {
      target = cell.ownerDocument.createElement('span');
      target.className = 'se-editor__cell-content';
      cell.replaceChildren(target);
      portalTargets.set(cell, target);
    }
    frame.portals.push({ key, target, content });
    return true;
  }
  const target = portalTargets.get(cell);
  if (target || cell.textContent !== text) cell.textContent = text;
  portalTargets.delete(cell);
  return false;
}

function markApproximate(cell: HTMLElement, value: CellValue) {
  const approximate = String(
    typeof value === 'object' && value !== null && value.kind === 'number' && value.approximate,
  );
  if (cell.dataset.approximate !== approximate) cell.dataset.approximate = approximate;
}

/** Paints one cell, unless it still shows what it was painted from. */
export function paintCell(frame: PaintFrame, cell: HTMLElement, at: { x: number; y: number }) {
  const state = cellState(frame, at);
  const painted = paintedCells.get(cell);
  if (painted && unchanged(painted, state, { settings: frame.settings, node: cell.firstChild })) {
    const target = portalTargets.get(cell);
    if (painted.content !== undefined && target)
      frame.portals.push({ key: state.key, target, content: painted.content });
    return;
  }
  clearCellPresentation(cell);
  const format: CellFormat | undefined =
    state.hint?.defaultUnit && !state.format
      ? { kind: 'unit', unit: state.hint.defaultUnit }
      : state.format;
  const calculatedText = formatWorksheetCell(state.value, format, frame.display);
  const text = cellText(frame, state, calculatedText);
  applyHint(cell, state.hint);
  const content = hostContent(frame, state, calculatedText);
  const portal = placeContent(frame, cell, { key: state.key, content, text });
  markApproximate(cell, state.value);
  paintedCells.set(cell, {
    settings: frame.settings,
    value: state.value,
    raw: state.input,
    ready: state.ready,
    drafting: state.drafting,
    hint: state.hint,
    format: state.format,
    node: cell.firstChild,
    content: portal ? content : undefined,
  });
}
