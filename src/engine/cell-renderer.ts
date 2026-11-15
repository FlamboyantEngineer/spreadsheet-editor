import type { CellValue } from '../calculation/values.js';
import type { ReactNode } from 'react';
import type { CellPresentationSnapshot, RenderCellContent } from '../editor/public-host.js';
import type { DisplayOptions } from '../calculation/format.js';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import { collaboratorColorIndex } from '../model/presence.js';
import type { CellSelection } from '../model/editing.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import type { CellFormat, SpreadsheetComment, SpreadsheetSuggestion } from '../model/review.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import type { JSpreadsheetHostHandle, JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

function worksheetFrom(handle: JSpreadsheetHostHandle): JSpreadsheetWorksheetInstance | undefined {
  const instance = handle.getInstance();
  if (!instance) return undefined;
  return Array.isArray(instance)
    ? instance[0]
    : 'worksheets' in instance
      ? instance.worksheets?.[0]
      : (instance as JSpreadsheetWorksheetInstance);
}
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
const previousPresentations = new WeakMap<JSpreadsheetHostHandle, Set<HTMLElement>>();
const portalTargets = new WeakMap<HTMLElement, HTMLElement>();
const presentationTitles = new WeakMap<HTMLElement, string>();
/** What a cell was last painted from, so an unchanged cell is not painted again. */
interface PaintedCell {
  settings: object;
  value: CellValue;
  raw: CellInput | undefined;
  ready: boolean;
  drafting: boolean;
  hint: unknown;
  format: CellFormat | undefined;
  /** The cell's first child after painting; the engine replaces it when it rewrites the cell. */
  node: ChildNode | null;
  content: ReactNode | undefined;
}
const paintedCells = new WeakMap<HTMLElement, PaintedCell>();
/**
 * What a row was last painted from. A values row keeps its identity only while no cell in it
 * changes value or readiness, so a row whose values, input, formats, hints and settings are all
 * the same objects is left as it is without visiting its cells (SE-PERF-1).
 */
interface PaintedRow {
  settings: object;
  /** The values row, or the input row when there are no values. */
  values: readonly (CellValue | CellInput)[];
  raw: readonly CellInput[] | undefined;
  formats: Record<string, CellFormat>;
  hints: unknown;
  cells: HTMLElement[];
  /** Each cell's first child after painting, as in PaintedCell. */
  nodes: (ChildNode | null)[];
  portals: CellPortal[];
}
const NO_FORMATS: Record<string, CellFormat> = Object.freeze({});
const paintedRows = new WeakMap<JSpreadsheetHostHandle, Map<number, PaintedRow>>();
/** Settings shared by every cell; a new object means every cell is painted again. */
const paintSettings = new WeakMap<
  JSpreadsheetHostHandle,
  { inputs: readonly unknown[]; token: object }
>();
function clearCellPresentation(cell: HTMLElement) {
  if (cell.dataset.sePresentation !== undefined) {
    if (cell.title === presentationTitles.get(cell)) cell.removeAttribute('title');
    presentationTitles.delete(cell);
    cell.removeAttribute('aria-description');
    delete cell.dataset.sePresentation;
    delete cell.dataset.seDecoration;
    delete cell.dataset.sePresentationStatus;
  }
}

/** Release only presentation owned by this renderer; authored styles remain untouched. */
export function clearCalculatedPresentation(handle: JSpreadsheetHostHandle) {
  for (const cell of previousPresentations.get(handle) ?? []) {
    clearCellPresentation(cell);
    paintedCells.delete(cell);
    const target = portalTargets.get(cell);
    if (target) target.remove();
    portalTargets.delete(cell);
  }
  previousPresentations.delete(handle);
  paintSettings.delete(handle);
  paintedRows.delete(handle);
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

/** The settings token for this paint: unchanged while every shared input is unchanged. */
function settingsToken(
  handle: JSpreadsheetHostHandle,
  inputs: readonly unknown[],
  force: boolean,
): object {
  const previous = paintSettings.get(handle);
  if (
    !force &&
    previous &&
    previous.inputs.every((input, index) => Object.is(input, inputs[index]))
  )
    return previous.token;
  const token = {};
  paintSettings.set(handle, { inputs, token });
  return token;
}

/** Only engine-owned cells are painted; raw data and formulas remain untouched. A cell
 * whose value, input, format, readiness and presentation are unchanged is left as it is.
 * `repaint` paints every cell again, for when other presentation has changed the cells.
 */
export function renderCalculatedCells(
  handle: JSpreadsheetHostHandle,
  values: CellValue[][] | undefined,
  display?: DisplayOptions,
  raw: CellInput[][] = [],
  unavailableText = '…',
  formats: Record<string, CellFormat> = NO_FORMATS,
  showFormulas = false,
  presentation?: GridCellPresentation,
  repaint = false,
): CellPortal[] {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.getCellFromCoords) return [];
  const portals: CellPortal[] = [];
  const touched = new Set<HTMLElement>();
  const previous = previousPresentations.get(handle) ?? new Set<HTMLElement>();
  const accepted = presentation?.publication;
  const hints =
    accepted &&
    presentation?.cellPresentation?.inputKey === accepted.inputKey &&
    presentation.cellPresentation.revision === accepted.revision
      ? presentation.cellPresentation.cells
      : undefined;
  const settings = settingsToken(
    handle,
    [
      display,
      unavailableText,
      showFormulas,
      presentation?.sheetId,
      presentation?.renderCellContent,
      // Host content receives the publication; text alone does not depend on it.
      presentation?.renderCellContent && accepted?.inputKey,
      presentation?.renderCellContent && accepted?.revision,
    ],
    repaint,
  );
  const hasFormats = Object.keys(formats).length > 0;
  const keyPrefix = `[${JSON.stringify(presentation?.sheetId)},`;
  const draft = presentation?.draft;
  const readyCells = presentation?.readyCells;
  const rows = values ?? raw;
  const rowMemo = paintedRows.get(handle) ?? new Map<number, PaintedRow>();
  paintedRows.set(handle, rowMemo);
  if (rowMemo.size > rows.length)
    for (const y of rowMemo.keys()) if (y >= rows.length) rowMemo.delete(y);
  for (let y = 0; y < rows.length; y++) {
    const row = rows[y],
      rawRow = raw[y];
    const memo = rowMemo.get(y);
    if (
      memo &&
      memo.values === row &&
      memo.raw === rawRow &&
      memo.settings === settings &&
      memo.formats === formats &&
      memo.hints === hints &&
      draft?.y !== y &&
      // A rebuilt row has new cells; the engine replaces every cell of a row it rebuilds.
      memo.cells[0] === worksheet.getCellFromCoords(0, y) &&
      // The engine writes its raw value into a cell when an editor closes.
      memo.cells.every((cell, index) => cell.firstChild === memo.nodes[index])
    ) {
      for (const cell of memo.cells) touched.add(cell);
      for (const portal of memo.portals) portals.push(portal);
      continue;
    }
    const rowCells: HTMLElement[] = [],
      rowPortals = portals.length;
    let rowSettled = draft?.y !== y;
    for (let x = 0; x < row.length; x++) {
      const cell = worksheet.getCellFromCoords(x, y);
      if (!cell) continue;
      if (cell.classList.contains('editor')) {
        // The engine rewrites the cell when its editor closes.
        paintedCells.delete(cell);
        rowSettled = false;
        continue;
      }
      rowCells.push(cell);
      const input = rawRow?.[x];
      const value: CellValue = values
        ? values[y][x]
        : typeof input === 'string' && input.trimStart().startsWith('=')
          ? unavailableText
          : typeof input === 'number'
            ? String(input)
            : (input as CellValue);
      // Equal to calculationCellKey(sheetId, x, y), without a JSON.stringify per cell.
      const key = `${keyPrefix}${x},${y}]`;
      const drafting = draft?.x === x && draft.y === y;
      const ready = !drafting && (!readyCells || readyCells.has(key));
      const hint = ready ? hints?.[key] : undefined;
      const authoredFormat = hasFormats ? formats[formatSpreadsheetAddress(x, y)] : undefined;
      touched.add(cell);
      const painted = paintedCells.get(cell);
      if (
        painted &&
        painted.settings === settings &&
        !drafting &&
        !painted.drafting &&
        painted.ready === ready &&
        painted.raw === input &&
        painted.hint === hint &&
        painted.format === authoredFormat &&
        painted.node === cell.firstChild &&
        sameCellValue(painted.value, value)
      ) {
        const target = portalTargets.get(cell);
        if (painted.content !== undefined && target)
          portals.push({ key, target, content: painted.content });
        continue;
      }
      clearCellPresentation(cell);
      const format: CellFormat | undefined =
        hint?.defaultUnit && !authoredFormat
          ? { kind: 'unit', unit: hint.defaultUnit }
          : authoredFormat;
      const calculatedText = formatWorksheetCell(value, format, display);
      const text = drafting
        ? draft.text
        : !ready && input !== null && input !== undefined && input !== ''
          ? unavailableText
          : showFormulas && typeof input === 'string' && input.trimStart().startsWith('=')
            ? input
            : (hint?.text ?? calculatedText);
      if (hint) {
        cell.dataset.sePresentation = '';
        if (hint.tooltip) {
          cell.title = hint.tooltip;
          presentationTitles.set(cell, hint.tooltip);
        }
        if (hint.ariaDescription) cell.setAttribute('aria-description', hint.ariaDescription);
        if (hint.decoration) cell.dataset.seDecoration = hint.decoration;
        if (hint.status) cell.dataset.sePresentationStatus = hint.status;
      }
      const content =
        accepted && ready && presentation?.renderCellContent
          ? presentation.renderCellContent({
              cellKey: key,
              sheetId: presentation.sheetId,
              x,
              y,
              inputKey: accepted.inputKey,
              revision: accepted.revision,
              rawInput: input ?? null,
              value,
              presentation: hint,
              calculatedText,
            })
          : undefined;
      const portal = content !== undefined && !showFormulas;
      if (portal) {
        let target = portalTargets.get(cell);
        if (!target || target.parentElement !== cell) {
          target = cell.ownerDocument.createElement('span');
          target.className = 'se-editor__cell-content';
          cell.replaceChildren(target);
          portalTargets.set(cell, target);
        }
        portals.push({ key, target, content });
      } else {
        const target = portalTargets.get(cell);
        if (target || cell.textContent !== text) cell.textContent = text;
        portalTargets.delete(cell);
      }
      const approximate = String(
        typeof value === 'object' && value !== null && value.kind === 'number' && value.approximate,
      );
      if (cell.dataset.approximate !== approximate) cell.dataset.approximate = approximate;
      paintedCells.set(cell, {
        settings,
        value,
        raw: input,
        ready,
        drafting,
        hint,
        format: authoredFormat,
        node: cell.firstChild,
        content: portal ? content : undefined,
      });
    }
    if (rowSettled)
      rowMemo.set(y, {
        settings,
        values: row,
        raw: rawRow,
        formats,
        hints,
        cells: rowCells,
        nodes: rowCells.map((cell) => cell.firstChild),
        portals: portals.slice(rowPortals),
      });
    else rowMemo.delete(y);
  }
  if (draft) {
    const { x, y, text } = draft;
    const cell = worksheet.getCellFromCoords(x, y);
    if (cell && !cell.classList.contains('editor')) {
      if (cell.textContent !== text) cell.textContent = text;
      cell.dataset.approximate = 'false';
      clearCellPresentation(cell);
      portalTargets.delete(cell);
      paintedCells.delete(cell);
      touched.add(cell);
    }
  }
  for (const cell of previous)
    if (!touched.has(cell)) {
      clearCellPresentation(cell);
      paintedCells.delete(cell);
      const target = portalTargets.get(cell);
      if (target) target.remove();
      portalTargets.delete(cell);
      if (!cell.classList.contains('editor')) cell.textContent = '';
    }
  previousPresentations.set(handle, touched);
  return portals;
}
/** Cells marked for each grid, and what they were marked from. */
const reviewMarkers = new WeakMap<
  JSpreadsheetHostHandle,
  { inputs: readonly unknown[]; cells: HTMLElement[] }
>();

/** Mark cells with open comments and pending suggestions. Returns false when nothing
 * changed: the same inputs on the same cells are left as they are.
 */
export function renderReviewMarkers(
  handle: JSpreadsheetHostHandle,
  sheetId: string,
  comments: SpreadsheetComment[] = [],
  suggestions: SpreadsheetSuggestion[] = [],
): boolean {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.getCellFromCoords) return false;
  const inputs = [sheetId, comments, suggestions, worksheet.records?.length];
  const previous = reviewMarkers.get(handle);
  if (
    previous &&
    previous.inputs.every((input, index) => Object.is(input, inputs[index])) &&
    previous.cells.every(
      (cell) =>
        cell.isConnected &&
        (cell.dataset.seComments !== undefined || cell.dataset.seSuggestions !== undefined),
    )
  )
    return false;
  for (const cell of previous?.cells ?? []) {
    delete cell.dataset.seComments;
    delete cell.dataset.seSuggestions;
    cell.removeAttribute('title');
  }
  const marked = new Set<HTMLElement>();
  for (const thread of comments) {
    if (thread.sheetId !== sheetId || thread.resolved || thread.orphaned || thread.suggestionId)
      continue;
    const range = parseSpreadsheetRange(thread.range ?? '');
    if (!range) continue;
    for (
      let y = range.start.y;
      y <= Math.min(range.end.y, (worksheet.records?.length ?? 0) - 1);
      y++
    ) {
      const width = worksheet.records?.[y]?.length ?? 0;
      for (let x = range.start.x; x <= Math.min(range.end.x, width - 1); x++) {
        const cell = worksheet.getCellFromCoords(x, y);
        if (cell) {
          cell.dataset.seComments = String(Number(cell.dataset.seComments ?? 0) + 1);
          cell.title = `${cell.dataset.seComments} open comment(s)`;
          marked.add(cell);
        }
      }
    }
  }
  for (const suggestion of suggestions) {
    if (suggestion.status !== 'pending' || suggestion.outdated) continue;
    const addresses = suggestion.changes
      .filter((change) => change.sheetId === sheetId)
      .map((change) => change.address);
    if (suggestion.structure) addresses.push('A1');
    for (const address of addresses) {
      const range = parseSpreadsheetRange(address),
        cell = range && worksheet.getCellFromCoords(range.start.x, range.start.y);
      if (cell) {
        cell.dataset.seSuggestions = String(Number(cell.dataset.seSuggestions ?? 0) + 1);
        cell.title = [cell.title, `Suggestion: ${suggestion.summary}`].filter(Boolean).join('; ');
        marked.add(cell);
      }
    }
  }
  reviewMarkers.set(handle, { inputs, cells: [...marked] });
  return (previous?.cells.length ?? 0) > 0 || marked.size > 0;
}

/** What each grid's contributor colours were last painted from. */
const paintedContributors = new WeakMap<JSpreadsheetHostHandle, readonly unknown[]>();

/** Contributor colours are presentation attributes, never authored cell styles. Repaints
 * only when the provenance, the edited cell, the visibility or the engine's cells change.
 */
export function renderContributors(
  handle: JSpreadsheetHostHandle,
  sheet: Worksheet,
  visible: boolean,
  editing?: CellSelection,
): void {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.records) return;
  const records = worksheet.records;
  const inputs = [
    visible,
    visible ? sheet.cellProvenance : undefined,
    visible ? editing?.x1 : undefined,
    visible ? editing?.y1 : undefined,
    // The cell being edited is skipped; the engine rewrites it when its editor closes.
    visible ? (worksheet as { edition?: unknown[] }).edition?.[0] : undefined,
    records.length,
    records[0]?.length,
    records[0]?.[0]?.element,
    records[records.length - 1]?.[0]?.element,
  ];
  const previous = paintedContributors.get(handle);
  // Hidden twice in a row: the cells carry no contributor attributes to clear.
  if (previous && !visible && !previous[0]) return;
  if (previous?.every((input, index) => Object.is(input, inputs[index]))) return;
  paintedContributors.set(handle, inputs);
  for (let y = 0; y < records.length; y++) {
    const row = records[y];
    if (!row) continue;
    for (let x = 0; x < row.length; x++) {
      const cell = row[x]?.element;
      if (!(cell instanceof HTMLElement)) continue;
      delete cell.dataset.seContributor;
      delete cell.dataset.seAi;
      if (
        !visible ||
        cell.classList.contains('editor') ||
        (editing && x === editing.x1 && y === editing.y1)
      )
        continue;
      const entry = sheet.cellProvenance?.[formatSpreadsheetAddress(x, y)];
      cell.dataset.seContributor = entry
        ? String(collaboratorColorIndex(entry.actor.id))
        : 'unknown';
      if (entry?.source === 'ai') cell.dataset.seAi = 'true';
    }
  }
}
