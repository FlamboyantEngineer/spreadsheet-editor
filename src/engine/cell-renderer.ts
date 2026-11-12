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
    const target = portalTargets.get(cell);
    if (target) target.remove();
    portalTargets.delete(cell);
  }
  previousPresentations.delete(handle);
}
/** Only engine-owned cells are painted; raw data and formulas remain untouched. */
export function renderCalculatedCells(
  handle: JSpreadsheetHostHandle,
  values: CellValue[][] | undefined,
  display?: DisplayOptions,
  raw: CellInput[][] = [],
  unavailableText = '…',
  formats: Record<string, CellFormat> = {},
  showFormulas = false,
  presentation?: GridCellPresentation,
): CellPortal[] {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.getCellFromCoords) return [];
  const portals: CellPortal[] = [];
  const touched = new Set<HTMLElement>();
  const previous = previousPresentations.get(handle) ?? new Set<HTMLElement>();
  for (const cell of previous) clearCellPresentation(cell);
  const accepted = presentation?.publication;
  const hints =
    accepted &&
    presentation?.cellPresentation?.inputKey === accepted.inputKey &&
    presentation.cellPresentation.revision === accepted.revision
      ? presentation.cellPresentation.cells
      : undefined;
  const visible: CellValue[][] =
    values ??
    raw.map((row) =>
      row.map((value) =>
        typeof value === 'string' && value.trimStart().startsWith('=')
          ? unavailableText
          : typeof value === 'number'
            ? String(value)
            : value,
      ),
    );
  visible.forEach((row, y) =>
    row.forEach((value, x) => {
      const cell = worksheet.getCellFromCoords?.(x, y);
      if (!cell || cell.classList.contains('editor')) return;
      const key = JSON.stringify([presentation?.sheetId, x, y]);
      const draft = presentation?.draft;
      const drafting = draft?.x === x && draft.y === y;
      const ready = !drafting && (!presentation?.readyCells || presentation.readyCells.has(key));
      const hint = ready ? hints?.[key] : undefined;
      const authoredFormat = formats[formatSpreadsheetAddress(x, y)];
      const format: CellFormat | undefined =
        hint?.defaultUnit && !authoredFormat
          ? { kind: 'unit', unit: hint.defaultUnit }
          : authoredFormat;
      const calculatedText = formatWorksheetCell(value, format, display);
      const text = drafting
        ? draft.text
        : !ready && raw[y]?.[x] !== null && raw[y]?.[x] !== undefined && raw[y]?.[x] !== ''
          ? unavailableText
          : showFormulas &&
              typeof raw[y]?.[x] === 'string' &&
              String(raw[y][x]).trimStart().startsWith('=')
            ? String(raw[y][x])
            : (hint?.text ?? calculatedText);
      touched.add(cell);
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
              rawInput: raw[y]?.[x] ?? null,
              value,
              presentation: hint,
              calculatedText,
            })
          : undefined;
      if (content !== undefined && !showFormulas) {
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
      cell.dataset.approximate = String(
        typeof value === 'object' && value !== null && value.kind === 'number' && value.approximate,
      );
    }),
  );
  if (presentation?.draft) {
    const { x, y, text } = presentation.draft;
    const cell = worksheet.getCellFromCoords(x, y);
    if (cell && !cell.classList.contains('editor')) {
      cell.textContent = text;
      cell.dataset.approximate = 'false';
      clearCellPresentation(cell);
      portalTargets.delete(cell);
      touched.add(cell);
    }
  }
  for (const cell of previous)
    if (!touched.has(cell)) {
      const target = portalTargets.get(cell);
      if (target) target.remove();
      portalTargets.delete(cell);
      if (!cell.classList.contains('editor')) cell.textContent = '';
    }
  previousPresentations.set(handle, touched);
  return portals;
}
export function renderReviewMarkers(
  handle: JSpreadsheetHostHandle,
  sheetId: string,
  comments: SpreadsheetComment[] = [],
  suggestions: SpreadsheetSuggestion[] = [],
): void {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.getCellFromCoords) return;
  const root = worksheet.element ?? worksheet.root;
  root
    ?.querySelectorAll<HTMLElement>('[data-se-comments], [data-se-suggestions]')
    .forEach((cell) => {
      delete cell.dataset.seComments;
      delete cell.dataset.seSuggestions;
      cell.removeAttribute('title');
    });
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
      }
    }
  }
}

/** Contributor colours are presentation attributes, never authored cell styles. */
export function renderContributors(
  handle: JSpreadsheetHostHandle,
  sheet: Worksheet,
  visible: boolean,
  editing?: CellSelection,
): void {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.records) return;
  for (let y = 0; y < worksheet.records.length; y++) {
    const row = worksheet.records[y];
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
