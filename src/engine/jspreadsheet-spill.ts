import type { JSpreadsheetInstance, JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { GRID_SPACER_CLASS } from './row-window.js';
import {
  resolveSpreadsheetRowSpillWidths,
  type SpreadsheetSpillCell,
} from './spreadsheet-cell-spill.js';

/**
 * How far a text cell may paint past its right edge, in whole pixels; a
 * generated rule per width sets the cell's clip margin. It is presentation
 * only: the vendor style map, and so the saved workbook, never sees it.
 */
const JSPREADSHEET_CELL_SPILL_ATTRIBUTE = 'data-spreadsheet-spill';

function listJSpreadsheetWorksheets(
  instance: JSpreadsheetInstance | null,
): JSpreadsheetWorksheetInstance[] {
  if (!instance) return [];
  if (Array.isArray(instance)) return instance;
  if ('worksheets' in instance && Array.isArray(instance.worksheets)) return instance.worksheets;
  return [instance as JSpreadsheetWorksheetInstance];
}

function readJSpreadsheetSpillColumnWidth(colElement: HTMLElement | undefined): number {
  if (!colElement || colElement.style.display === 'none') return 0;
  const width = Number.parseFloat(colElement.style.width || colElement.getAttribute('width') || '');
  return Number.isFinite(width) ? width : 0;
}

/**
 * Classes that leave an empty cell open to spill: a read-only grid's marker and
 * a comment, whose corner is a background image. The adapter stylesheet takes
 * the same cells out of the positioned layer.
 */
const JSPREADSHEET_SPILL_TRANSPARENT_CLASSES = new Set(['readonly', 'jss_comments']);

function describeJSpreadsheetSpillCell(
  cell: HTMLTableCellElement | undefined,
): SpreadsheetSpillCell {
  if (!cell || cell.style.display === 'none') {
    return { rendered: false, empty: true, decorated: false, spillable: false };
  }
  return {
    rendered: true,
    empty: !cell.firstChild,
    decorated: Array.from(cell.classList).some(
      (name) => !JSPREADSHEET_SPILL_TRANSPARENT_CLASSES.has(name),
    ),
    spillable:
      !cell.firstElementChild &&
      cell.colSpan === 1 &&
      cell.rowSpan === 1 &&
      !cell.classList.contains('editor') &&
      !cell.classList.contains('jss_freezed'),
  };
}

/**
 * One stylesheet rule per spill width, shared by a document's grids, sets the
 * clip margin: typed `attr()` could read the width from the attribute itself,
 * but only Chromium supports it.
 */
const jspreadsheetSpillRuleSheets = new WeakMap<
  Document,
  { sheet: CSSStyleSheet; widths: Set<number> }
>();

function ensureJSpreadsheetSpillRule(ownerDocument: Document, width: number): void {
  let rules = jspreadsheetSpillRuleSheets.get(ownerDocument);
  if (!rules) {
    const StyleSheet = ownerDocument.defaultView?.CSSStyleSheet;
    if (!StyleSheet || !Array.isArray(ownerDocument.adoptedStyleSheets)) return;
    rules = { sheet: new StyleSheet(), widths: new Set<number>() };
    ownerDocument.adoptedStyleSheets = [...ownerDocument.adoptedStyleSheets, rules.sheet];
    jspreadsheetSpillRuleSheets.set(ownerDocument, rules);
  }
  if (rules.widths.has(width)) return;
  rules.widths.add(width);
  rules.sheet.insertRule(
    `[data-jspreadsheet-host] .jss_worksheet > tbody > tr > td[${JSPREADSHEET_CELL_SPILL_ATTRIBUTE}="${width}"] { overflow-clip-margin: ${width}px; }`,
    rules.sheet.cssRules.length,
  );
}

/**
 * Write one cell's spill width, in whole pixels, on jspreadsheet's own `<td>`,
 * read from the worksheet's records; a zero width removes it.
 */
function applyJSpreadsheetCellSpill(
  worksheet: Pick<JSpreadsheetWorksheetInstance, 'records'>,
  { rowIndex, columnIndex }: { rowIndex: number; columnIndex: number },
  spillWidth: number,
): void {
  const cellElement = worksheet.records?.[rowIndex]?.[columnIndex]?.element;
  if (!(cellElement instanceof HTMLElement)) return;
  const width = Math.round(spillWidth);
  const spill = width > 0 ? String(width) : null;
  if (width > 0) ensureJSpreadsheetSpillRule(cellElement.ownerDocument, width);
  if (cellElement.getAttribute(JSPREADSHEET_CELL_SPILL_ATTRIBUTE) === spill) return;
  if (spill === null) cellElement.removeAttribute(JSPREADSHEET_CELL_SPILL_ATTRIBUTE);
  else cellElement.setAttribute(JSPREADSHEET_CELL_SPILL_ATTRIBUTE, spill);
}

/** Rewrites the spill width of each cell in the given rows, or in every row. */
function refreshJSpreadsheetWorksheetSpill(
  worksheet: JSpreadsheetWorksheetInstance,
  rowIndexes?: Iterable<number>,
): void {
  const records = worksheet.records;
  if (!records) return;
  const columnWidths = (worksheet.cols ?? []).map((column) =>
    readJSpreadsheetSpillColumnWidth(column?.colElement),
  );
  for (const rowIndex of rowIndexes ?? records.keys()) {
    const row = records[rowIndex];
    if (!row) continue;
    const spillWidths = resolveSpreadsheetRowSpillWidths(
      row.map((record) => describeJSpreadsheetSpillCell(record?.element)),
      columnWidths,
    );
    spillWidths.forEach((spillWidth, columnIndex) => {
      applyJSpreadsheetCellSpill(worksheet, { rowIndex, columnIndex }, spillWidth);
    });
  }
}

type RowOwner = { worksheet: JSpreadsheetWorksheetInstance; rowIndex: number };
type ChangedRows = Map<JSpreadsheetWorksheetInstance, Set<number>>;
/** A mutation's effect: rows to refresh were added to the map, or every row must refresh. */
type SpillChange = 'rows' | 'all';

const isSpacer = (node: Node) =>
  node instanceof Element && node.classList.contains(GRID_SPACER_CLASS);

function addChangedRow(changedRows: ChangedRows, owner: RowOwner) {
  const rows = changedRows.get(owner.worksheet) ?? new Set<number>();
  rows.add(owner.rowIndex);
  changedRows.set(owner.worksheet, rows);
}

/**
 * The row window attaches and detaches known rows; rows edited while detached were not
 * observed, so a row refreshes when it is attached. New rows refresh every row.
 */
function attachedRows(
  mutation: MutationRecord,
  rowOwners: WeakMap<Node, RowOwner>,
  changedRows: ChangedRows,
): SpillChange {
  for (const node of mutation.addedNodes) {
    const owner = rowOwners.get(node);
    if (owner) addChangedRow(changedRows, owner);
    else if (!isSpacer(node)) return 'all';
  }
  return 'rows';
}

/** Walk from a text node to its cell and row, or stop at a changed column or row list. */
function changedRow(
  mutation: MutationRecord,
  rowOwners: WeakMap<Node, RowOwner>,
  changedRows: ChangedRows,
): SpillChange {
  let node: Node | null = mutation.target;
  for (let depth = 0; node && depth < 3; depth += 1, node = node.parentNode) {
    if (isSpacer(node)) break;
    const owner = rowOwners.get(node);
    if (owner) {
      addChangedRow(changedRows, owner);
      break;
    }
    if (node.nodeName === 'COL' || node.nodeName === 'COLGROUP' || node.nodeName === 'TBODY')
      return 'all';
  }
  return 'rows';
}

function spillChange(
  mutation: MutationRecord,
  rowOwners: WeakMap<Node, RowOwner>,
  changedRows: ChangedRows,
): SpillChange {
  return mutation.type === 'childList' && mutation.target.nodeName === 'TBODY'
    ? attachedRows(mutation, rowOwners, changedRows)
    : changedRow(mutation, rowOwners, changedRows);
}

/**
 * Keeps each text cell's spill width current as the engine renders values,
 * edits, decorations and column widths: a changed cell refreshes its own row,
 * and a changed column, row list or worksheet list refreshes every row.
 */
export function observeJSpreadsheetCellSpill(
  rootEl: HTMLElement,
  instance: JSpreadsheetInstance,
): () => void {
  const MutationObserverClass = rootEl.ownerDocument.defaultView?.MutationObserver;
  if (!MutationObserverClass) return () => {};
  const rowOwners = new WeakMap<Node, RowOwner>();
  let worksheetCount = 0;
  const refreshAll = (): void => {
    const worksheets = listJSpreadsheetWorksheets(instance);
    worksheetCount = worksheets.length;
    worksheets.forEach((worksheet) => {
      worksheet.rows?.forEach((row, rowIndex) => {
        if (row?.element) rowOwners.set(row.element, { worksheet, rowIndex });
      });
      refreshJSpreadsheetWorksheetSpill(worksheet);
    });
  };
  const observer = new MutationObserverClass((mutations) => {
    if (listJSpreadsheetWorksheets(instance).length !== worksheetCount) return refreshAll();
    const changedRows: ChangedRows = new Map();
    for (const mutation of mutations)
      if (spillChange(mutation, rowOwners, changedRows) === 'all') return refreshAll();
    changedRows.forEach((rowIndexes, worksheet) =>
      refreshJSpreadsheetWorksheetSpill(worksheet, rowIndexes),
    );
  });
  observer.observe(rootEl, {
    subtree: true,
    childList: true,
    characterData: true,
    attributes: true,
    attributeFilter: ['class', 'style', 'width', 'colspan', 'rowspan'],
  });
  refreshAll();
  return () => observer.disconnect();
}
