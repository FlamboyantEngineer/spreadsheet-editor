import type { CellSelection } from '../model/editing.js';
import type { SpreadsheetComment, SpreadsheetSuggestion } from '../model/review.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import type { Worksheet } from '../model/workbook.js';
import { worksheetFrom } from './cell-renderer.js';
import { collaboratorColorIndex } from './collaborator-color.js';
import type {
  JSpreadsheetHostHandle,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';

type MarkedWorksheet = JSpreadsheetWorksheetInstance &
  Required<Pick<JSpreadsheetWorksheetInstance, 'getCellFromCoords'>>;
/** The review items a grid marks. */
interface ReviewMarks {
  comments?: SpreadsheetComment[] | undefined;
  suggestions?: SpreadsheetSuggestion[] | undefined;
}
/** Whether contributor colours show, and the cell being edited, which is skipped. */
interface ContributorPaint {
  visible: boolean;
  editing?: CellSelection | undefined;
}

/** Cells marked for each grid, and what they were marked from. */
const reviewMarkers = new WeakMap<
  JSpreadsheetHostHandle,
  { inputs: readonly unknown[]; cells: HTMLElement[] }
>();

const markedCell = (cell: HTMLElement) =>
  cell.isConnected &&
  (cell.dataset.seComments !== undefined || cell.dataset.seSuggestions !== undefined);

function unmark(cells: HTMLElement[]) {
  for (const cell of cells) {
    delete cell.dataset.seComments;
    delete cell.dataset.seSuggestions;
    cell.removeAttribute('title');
  }
}

const markedThread = (thread: SpreadsheetComment, sheetId: string) =>
  thread.sheetId === sheetId &&
  thread.state !== 'resolved' &&
  !thread.orphaned &&
  !thread.suggestion;

/** Counts an open thread on every cell of its range that the engine has. */
function markThread(worksheet: MarkedWorksheet, range: string, marked: Set<HTMLElement>) {
  const bounds = parseSpreadsheetRange(range);
  if (!bounds) return;
  const lastRow = Math.min(bounds.end.y, (worksheet.records?.length ?? 0) - 1);
  for (let y = bounds.start.y; y <= lastRow; y++) {
    const lastColumn = Math.min(bounds.end.x, (worksheet.records?.[y]?.length ?? 0) - 1);
    for (let x = bounds.start.x; x <= lastColumn; x++) {
      const cell = worksheet.getCellFromCoords(x, y);
      if (!cell) continue;
      cell.dataset.seComments = String(Number(cell.dataset.seComments ?? 0) + 1);
      cell.title = `${cell.dataset.seComments} open comment(s)`;
      marked.add(cell);
    }
  }
}

/** Counts a pending suggestion on each cell it changes; a structural one marks A1. */
function markSuggestion(
  worksheet: MarkedWorksheet,
  { suggestion, sheetId }: { suggestion: SpreadsheetSuggestion; sheetId: string },
  marked: Set<HTMLElement>,
) {
  const addresses = suggestion.changes
    .filter((change) => change.sheetId === sheetId)
    .map((change) => change.address);
  if (suggestion.structure) addresses.push('A1');
  for (const address of addresses) {
    const range = parseSpreadsheetRange(address),
      cell = range && worksheet.getCellFromCoords(range.start.x, range.start.y);
    if (!cell) continue;
    cell.dataset.seSuggestions = String(Number(cell.dataset.seSuggestions ?? 0) + 1);
    cell.title = [cell.title, `Suggestion: ${suggestion.summary}`].filter(Boolean).join('; ');
    marked.add(cell);
  }
}

/** Marks every open thread and pending suggestion of the sheet, returning the marked cells. */
function markReview(
  worksheet: MarkedWorksheet,
  sheetId: string,
  { comments, suggestions }: Required<ReviewMarks>,
): Set<HTMLElement> {
  const marked = new Set<HTMLElement>();
  for (const thread of comments ?? [])
    if (markedThread(thread, sheetId)) markThread(worksheet, thread.range ?? '', marked);
  for (const suggestion of suggestions ?? [])
    if (suggestion.status === 'pending' && !suggestion.outdated)
      markSuggestion(worksheet, { suggestion, sheetId }, marked);
  return marked;
}

const hasCellLookup = (
  worksheet: JSpreadsheetWorksheetInstance | undefined,
): worksheet is MarkedWorksheet => !!worksheet?.getCellFromCoords;

/** Mark cells with open comments and pending suggestions. Returns false when nothing
 * changed: the same inputs on the same cells are left as they are.
 */
export function renderReviewMarkers(
  handle: JSpreadsheetHostHandle,
  sheetId: string,
  marks: ReviewMarks,
): boolean {
  const worksheet = worksheetFrom(handle);
  if (!hasCellLookup(worksheet)) return false;
  const comments = marks.comments ?? [],
    suggestions = marks.suggestions ?? [];
  const inputs = [sheetId, comments, suggestions, worksheet.records?.length];
  const previous = reviewMarkers.get(handle);
  if (
    previous?.inputs.every((input, index) => Object.is(input, inputs[index])) &&
    previous.cells.every(markedCell)
  )
    return false;
  unmark(previous?.cells ?? []);
  const marked = markReview(worksheet, sheetId, { comments, suggestions });
  reviewMarkers.set(handle, { inputs, cells: [...marked] });
  return (previous?.cells.length ?? 0) > 0 || marked.size > 0;
}

/** What each grid's contributor colours were last painted from. */
const paintedContributors = new WeakMap<JSpreadsheetHostHandle, readonly unknown[]>();

type Records = NonNullable<JSpreadsheetWorksheetInstance['records']>;

function contributorInputs(
  worksheet: JSpreadsheetWorksheetInstance,
  sheet: Worksheet,
  { visible, editing }: ContributorPaint,
): unknown[] {
  const records = worksheet.records ?? [];
  return [
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
}

function paintContributor(
  cell: HTMLElement,
  { sheet, x, y }: { sheet: Worksheet; x: number; y: number },
  { visible, editing }: ContributorPaint,
) {
  delete cell.dataset.seContributor;
  delete cell.dataset.seAi;
  if (
    !visible ||
    cell.classList.contains('editor') ||
    (editing && x === editing.x1 && y === editing.y1)
  )
    return;
  const entry = sheet.cellProvenance?.[formatSpreadsheetAddress(x, y)];
  cell.dataset.seContributor = entry ? String(collaboratorColorIndex(entry.actor.id)) : 'unknown';
  if (entry?.source === 'ai') cell.dataset.seAi = 'true';
}

function paintContributors(records: Records, sheet: Worksheet, paint: ContributorPaint) {
  for (let y = 0; y < records.length; y++) {
    const row = records[y] ?? [];
    for (let x = 0; x < row.length; x++) {
      const cell = row[x]?.element;
      if (cell instanceof HTMLElement) paintContributor(cell, { sheet, x, y }, paint);
    }
  }
}

/** Contributor colours are presentation attributes, never authored cell styles. Repaints
 * only when the provenance, the edited cell, the visibility or the engine's cells change.
 */
export function renderContributors(
  handle: JSpreadsheetHostHandle,
  sheet: Worksheet,
  paint: ContributorPaint,
): void {
  const worksheet = worksheetFrom(handle);
  if (!worksheet?.records) return;
  const inputs = contributorInputs(worksheet, sheet, paint);
  const previous = paintedContributors.get(handle);
  // Hidden twice in a row: the cells carry no contributor attributes to clear.
  if (previous && !paint.visible && !previous[0]) return;
  if (previous?.every((input, index) => Object.is(input, inputs[index]))) return;
  paintedContributors.set(handle, inputs);
  paintContributors(worksheet.records, sheet, paint);
}
