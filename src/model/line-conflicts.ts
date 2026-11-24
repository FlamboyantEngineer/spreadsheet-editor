import { authoredAddresses, authoredCellState } from './cell-operations.js';
import { equal, sheetOf, WorkbookOperationConflict } from './operation-delta.js';
import type { AuthoredCellState, WorkbookChangeContext, WorkbookOperation } from './operations.js';
import { requireDefined } from './require-defined.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { Workbook, Worksheet } from './workbook.js';

type Axis = 'row' | 'column';
type LineOperation = Extract<WorkbookOperation, { kind: 'lines' }>;

interface LineContent {
  cells: Record<string, AuthoredCellState>;
  geometry: unknown;
  merges: unknown[];
}

/** Index targeted lines once, then visit their authored values and merge coverage. */
function lineContents(sheet: Worksheet, { axis, ids }: LineOperation) {
  const wanted = new Set(ids);
  const lines = axis === 'row' ? sheet.rowIds : sheet.columnIds;
  const byId = new Map<string, LineContent>();
  const byPosition = new Map<number, LineContent>();
  lines?.forEach((id, index) => {
    if (!wanted.has(id)) return;
    const geometry = axis === 'row' ? sheet.rows?.[String(index)] : sheet.columns?.[index];
    const content = { cells: {}, geometry, merges: [] };
    byId.set(id, content);
    byPosition.set(index, content);
  });
  collectCells(sheet, axis, byPosition);
  collectMerges(sheet, axis, byPosition);
  return byId;
}

function collectCells(sheet: Worksheet, axis: Axis, lines: ReadonlyMap<number, LineContent>) {
  for (const address of authoredAddresses(sheet)) {
    const { x, y } = requireDefined(parseSpreadsheetAddress(address));
    const content = lines.get(axis === 'row' ? y : x);
    if (!content) continue;
    const row = sheet.rowIds?.[y],
      column = sheet.columnIds?.[x];
    if (!row || !column) throw new WorkbookOperationConflict('line-content-identities');
    content.cells[JSON.stringify([row, column])] = authoredCellState(sheet, address);
  }
}

function collectMerges(sheet: Worksheet, axis: Axis, lines: ReadonlyMap<number, LineContent>) {
  for (const [address, [width, height]] of Object.entries(sheet.mergeCells ?? {})) {
    const { x, y } = requireDefined(parseSpreadsheetAddress(address));
    const begin = axis === 'row' ? y : x;
    const end = begin + (axis === 'row' ? height : width);
    const coverage = [sheet.rowIds?.slice(y, y + height), sheet.columnIds?.slice(x, x + width)];
    for (const [position, content] of lines)
      if (position >= begin && position < end) content.merges.push(coverage);
  }
  for (const content of lines.values())
    content.merges.sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
}

/** A deleted line someone changed since the baseline is rejected without mutation. */
export function checkDeletedLines(current: Workbook, transaction: WorkbookChangeContext) {
  for (const operation of transaction.operations) {
    if (operation.kind !== 'lines' || operation.action !== 'delete') continue;
    const previous = sheetOf(transaction.baseline, operation.sheetId);
    const latest = sheetOf(current, operation.sheetId);
    if (previous === latest || equal(previous, latest)) continue;
    const priorLines = lineContents(previous, operation);
    for (const [id, content] of lineContents(latest, operation))
      if (!equal(priorLines.get(id), content))
        throw new WorkbookOperationConflict(`modified-line/${id}`);
  }
}
