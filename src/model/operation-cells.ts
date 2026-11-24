import { writeAuthoredCellState } from './authored-cell.js';
import { restoreCellShape } from './operation-cell-shape.js';
import { delta, WorkbookOperationConflict } from './operation-delta.js';
import {
  type AuthoredCellState,
  authoredCellState,
  type WorkbookCellChange,
  type WorkbookOperation,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { Worksheet } from './workbook.js';

type AxisIndex = ReadonlyMap<string, number> | undefined;

export function lineIndex(ids: readonly string[] | undefined): AxisIndex {
  return ids && new Map(ids.map((id, index) => [id, index]));
}

/** Where the recorded cell lands now. Stable identities win over its recorded position. */
function changeAddress(
  change: WorkbookCellChange,
  { rows, columns }: { rows: AxisIndex; columns: AxisIndex },
  sheetId: string,
): string {
  const position = requireDefined(parseSpreadsheetAddress(change.address));
  const y = change.rowId ? rows?.get(change.rowId) : position.y;
  const x = change.columnId ? columns?.get(change.columnId) : position.x;
  if (x === undefined || y === undefined)
    throw new WorkbookOperationConflict(`cells/${sheetId}/${change.address}`);
  return formatSpreadsheetAddress(x, y);
}

export function applyCells(
  sheet: Worksheet,
  operation: Extract<WorkbookOperation, { kind: 'cells' }>,
) {
  const indices = { rows: lineIndex(sheet.rowIds), columns: lineIndex(sheet.columnIds) };
  const restore = operation.shape && restoreCellShape(sheet, operation.shape, indices.rows);
  for (const change of operation.changes) {
    const address = changeAddress(change, indices, sheet.id);
    const merged = delta(authoredCellState(sheet, address), change, `cells/${sheet.id}/${address}`);
    writeAuthoredCellState(sheet, address, {
      ...(merged as AuthoredCellState),
      materialize: change.materializeValue,
    });
  }
  restore?.();
}
