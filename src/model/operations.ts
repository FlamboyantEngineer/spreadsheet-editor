import type { NumericValue } from '../calculation/values.js';
import type { CellInput, CellProvenance, Workbook, Worksheet } from './workbook.js';
import { copyWorkbook } from './workbook.js';
import type { CellFormat } from './review.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import { changeWorkbookStructure, type StructureChange } from './structure.js';
import {
  materializeWorksheetExtent,
  worksheetExtent,
  createWorkbookEntityId,
  type WorkbookIdFactory,
} from './extent.js';
import { writeCells } from './editing.js';

export interface AuthoredCellState {
  value: CellInput;
  style?: string;
  format?: CellFormat;
  inputValue?: NumericValue;
  provenance?: CellProvenance;
}
export interface WorkbookCellChange {
  address: string;
  /** Old coordinates may differ after a line command; stable IDs identify the same cell. */
  beforeAddress?: string;
  rowId?: string;
  columnId?: string;
  before: AuthoredCellState;
  after: AuthoredCellState;
}
export type WorksheetGeometry = Pick<
  Worksheet,
  'minDimensions' | 'columns' | 'rows' | 'freezeColumns'
>;
export type WorkbookOperation =
  | { kind: 'cells'; sheetId: string; changes: WorkbookCellChange[] }
  | {
      kind: 'lines';
      sheetId: string;
      axis: 'row' | 'column';
      action: 'insert' | 'delete';
      index: number;
      count: number;
      ids: string[];
      beforeId?: string;
      afterId?: string;
      /** Removed authored cells and sparse geometry support semantic standalone undo. */
      deletedCells?: WorkbookCellChange[];
      beforeGeometry?: WorksheetGeometry;
      afterGeometry?: WorksheetGeometry;
      beforeMerges?: NonNullable<Worksheet['mergeCells']>;
      afterMerges?: NonNullable<Worksheet['mergeCells']>;
    }
  | {
      kind: 'sheet-add';
      sheet: Worksheet;
      index: number;
      beforeOrder: string[];
      afterOrder: string[];
    }
  | {
      kind: 'sheet-delete';
      sheet: Worksheet;
      index: number;
      beforeOrder: string[];
      afterOrder: string[];
    }
  | { kind: 'sheet-rename'; sheetId: string; before: string; after: string }
  | { kind: 'sheet-reorder'; before: string[]; after: string[] }
  | { kind: 'geometry'; sheetId: string; before: WorksheetGeometry; after: WorksheetGeometry }
  | {
      kind: 'merges';
      sheetId: string;
      before: NonNullable<Worksheet['mergeCells']>;
      after: NonNullable<Worksheet['mergeCells']>;
    }
  | {
      kind: 'annotations';
      before: Pick<Workbook, 'comments' | 'suggestions'>;
      after: Pick<Workbook, 'comments' | 'suggestions'>;
    }
  | {
      kind: 'replace';
      reason: 'import' | 'new-workbook' | 'reset';
      before: Workbook;
      after: Workbook;
    };
export interface WorkbookChangeContext {
  transactionId: string;
  source: 'user' | 'undo' | 'redo' | 'host-command';
  baseline: Workbook;
  operations: WorkbookOperation[];
}
export interface WorkbookCommandResult {
  workbook: Workbook;
  operations: WorkbookOperation[];
}

export function authoredCellState(sheet: Worksheet, address: string): AuthoredCellState {
  const cell = parseSpreadsheetAddress(address);
  if (!cell) throw new TypeError('Invalid cell address.');
  return structuredClone({
    value: sheet.data[cell.y]?.[cell.x] ?? null,
    ...(sheet.style?.[address] !== undefined ? { style: sheet.style[address] } : {}),
    ...(sheet.formats?.[address] !== undefined ? { format: sheet.formats[address] } : {}),
    ...(sheet.inputValues?.[address] !== undefined
      ? { inputValue: sheet.inputValues[address] }
      : {}),
    ...(sheet.cellProvenance?.[address] !== undefined
      ? { provenance: sheet.cellProvenance[address] }
      : {}),
  });
}
function equal(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}
function authoredAddresses(sheet: Worksheet): Set<string> {
  const addresses = new Set<string>();
  sheet.data.forEach((row, y) =>
    row.forEach((_value, x) => addresses.add(formatSpreadsheetAddress(x, y))),
  );
  for (const map of [sheet.style, sheet.formats, sheet.inputValues, sheet.cellProvenance])
    for (const address of Object.keys(map ?? {})) addresses.add(address);
  return addresses;
}
/** Only authored/metadata positions are visited, never an implicit blank matrix. Structure comes from the original command. */
export function collectCellOperations(
  before: Workbook,
  after: Workbook,
  targets?: ReadonlyMap<string, Iterable<string>>,
): WorkbookOperation[] {
  const operations: WorkbookOperation[] = [];
  for (const sheet of before.worksheets) {
    const next = after.worksheets.find((candidate) => candidate.id === sheet.id);
    if (!next) continue;
    const addresses =
      targets?.get(sheet.id) ?? new Set([...authoredAddresses(sheet), ...authoredAddresses(next)]);
    if (targets && !targets.has(sheet.id)) continue;
    const changes: WorkbookCellChange[] = [];
    for (const address of addresses) {
      const previous = authoredCellState(sheet, address),
        current = authoredCellState(next, address);
      if (equal(previous, current)) continue;
      const cell = parseSpreadsheetAddress(address)!;
      changes.push({
        address,
        rowId: sheet.rowIds?.[cell.y],
        columnId: sheet.columnIds?.[cell.x],
        before: previous,
        after: current,
      });
    }
    if (changes.length) operations.push({ kind: 'cells', sheetId: sheet.id, changes });
  }
  return operations;
}
export function writeCellsWithOperations(
  workbook: Workbook,
  sheetId: string,
  x: number,
  y: number,
  values: CellInput[][],
): WorkbookCommandResult {
  const next = writeCells(workbook, sheetId, x, y, values),
    addresses = values.flatMap((row, dy) =>
      row.map((_cell, dx) => formatSpreadsheetAddress(x + dx, y + dy)),
    );
  return {
    workbook: next,
    operations: collectCellOperations(workbook, next, new Map([[sheetId, addresses]])),
  };
}

/** A structural operation is recorded directly, including on identical blank rows. */
export function changeWorkbookStructureWithOperations(
  workbook: Workbook,
  sheetId: string,
  change: StructureChange,
  idFactory: WorkbookIdFactory = createWorkbookEntityId,
): WorkbookCommandResult {
  const materialized = materializeWorksheetExtent(workbook, sheetId),
    sheet = materialized.worksheets.find((sheet) => sheet.id === sheetId)!;
  const extent = worksheetExtent(sheet, materialized);
  if (sheet.rowIds) while (sheet.rowIds.length < extent.rows) sheet.rowIds.push(idFactory('row'));
  if (sheet.columnIds)
    while (sheet.columnIds.length < extent.columns) sheet.columnIds.push(idFactory('column'));
  let oldIds = change.axis === 'row' ? sheet.rowIds : sheet.columnIds;
  const extentSize = change.axis === 'row' ? extent.rows : extent.columns;
  const count = change.kind === 'delete' ? Math.min(change.count, extentSize - 1) : change.count;
  if (!count) return { workbook: copyWorkbook(workbook), operations: [] };
  const ids =
    change.kind === 'insert'
      ? [...(change.ids ?? Array.from({ length: count }, () => idFactory(change.axis)))]
      : (oldIds?.slice(change.index, change.index + count) ?? []);
  if (!oldIds) {
    oldIds = Array.from({ length: extentSize }, () => idFactory(change.axis));
    if (change.axis === 'row') sheet.rowIds = oldIds;
    else sheet.columnIds = oldIds;
    if (change.kind === 'delete') ids.push(...oldIds.slice(change.index, change.index + count));
  }
  const actual = { ...change, count, ...(change.kind === 'insert' ? { ids } : {}) };
  const next = changeWorkbookStructure(materialized, sheetId, actual);
  const operation: WorkbookOperation = {
    kind: 'lines',
    sheetId,
    axis: change.axis,
    action: change.kind,
    index: change.index,
    count,
    ids,
    beforeId: oldIds?.[change.index - 1],
    afterId: oldIds?.[change.kind === 'insert' ? change.index : change.index + count],
    beforeGeometry: (
      geometryOperation(
        sheet,
        next.worksheets.find((candidate) => candidate.id === sheetId)!,
      ) as Extract<WorkbookOperation, { kind: 'geometry' }>
    ).before,
    afterGeometry: (
      geometryOperation(
        sheet,
        next.worksheets.find((candidate) => candidate.id === sheetId)!,
      ) as Extract<WorkbookOperation, { kind: 'geometry' }>
    ).after,
    beforeMerges: structuredClone(sheet.mergeCells ?? {}),
    afterMerges: structuredClone(
      next.worksheets.find((candidate) => candidate.id === sheetId)!.mergeCells ?? {},
    ),
    ...(change.kind === 'delete'
      ? {
          deletedCells: [...authoredAddresses(sheet)].flatMap((address) => {
            const cell = parseSpreadsheetAddress(address)!,
              coordinate = change.axis === 'row' ? cell.y : cell.x;
            return coordinate >= change.index && coordinate < change.index + count
              ? [
                  {
                    address,
                    rowId: sheet.rowIds?.[cell.y],
                    columnId: sheet.columnIds?.[cell.x],
                    before: authoredCellState(sheet, address),
                    after: { value: null },
                  } satisfies WorkbookCellChange,
                ]
              : [];
          }),
        }
      : {}),
  };
  // Compare surviving cells by their structural mapping, so moving a cell is not emitted as a second write.
  const operations: WorkbookOperation[] = [operation];
  for (const previousSheet of workbook.worksheets) {
    const currentSheet = next.worksheets.find((candidate) => candidate.id === previousSheet.id)!;
    const changes: WorkbookCellChange[] = [];
    for (const address of authoredAddresses(previousSheet)) {
      const position = parseSpreadsheetAddress(address)!;
      let coordinate = change.axis === 'row' ? position.y : position.x;
      if (previousSheet.id === sheetId) {
        if (
          change.kind === 'delete' &&
          coordinate >= change.index &&
          coordinate < change.index + count
        )
          continue;
        if (coordinate >= change.index) coordinate += change.kind === 'insert' ? count : -count;
      }
      const destination =
        previousSheet.id === sheetId
          ? formatSpreadsheetAddress(
              change.axis === 'column' ? coordinate : position.x,
              change.axis === 'row' ? coordinate : position.y,
            )
          : address;
      const before = authoredCellState(previousSheet, address),
        after = authoredCellState(currentSheet, destination);
      if (!equal(before, after))
        changes.push({
          address: destination,
          beforeAddress: address,
          rowId: previousSheet.rowIds?.[position.y],
          columnId: previousSheet.columnIds?.[position.x],
          before,
          after,
        });
    }
    if (changes.length) operations.push({ kind: 'cells', sheetId: previousSheet.id, changes });
  }
  if (!equal(workbook.comments, next.comments) || !equal(workbook.suggestions, next.suggestions))
    operations.push(annotationOperation(workbook, next));
  return { workbook: next, operations };
}
export function annotationOperation(before: Workbook, after: Workbook): WorkbookOperation {
  return {
    kind: 'annotations',
    before: structuredClone({ comments: before.comments, suggestions: before.suggestions }),
    after: structuredClone({ comments: after.comments, suggestions: after.suggestions }),
  };
}
export function geometryOperation(before: Worksheet, after: Worksheet): WorkbookOperation {
  const geometry = (sheet: Worksheet): WorksheetGeometry =>
    structuredClone({
      minDimensions: sheet.minDimensions,
      columns: sheet.columns,
      rows: sheet.rows,
      freezeColumns: sheet.freezeColumns,
    });
  return { kind: 'geometry', sheetId: before.id, before: geometry(before), after: geometry(after) };
}
export function mergeOperation(before: Worksheet, after: Worksheet): WorkbookOperation {
  return {
    kind: 'merges',
    sheetId: before.id,
    before: structuredClone(before.mergeCells ?? {}),
    after: structuredClone(after.mergeCells ?? {}),
  };
}
export function sheetOperation(
  kind: 'sheet-add' | 'sheet-delete' | 'sheet-rename' | 'sheet-reorder',
  before: Workbook,
  after: Workbook,
  sheetId?: string,
): WorkbookOperation {
  const beforeOrder = before.worksheets.map((sheet) => sheet.id),
    afterOrder = after.worksheets.map((sheet) => sheet.id);
  if (kind === 'sheet-reorder') return { kind, before: beforeOrder, after: afterOrder };
  const old = before.worksheets.find((sheet) => sheet.id === sheetId),
    next = after.worksheets.find((sheet) => sheet.id === sheetId);
  if (kind === 'sheet-rename') {
    if (!old || !next) throw new RangeError('Unknown worksheet.');
    return { kind, sheetId: old.id, before: old.name, after: next.name };
  }
  const sheet = kind === 'sheet-add' ? next : old;
  if (!sheet) throw new RangeError('Unknown worksheet.');
  return {
    kind,
    sheet: structuredClone(sheet),
    index: (kind === 'sheet-add' ? afterOrder : beforeOrder).indexOf(sheet.id),
    beforeOrder,
    afterOrder,
  };
}

/** Invert recorded command semantics; structure remains explicit even for blank lines. */
export function invertWorkbookOperations(
  operations: readonly WorkbookOperation[],
): WorkbookOperation[] {
  const inverse: WorkbookOperation[] = [];
  for (const operation of [...operations].reverse()) {
    switch (operation.kind) {
      case 'cells':
        inverse.push({
          ...operation,
          changes: operation.changes.map((change) => ({
            ...change,
            address: change.beforeAddress ?? change.address,
            beforeAddress: change.address,
            before: structuredClone(change.after),
            after: structuredClone(change.before),
          })),
        });
        break;
      case 'lines': {
        const line: WorkbookOperation = {
          ...operation,
          action: operation.action === 'insert' ? 'delete' : 'insert',
          beforeGeometry: operation.afterGeometry,
          afterGeometry: operation.beforeGeometry,
          beforeMerges: operation.afterMerges,
          afterMerges: operation.beforeMerges,
        };
        const prior = inverse.splice(0);
        inverse.push(line);
        if (operation.action === 'delete' && operation.deletedCells?.length)
          inverse.push({
            kind: 'cells',
            sheetId: operation.sheetId,
            changes: operation.deletedCells.map((change) => ({
              ...change,
              before: { value: null },
              after: structuredClone(change.before),
            })),
          });
        if (operation.beforeGeometry && operation.afterGeometry)
          inverse.push({
            kind: 'geometry',
            sheetId: operation.sheetId,
            before: structuredClone(operation.afterGeometry),
            after: structuredClone(operation.beforeGeometry),
          });
        if (operation.beforeMerges && operation.afterMerges)
          inverse.push({
            kind: 'merges',
            sheetId: operation.sheetId,
            before: structuredClone(operation.afterMerges),
            after: structuredClone(operation.beforeMerges),
          });
        inverse.push(...prior);
        break;
      }
      case 'sheet-add':
        inverse.push({
          ...operation,
          kind: 'sheet-delete',
          beforeOrder: [...operation.afterOrder],
          afterOrder: [...operation.beforeOrder],
        });
        break;
      case 'sheet-delete':
        inverse.push({
          ...operation,
          kind: 'sheet-add',
          beforeOrder: [...operation.afterOrder],
          afterOrder: [...operation.beforeOrder],
        });
        break;
      default:
        inverse.push({
          ...operation,
          before: structuredClone(operation.after),
          after: structuredClone(operation.before),
        } as WorkbookOperation);
        break;
    }
  }
  return inverse;
}
