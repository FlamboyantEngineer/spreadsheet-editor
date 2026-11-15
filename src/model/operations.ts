import type { NumericValue } from '../calculation/values.js';
import { writeCells } from './editing.js';
import {
  createWorkbookEntityId,
  materializeWorksheetExtent,
  type WorkbookIdFactory,
  worksheetExtent,
} from './extent.js';
import { requireDefined } from './require-defined.js';
import type { CellFormat } from './review.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import { changeWorkbookStructure, type StructureChange } from './structure.js';
import {
  type CellInput,
  type CellProvenance,
  copyWorkbook,
  type Workbook,
  type Worksheet,
} from './workbook.js';

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
  | { kind: 'workbook-metadata'; before: Pick<Workbook, 'name'>; after: Pick<Workbook, 'name'> }
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
/** Ordinary writes can append identified lines before their cell targets are introduced. */
export function collectLineGrowthOperations(
  before: Worksheet,
  after: Worksheet,
): WorkbookOperation[] {
  const operations: WorkbookOperation[] = [];
  for (const axis of ['row', 'column'] as const) {
    const previous = axis === 'row' ? before.rowIds : before.columnIds;
    const current = axis === 'row' ? after.rowIds : after.columnIds;
    if (
      !previous ||
      !current ||
      current.length <= previous.length ||
      previous.some((id, index) => current[index] !== id)
    )
      continue;
    operations.push({
      kind: 'lines',
      sheetId: before.id,
      axis,
      action: 'insert',
      index: previous.length,
      count: current.length - previous.length,
      ids: current.slice(previous.length),
      beforeId: previous.at(-1),
    });
  }
  return operations;
}
const METADATA = ['style', 'formats', 'inputValues', 'cellProvenance'] as const;
function sameEntry(a: unknown, b: unknown): boolean {
  return a === b || (typeof a === 'object' && typeof b === 'object' && equal(a, b));
}
/**
 * Addresses whose authored state may differ, found by comparing values in place: data by
 * position, then each metadata map by key. Nothing is copied or serialized for an unchanged
 * cell, so an untargeted diff stays cheap on a large sheet (SE-PERF-1).
 */
function changedAddresses(sheet: Worksheet, next: Worksheet): Set<string> {
  const addresses = new Set<string>();
  if (sheet.data !== next.data)
    for (let y = 0; y < Math.max(sheet.data.length, next.data.length); y++) {
      const before = sheet.data[y],
        after = next.data[y];
      if (before === after) continue;
      for (let x = 0; x < Math.max(before?.length ?? 0, after?.length ?? 0); x++)
        if ((before?.[x] ?? null) !== (after?.[x] ?? null))
          addresses.add(formatSpreadsheetAddress(x, y));
    }
  for (const field of METADATA) {
    const before: Record<string, unknown> = sheet[field] ?? {},
      after: Record<string, unknown> = next[field] ?? {};
    if (before === after) continue;
    for (const address of Object.keys(before))
      if (!sameEntry(before[address], after[address])) addresses.add(address);
    for (const address of Object.keys(after))
      if (!Object.hasOwn(before, address)) addresses.add(address);
  }
  return addresses;
}
/**
 * Cell patches between two workbooks. `targets` limits the comparison to the addresses a command
 * wrote; without it, only authored positions that differ are visited, never an implicit blank
 * matrix. Structure comes from the original command.
 */
export function collectCellOperations(
  before: Workbook,
  after: Workbook,
  targets?: ReadonlyMap<string, Iterable<string>>,
): WorkbookOperation[] {
  const operations: WorkbookOperation[] = [];
  for (const sheet of before.worksheets) {
    if (targets && !targets.has(sheet.id)) continue;
    const next = after.worksheets.find((candidate) => candidate.id === sheet.id);
    if (!next || next === sheet) continue;
    const changes: WorkbookCellChange[] = [];
    for (const address of targets?.get(sheet.id) ?? changedAddresses(sheet, next)) {
      const previous = authoredCellState(sheet, address),
        current = authoredCellState(next, address);
      if (equal(previous, current)) continue;
      const cell = requireDefined(parseSpreadsheetAddress(address));
      changes.push({
        address,
        rowId: sheet.rowIds?.[cell.y] ?? next.rowIds?.[cell.y],
        columnId: sheet.columnIds?.[cell.x] ?? next.columnIds?.[cell.x],
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
  idFactory?: WorkbookIdFactory,
): WorkbookCommandResult {
  const next = writeCells(workbook, sheetId, x, y, values, idFactory),
    addresses = values.flatMap((row, dy) =>
      row.map((_cell, dx) => formatSpreadsheetAddress(x + dx, y + dy)),
    );
  return {
    workbook: next,
    operations: [
      ...collectLineGrowthOperations(
        requireDefined(workbook.worksheets.find((sheet) => sheet.id === sheetId)),
        requireDefined(next.worksheets.find((sheet) => sheet.id === sheetId)),
      ),
      ...collectCellOperations(workbook, next, new Map([[sheetId, addresses]])),
    ],
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
    sheet = requireDefined(materialized.worksheets.find((sheet) => sheet.id === sheetId));
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
  const nextSheet = requireDefined(next.worksheets.find((candidate) => candidate.id === sheetId));
  const geometry = geometryOperation(sheet, nextSheet) as Extract<
    WorkbookOperation,
    { kind: 'geometry' }
  >;
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
    beforeGeometry: geometry.before,
    afterGeometry: geometry.after,
    beforeMerges: structuredClone(sheet.mergeCells ?? {}),
    afterMerges: structuredClone(nextSheet.mergeCells ?? {}),
    ...(change.kind === 'delete'
      ? {
          deletedCells: [...authoredAddresses(sheet)].flatMap((address) => {
            const cell = requireDefined(parseSpreadsheetAddress(address)),
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
    const currentSheet = requireDefined(
      next.worksheets.find((candidate) => candidate.id === previousSheet.id),
    );
    const changes: WorkbookCellChange[] = [];
    for (const address of authoredAddresses(previousSheet)) {
      const position = requireDefined(parseSpreadsheetAddress(address));
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
  const inserted = new Map<string, { row: Set<string>; column: Set<string> }>();
  for (const operation of operations) {
    if (operation.kind !== 'lines' || operation.action !== 'insert') continue;
    let axes = inserted.get(operation.sheetId);
    if (!axes) {
      axes = { row: new Set(), column: new Set() };
      inserted.set(operation.sheetId, axes);
    }
    for (const id of operation.ids) axes[operation.axis].add(id);
  }
  for (const operation of [...operations].reverse()) {
    switch (operation.kind) {
      case 'cells': {
        // Removing inserted lines removes their new cells too; do not write to those identities
        // after the inverse line command has removed them.
        const axes = inserted.get(operation.sheetId);
        const changes = operation.changes
          .filter(
            (change) =>
              !axes?.row.has(change.rowId ?? '') && !axes?.column.has(change.columnId ?? ''),
          )
          .map((change) => ({
            ...change,
            address: change.beforeAddress ?? change.address,
            beforeAddress: change.address,
            before: structuredClone(change.after),
            after: structuredClone(change.before),
          }));
        if (changes.length) inverse.push({ ...operation, changes });
        break;
      }
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
