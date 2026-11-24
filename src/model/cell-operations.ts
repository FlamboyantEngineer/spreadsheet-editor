import { type CellPlacement, writeCells } from './editing.js';
import { collectCellShape } from './operation-cell-shape.js';
import type {
  AuthoredCellState,
  WorkbookCellChange,
  WorkbookCommandResult,
  WorkbookOperation,
} from './operations.js';
import { requireDefined } from './require-defined.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { CellInput, Workbook, Worksheet } from './workbook.js';

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
export function equal(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}
export function authoredAddresses(sheet: Worksheet, includeBlanks = false): Set<string> {
  const addresses = new Set<string>();
  sheet.data.forEach((row, y) =>
    row.forEach((value, x) => {
      if (includeBlanks || (value !== null && value !== ''))
        addresses.add(formatSpreadsheetAddress(x, y));
    }),
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
      extentOnly: true,
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
/** Adds the addresses of one row's positions whose values differ. */
function changedInRow(
  addresses: Set<string>,
  [before, after]: [CellInput[] | undefined, CellInput[] | undefined],
  y: number,
) {
  if (before === after) return;
  for (let x = 0; x < Math.max(before?.length ?? 0, after?.length ?? 0); x++)
    if ((before?.[x] ?? null) !== (after?.[x] ?? null))
      addresses.add(formatSpreadsheetAddress(x, y));
}
/** Adds the keys whose entries differ between two versions of a metadata map. */
function changedInMap(
  addresses: Set<string>,
  before: Record<string, unknown>,
  after: Record<string, unknown>,
) {
  if (before === after) return;
  for (const address of Object.keys(before))
    if (!sameEntry(before[address], after[address])) addresses.add(address);
  for (const address of Object.keys(after))
    if (!Object.hasOwn(before, address)) addresses.add(address);
}
/**
 * Addresses whose authored state may differ, found by comparing values in place: data by
 * position, then each metadata map by key. Nothing is copied or serialized for an unchanged
 * cell, so an untargeted diff stays cheap on a large sheet (SE-PERF-2.1).
 */
function changedAddresses(sheet: Worksheet, next: Worksheet): Set<string> {
  const addresses = new Set<string>();
  if (sheet.data !== next.data)
    for (let y = 0; y < Math.max(sheet.data.length, next.data.length); y++)
      changedInRow(addresses, [sheet.data[y], next.data[y]], y);
  for (const field of METADATA) changedInMap(addresses, sheet[field] ?? {}, next[field] ?? {});
  return addresses;
}
/** The cell changes between two versions of a sheet at the given addresses. */
function sheetCellChanges(
  sheet: Worksheet,
  next: Worksheet,
  targets: Iterable<string>,
): WorkbookCellChange[] {
  const changes: WorkbookCellChange[] = [];
  for (const address of targets) {
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
  return changes;
}

function sheetOperations(sheet: Worksheet, next: Worksheet, targets?: Iterable<string>) {
  const addresses = targets ? new Set(targets) : changedAddresses(sheet, next);
  const changes = sheetCellChanges(sheet, next, addresses);
  const rows = targets
    ? [
        ...new Set(
          [...addresses].map((address) => requireDefined(parseSpreadsheetAddress(address)).y),
        ),
      ]
    : Array.from({ length: Math.max(sheet.data.length, next.data.length) }, (_, index) => index);
  const shape = collectCellShape(sheet, next, { rows, addresses: targets ? addresses : undefined });
  return changes.length || shape
    ? { kind: 'cells' as const, sheetId: sheet.id, changes, ...(shape && { shape }) }
    : undefined;
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
    const operation = sheetOperations(sheet, next, targets?.get(sheet.id));
    if (operation) operations.push(operation);
  }
  return operations;
}
export function writeCellsWithOperations(
  workbook: Workbook,
  values: CellInput[][],
  placement: CellPlacement,
): WorkbookCommandResult {
  const { sheetId, x, y } = placement;
  const next = writeCells(workbook, values, placement),
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
