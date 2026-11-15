import { createWorkbookEntityId, type WorksheetExtent,worksheetExtent } from './extent.js';
import {
  interval,
  mapReferences,
  moveAddress,
  moveRange,
  rewriteStructureReferences,
} from './formula-references.js';
import { requireDefined } from './require-defined.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
  parseSpreadsheetRange,
} from './spreadsheet-addressing.js';
import {
  assertWorkbook,
  type CellInput,
  copyWorkbook,
  DEFAULT_WORKBOOK_LIMITS,
  type Workbook,
  type Worksheet,
  worksheetNameKey,
} from './workbook.js';

export interface StructureChange {
  axis: 'row' | 'column';
  kind: 'insert' | 'delete';
  index: number;
  count: number;
  /** IDs generated once by the command owner for newly inserted lines. */
  ids?: readonly string[];
}

function checkChange(change: StructureChange) {
  if (
    !Number.isSafeInteger(change.index) ||
    !Number.isSafeInteger(change.count) ||
    change.index < 0 ||
    change.index > 50_000 ||
    change.count < 1 ||
    change.count > 50_000
  )
    throw new RangeError('Invalid structural change.');
}

/** The change within the sheet's extent; a deletion always leaves one line. */
function boundedChange(change: StructureChange, axisExtent: number): StructureChange {
  if (change.index > axisExtent || (change.kind === 'delete' && change.index >= axisExtent))
    throw new RangeError('Structural change is outside the worksheet extent.');
  if (change.kind === 'delete' && change.index + change.count > axisExtent)
    throw new RangeError('Structural deletion is outside the worksheet extent.');
  return {
    ...change,
    count: change.kind === 'delete' ? Math.min(change.count, axisExtent - 1) : change.count,
  };
}

/** Records the sheet's new dimensions, within the cell limit. */
function resize(sheet: Worksheet, extent: WorksheetExtent, change: StructureChange) {
  const delta = change.kind === 'insert' ? change.count : -change.count;
  const columns = extent.columns + (change.axis === 'column' ? delta : 0),
    rows = extent.rows + (change.axis === 'row' ? delta : 0);
  if (columns > 16384 || rows > 1048576 || columns * rows > DEFAULT_WORKBOOK_LIMITS.maxCells)
    throw new RangeError('Worksheet exceeds the cell limit.');
  sheet.minDimensions = [columns, rows];
}

function invalidIds(ids: readonly string[], change: StructureChange, existing?: string[]) {
  return (
    change.kind !== 'insert' ||
    ids.length !== change.count ||
    ids.some((id) => typeof id !== 'string' || !id || id.length > 256) ||
    new Set(ids).size !== ids.length ||
    ids.some((id) => existing?.includes(id))
  );
}

/** Inserts or removes line identities; supplied ids start tracking on a sheet without them. */
function changeLineIds(sheet: Worksheet, change: StructureChange, axisExtent: number) {
  let existing = change.axis === 'row' ? sheet.rowIds : sheet.columnIds;
  if (!existing && change.ids) {
    existing = Array.from({ length: axisExtent }, () => createWorkbookEntityId(change.axis));
    if (change.axis === 'row') sheet.rowIds = existing;
    else sheet.columnIds = existing;
  }
  if (change.ids !== undefined && invalidIds(change.ids, change, existing))
    throw new TypeError('Invalid inserted line identities.');
  if (!existing) return;
  const inserted =
    change.kind === 'insert'
      ? (change.ids ??
        Array.from({ length: change.count }, () => createWorkbookEntityId(change.axis)))
      : [];
  existing.splice(change.index, change.kind === 'delete' ? change.count : 0, ...inserted);
}

/** Inserts blank entries into a list that reaches the change, or removes the deleted ones. */
function spliceLines<T>(list: T[], change: StructureChange, blank: () => T) {
  if (change.kind === 'delete') list.splice(change.index, change.count);
  else if (change.index < list.length)
    list.splice(change.index, 0, ...Array.from({ length: change.count }, blank));
}

function changeLines(sheet: Worksheet, change: StructureChange) {
  if (change.axis === 'row') {
    spliceLines<CellInput[]>(sheet.data, change, () => []);
    if (!sheet.rows) return;
    const rowSettings: NonNullable<Worksheet['rows']> = {};
    for (const [index, settings] of Object.entries(sheet.rows)) {
      const moved = interval(Number(index), Number(index), change);
      if (moved) rowSettings[String(moved[0])] = settings;
    }
    sheet.rows = rowSettings;
    return;
  }
  sheet.data.forEach((row) => spliceLines<CellInput>(row, change, () => null));
  if (sheet.columns) spliceLines(sheet.columns, change, () => ({}));
}

/** Moves every cell-keyed map and merge with the change; deleted cells drop out. */
function moveCellMaps(sheet: Worksheet, change: StructureChange) {
  const remap = <T>(map: Record<string, T> | undefined): Record<string, T> =>
    Object.fromEntries(
      Object.entries(map ?? {}).flatMap(([address, value]) => {
        const next = moveAddress(address, change);
        return next ? [[next, value]] : [];
      }),
    );
  // A map the worksheet did not have stays absent, so an inverse restores it exactly (SE-HOST-5.1).
  if (sheet.style) sheet.style = remap(sheet.style);
  if (sheet.formats) sheet.formats = remap(sheet.formats);
  if (sheet.cellProvenance) sheet.cellProvenance = remap(sheet.cellProvenance);
  if (sheet.inputValues) sheet.inputValues = remap(sheet.inputValues);
  if (!sheet.mergeCells) return;
  const merges: NonNullable<Worksheet['mergeCells']> = {};
  for (const [address, [columns, rows]] of Object.entries(sheet.mergeCells)) {
    const anchor = requireDefined(parseSpreadsheetAddress(address));
    const moved = moveRange(
      `${address}:${formatSpreadsheetAddress(anchor.x + columns - 1, anchor.y + rows - 1)}`,
      change,
    );
    const span = moved && parseSpreadsheetRange(moved);
    if (span)
      merges[formatSpreadsheetAddress(span.start.x, span.start.y)] = [
        span.end.x - span.start.x + 1,
        span.end.y - span.start.y + 1,
      ];
  }
  sheet.mergeCells = merges;
}

/** Rewrites formulas, comment ranges and suggestions across the workbook for the change. */
function followChange(next: Workbook, sheet: Worksheet, change: StructureChange) {
  const rewrite = (formula: string, currentSheet: string) =>
    rewriteStructureReferences(formula, change, { currentSheet, targetSheet: sheet.name });
  next.worksheets.forEach((current) =>
    current.data.forEach((row) =>
      row.forEach((value, index) => {
        if (typeof value === 'string') row[index] = rewrite(value, current.name);
      }),
    ),
  );
  next.comments?.forEach((thread) => {
    if (thread.sheetId !== sheet.id || thread.orphaned) return;
    const moved = moveRange(thread.range ?? '', change);
    if (moved) thread.range = requireDefined(parseSpreadsheetRange(moved)).address;
    else thread.orphaned = true;
  });
  next.suggestions?.forEach((item) =>
    item.changes.forEach((cell) => {
      const current = next.worksheets.find((candidate) => candidate.id === cell.sheetId);
      if (cell.sheetId === sheet.id) {
        const moved = moveAddress(cell.address, change);
        if (moved) cell.address = moved;
        else item.outdated = true;
      }
      if (!current) return;
      for (const field of ['before', 'after'] as const) {
        const value = cell[field];
        if (typeof value === 'string') cell[field] = rewrite(value, current.name);
      }
    }),
  );
}

export function changeWorkbookStructure(
  workbook: Workbook,
  sheetId: string,
  requested: StructureChange,
): Workbook {
  checkChange(requested);
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  const extent = worksheetExtent(sheet, next),
    axisExtent = requested.axis === 'row' ? extent.rows : extent.columns;
  const change = boundedChange(requested, axisExtent);
  if (!change.count) return next;
  resize(sheet, extent, change);
  changeLineIds(sheet, change, axisExtent);
  changeLines(sheet, change);
  if (!sheet.data.length) sheet.data = [['']];
  if (!sheet.data.some((row) => row.length)) sheet.data[0].push(null);
  moveCellMaps(sheet, change);
  followChange(next, sheet, change);
  assertWorkbook(next);
  return next;
}

/** Return the normalized name accepted by worksheet renaming, or throw without changing the workbook. */
export function validateWorksheetName(workbook: Workbook, sheetId: string, name: string): string {
  const normalized = name.trim().replace(/\s+/g, ' ');
  if (
    !workbook.worksheets.some((sheet) => sheet.id === sheetId) ||
    !normalized ||
    normalized.length > 256 ||
    /[[\]:*?/\\!]/.test(normalized) ||
    workbook.worksheets.some(
      (other) =>
        other.id !== sheetId && worksheetNameKey(other.name) === worksheetNameKey(normalized),
    )
  )
    throw new TypeError('Choose a unique worksheet name without reserved characters.');
  return normalized;
}

export function renameWorksheet(workbook: Workbook, sheetId: string, name: string): Workbook {
  const normalized = validateWorksheetName(workbook, sheetId, name),
    next = copyWorkbook(workbook),
    sheet = requireDefined(next.worksheets.find((sheet) => sheet.id === sheetId));
  const previous = worksheetNameKey(sheet.name);
  sheet.name = normalized;
  const rename = (formula: string) =>
    mapReferences(formula, (reference, target, prefix) =>
      target !== null && worksheetNameKey(target) === previous
        ? `'${normalized.replace(/'/g, "''")}'!${reference}`
        : prefix + reference,
    );
  rewriteFormulas(next, rename);
  assertWorkbook(next);
  return next;
}

/** Apply a formula rewrite to every authored cell and to both sides of every suggested change. */
function rewriteFormulas(workbook: Workbook, rewrite: (formula: string) => string): void {
  for (const sheet of workbook.worksheets)
    for (const row of sheet.data)
      row.forEach((value, index) => {
        if (typeof value === 'string') row[index] = rewrite(value);
      });
  const changes = (workbook.suggestions ?? []).flatMap((item) => item.changes);
  for (const cell of changes)
    for (const field of ['before', 'after'] as const)
      if (typeof cell[field] === 'string') cell[field] = rewrite(cell[field] as string);
}

export function deleteWorksheet(workbook: Workbook, sheetId: string): Workbook {
  if (workbook.worksheets.length < 2) throw new Error('Keep at least one worksheet.');
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  next.worksheets = next.worksheets.filter((sheet) => sheet.id !== sheetId);
  const deleted = worksheetNameKey(sheet.name);
  rewriteFormulas(next, (formula) =>
    mapReferences(formula, (reference, target, prefix) =>
      target !== null && worksheetNameKey(target) === deleted ? '#REF!' : prefix + reference,
    ),
  );
  next.comments?.forEach((thread) => {
    if (thread.sheetId === sheetId) thread.orphaned = true;
  });
  next.suggestions?.forEach((item) => {
    if (item.changes.some((cell) => cell.sheetId === sheetId)) item.outdated = true;
  });
  return next;
}
