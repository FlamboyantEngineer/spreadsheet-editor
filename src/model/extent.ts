import { DEFAULT_WORKBOOK_LIMITS } from './limits.js';
import { parseSpreadsheetAddress, parseSpreadsheetRange } from './spreadsheet-addressing.js';
import type { Workbook, Worksheet } from './workbook.js';

export interface WorksheetExtent {
  columns: number;
  rows: number;
}
export type WorkbookEntityKind = 'row' | 'column' | 'sheet' | 'transaction';
export type WorkbookIdFactory = (kind: WorkbookEntityKind) => string;
export const createWorkbookEntityId: WorkbookIdFactory = () => crypto.randomUUID();

/** Grows an extent to reach a cell, or a block anchored at it. */
function include(extent: WorksheetExtent, address: string, [width, height] = [1, 1]) {
  const cell = parseSpreadsheetAddress(address);
  if (!cell) return;
  extent.columns = Math.max(extent.columns, cell.x + width);
  extent.rows = Math.max(extent.rows, cell.y + height);
}

/** The extent of the sheet's data, declared dimensions and row and column metadata. */
function dataExtent(sheet: Worksheet): WorksheetExtent {
  const columns = Math.max(
    1,
    sheet.minDimensions?.[0] ?? 0,
    sheet.columns?.length ?? 0,
    sheet.freezeColumns ?? 0,
    sheet.columnIds?.length ?? 0,
    sheet.data.reduce((widest, row) => Math.max(widest, row.length), 0),
  );
  const rows = Math.max(
    1,
    sheet.minDimensions?.[1] ?? 0,
    sheet.data.length,
    sheet.rowIds?.length ?? 0,
    Object.keys(sheet.rows ?? {}).reduce((last, index) => Math.max(last, Number(index) + 1), 0),
  );
  return { columns, rows };
}

/** Grows an extent to reach the sheet's live comment ranges and current suggestions. */
function includeReview(extent: WorksheetExtent, sheet: Worksheet, workbook: Workbook) {
  for (const thread of workbook.comments ?? []) {
    if (thread.sheetId !== sheet.id || thread.orphaned) continue;
    const range = parseSpreadsheetRange(thread.range ?? '');
    if (!range) continue;
    extent.columns = Math.max(extent.columns, range.end.x + 1);
    extent.rows = Math.max(extent.rows, range.end.y + 1);
  }
  const changes = (workbook.suggestions ?? [])
    .filter((suggestion) => !suggestion.outdated)
    .flatMap((suggestion) => suggestion.changes);
  for (const change of changes) if (change.sheetId === sheet.id) include(extent, change.address);
}

/** Logical geometry includes sparse metadata, not just populated data. */
export function worksheetExtent(sheet: Worksheet, workbook?: Workbook): WorksheetExtent {
  const extent = dataExtent(sheet);
  for (const map of [sheet.style, sheet.formats, sheet.inputValues, sheet.cellProvenance])
    for (const address of Object.keys(map ?? {})) include(extent, address);
  for (const [address, size] of Object.entries(sheet.mergeCells ?? {}))
    include(extent, address, size);
  if (workbook) includeReview(extent, sheet, workbook);
  return extent;
}

/** Mounted geometry, with the initial blank viewport only when it fits the cell budget. */
export function gridWorksheetExtent(
  sheet: Worksheet,
  logicalExtent?: WorksheetExtent,
): [number, number] {
  const extent = worksheetExtent(sheet);
  const columns = Math.max(extent.columns, logicalExtent?.columns ?? 0);
  const rows = Math.max(extent.rows, logicalExtent?.rows ?? 0);
  const padded: [number, number] = [Math.max(columns, 26), Math.max(rows, 100)];
  return sheet.minDimensions || padded[0] * padded[1] > DEFAULT_WORKBOOK_LIMITS.maxCells
    ? [columns, rows]
    : padded;
}

/** Calculation ranges ignore annotation and decoration anchors which only extend the viewport. */
export function calculationFormats(sheet: Worksheet) {
  const formats = Object.fromEntries(
    Object.entries(sheet.formats ?? {})
      .filter(([, format]) => format.kind === 'unit')
      .map(([address, format]) => [
        address,
        {
          kind: format.kind,
          unit: format.unit,
          inputUnit: format.inputUnit,
          expectedDimension: format.expectedDimension,
          quantityKind: format.quantityKind,
        },
      ]),
  );
  return Object.keys(formats).length ? formats : undefined;
}
export function worksheetCalculationExtent(sheet: Worksheet): WorksheetExtent {
  return worksheetExtent({
    ...sheet,
    style: undefined,
    cellProvenance: undefined,
    formats: calculationFormats(sheet),
  });
}

/** Only editor structure commands materialize their implicit initial viewport. */
export function materializeWorksheetExtent(
  workbook: Workbook,
  sheetId: string,
  defaults: readonly [number, number] = [26, 100],
): Workbook {
  const next = structuredClone(workbook),
    sheet = next.worksheets.find((candidate) => candidate.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  const extent = worksheetExtent(sheet, next);
  const padded: [number, number] = [
    Math.max(defaults[0], extent.columns),
    Math.max(defaults[1], extent.rows),
  ];
  // The viewport padding is a convenience: a sheet already near the cell limit keeps its logical extent.
  sheet.minDimensions =
    sheet.minDimensions || padded[0] * padded[1] > DEFAULT_WORKBOOK_LIMITS.maxCells
      ? [extent.columns, extent.rows]
      : padded;
  return next;
}

export function extendWorksheetIdentities(
  sheet: Worksheet,
  factory: WorkbookIdFactory = createWorkbookEntityId,
  workbook?: Workbook,
): void {
  if (!sheet.rowIds && !sheet.columnIds) return;
  const extent = worksheetExtent(sheet, workbook);
  for (const [axis, length] of [
    ['row', extent.rows],
    ['column', extent.columns],
  ] as const) {
    const ids = axis === 'row' ? sheet.rowIds : sheet.columnIds;
    if (!ids) continue;
    while (ids.length < length) ids.push(factory(axis));
  }
}
