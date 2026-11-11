import type { Workbook, Worksheet } from './workbook.js';
import { parseSpreadsheetAddress, parseSpreadsheetRange } from './spreadsheet-addressing.js';

export interface WorksheetExtent {
  columns: number;
  rows: number;
}
export type WorkbookEntityKind = 'row' | 'column' | 'sheet' | 'transaction';
export type WorkbookIdFactory = (kind: WorkbookEntityKind) => string;
export const createWorkbookEntityId: WorkbookIdFactory = () => crypto.randomUUID();

/** Logical geometry includes sparse metadata, not just populated data. */
export function worksheetExtent(sheet: Worksheet, workbook?: Workbook): WorksheetExtent {
  let columns = Math.max(
      1,
      sheet.minDimensions?.[0] ?? 0,
      sheet.columns?.length ?? 0,
      sheet.freezeColumns ?? 0,
      sheet.columnIds?.length ?? 0,
    ),
    rows = Math.max(1, sheet.minDimensions?.[1] ?? 0, sheet.data.length, sheet.rowIds?.length ?? 0);
  for (const row of sheet.data) columns = Math.max(columns, row.length);
  for (const index of Object.keys(sheet.rows ?? {})) rows = Math.max(rows, Number(index) + 1);
  const include = (address: string, width = 1, height = 1) => {
    const cell = parseSpreadsheetAddress(address);
    if (cell) {
      columns = Math.max(columns, cell.x + width);
      rows = Math.max(rows, cell.y + height);
    }
  };
  for (const map of [sheet.style, sheet.formats, sheet.inputValues, sheet.cellProvenance])
    for (const address of Object.keys(map ?? {})) include(address);
  for (const [address, [width, height]] of Object.entries(sheet.mergeCells ?? {}))
    include(address, width, height);
  for (const thread of workbook?.comments ?? []) {
    if (thread.sheetId !== sheet.id || thread.orphaned) continue;
    const range = parseSpreadsheetRange(thread.range ?? '');
    if (range) {
      columns = Math.max(columns, range.end.x + 1);
      rows = Math.max(rows, range.end.y + 1);
    }
  }
  for (const suggestion of workbook?.suggestions ?? []) {
    if (suggestion.outdated) continue;
    for (const change of suggestion.changes)
      if (change.sheetId === sheet.id) include(change.address);
  }
  return { columns, rows };
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
  sheet.minDimensions = sheet.minDimensions
    ? [extent.columns, extent.rows]
    : [Math.max(defaults[0], extent.columns), Math.max(defaults[1], extent.rows)];
  return next;
}

export function extendWorksheetIdentities(
  sheet: Worksheet,
  factory: WorkbookIdFactory = createWorkbookEntityId,
): void {
  const extent = worksheetExtent(sheet);
  for (const [axis, length] of [
    ['row', extent.rows],
    ['column', extent.columns],
  ] as const) {
    const ids = axis === 'row' ? sheet.rowIds : sheet.columnIds;
    if (!ids) continue;
    while (ids.length < length) ids.push(factory(axis));
  }
}
