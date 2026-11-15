import type { NumericValue } from '../calculation/values.js';
import { assertNumericValue } from '../calculation/values.js';
import { assertCellFormat } from './cell-format.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import { worksheetExtent } from './extent.js';
import {
  assertReviewData,
  type CellFormat,
  type SpreadsheetComment,
  type SpreadsheetSuggestion,
} from './review.js';

export type CellInput = string | number | boolean | null;

export interface CellProvenance {
  actor: { id: string; name: string };
  source: 'human' | 'ai' | 'system';
  at?: string;
}

export interface Worksheet {
  id: string;
  name: string;
  /** Keep numeric literals as strings to preserve all authored digits. */
  data: CellInput[][];
  inputValues?: Record<string, NumericValue>;
  rowIds?: string[];
  columnIds?: string[];
  minDimensions?: [number, number];
  columns?: { width?: number; title?: string }[];
  rows?: Record<string, { height: number }>;
  mergeCells?: Record<string, [number, number]>;
  style?: Record<string, string>;
  formats?: Record<string, CellFormat>;
  freezeColumns?: number;
  cellProvenance?: Record<string, CellProvenance>;
}

export interface Workbook {
  schemaVersion: 1;
  worksheets: Worksheet[];
  name?: string;
  comments?: SpreadsheetComment[];
  suggestions?: SpreadsheetSuggestion[];
}

export interface WorkbookLimits {
  maxCells?: number;
  maxCellLength?: number;
  maxWorksheets?: number;
}

export const DEFAULT_WORKBOOK_LIMITS = Object.freeze({
  maxCells: 250_000,
  maxCellLength: 10_000,
  maxWorksheets: 200,
});

/** The form worksheet names are compared in: spaces collapsed, case folded. */
export function worksheetNameKey(name: string): string {
  return name.trim().replace(/\s+/g, ' ').toUpperCase();
}

export function createWorkbook(data: CellInput[][] = [['']]): Workbook {
  return { schemaVersion: 1, worksheets: [{ id: 'sheet-1', name: 'Sheet1', data }] };
}

/** Validate the same bounded JSON contract on the client and at the server boundary. */
export function assertWorkbook(
  value: unknown,
  limits: WorkbookLimits = {},
): asserts value is Workbook {
  const bounds = { ...DEFAULT_WORKBOOK_LIMITS, ...limits };
  if (!value || typeof value !== 'object') throw new TypeError('Expected a workbook.');
  const workbook = value as Partial<Workbook>;
  if (
    workbook.schemaVersion !== 1 ||
    !Array.isArray(workbook.worksheets) ||
    !workbook.worksheets.length ||
    workbook.worksheets.length > bounds.maxWorksheets
  ) {
    throw new TypeError('Invalid workbook version or worksheet count.');
  }
  assertReviewData(workbook as Workbook);
  if (
    workbook.name !== undefined &&
    (typeof workbook.name !== 'string' || workbook.name.length > 512)
  )
    throw new TypeError('Invalid workbook name.');
  const ids = new Set<string>(),
    names = new Set<string>();
  let cells = 0,
    rowCount = 0;
  for (const sheet of workbook.worksheets) {
    if (
      !sheet ||
      typeof sheet.id !== 'string' ||
      !sheet.id ||
      sheet.id.length > 256 ||
      typeof sheet.name !== 'string' ||
      !sheet.name.trim() ||
      sheet.name.length > 256 ||
      ids.has(sheet.id) ||
      names.has(worksheetNameKey(sheet.name)) ||
      !Array.isArray(sheet.data)
    ) {
      throw new TypeError('Worksheets need unique ids, unique names, and row data.');
    }
    if (
      sheet.minDimensions !== undefined &&
      (!Array.isArray(sheet.minDimensions) ||
        sheet.minDimensions.length !== 2 ||
        sheet.minDimensions[0] * sheet.minDimensions[1] > bounds.maxCells ||
        sheet.minDimensions.some(
          (n, i) => !Number.isInteger(n) || n < 1 || n > (i === 0 ? 16384 : 1048576),
        ))
    )
      throw new TypeError('Invalid worksheet dimensions.');
    for (const axis of ['rowIds', 'columnIds'] as const) {
      const list = sheet[axis];
      if (
        list !== undefined &&
        (!Array.isArray(list) ||
          list.some((id) => typeof id !== 'string' || !id || id.length > 256) ||
          new Set(list).size !== list.length)
      )
        throw new TypeError('Invalid worksheet line identities.');
    }
    ids.add(sheet.id);
    names.add(worksheetNameKey(sheet.name));
    rowCount += sheet.data.length;
    if (rowCount > bounds.maxCells || sheet.data.length > 1_048_576)
      throw new RangeError('Worksheet exceeds the row limit.');
    if (
      sheet.columns !== undefined &&
      (!Array.isArray(sheet.columns) ||
        sheet.columns.length > 16_384 ||
        sheet.columns.some(
          (column) =>
            !column ||
            (column.width !== undefined &&
              (typeof column.width !== 'number' ||
                !Number.isFinite(column.width) ||
                column.width < 1)) ||
            (column.title !== undefined && typeof column.title !== 'string'),
        ))
    )
      throw new TypeError('Invalid worksheet columns.');
    if (
      sheet.mergeCells !== undefined &&
      (!sheet.mergeCells || typeof sheet.mergeCells !== 'object' || Array.isArray(sheet.mergeCells))
    )
      throw new TypeError('Invalid merged cells.');
    for (const [address, span] of Object.entries(sheet.mergeCells ?? {})) {
      if (
        !/^[A-Z]+[1-9]\d*$/.test(address) ||
        !Array.isArray(span) ||
        span.length !== 2 ||
        span[0] * span[1] > bounds.maxCells ||
        span.some((value) => !Number.isSafeInteger(value) || value < 1 || value > bounds.maxCells)
      )
        throw new TypeError('Invalid merged cells.');
    }
    if (
      sheet.rows !== undefined &&
      (!sheet.rows || typeof sheet.rows !== 'object' || Array.isArray(sheet.rows))
    )
      throw new TypeError('Invalid worksheet rows.');
    for (const [index, row] of Object.entries(sheet.rows ?? {})) {
      if (!/^\d+$/.test(index) || !row || !Number.isFinite(row.height) || row.height < 1)
        throw new TypeError('Invalid row height.');
    }
    if (
      sheet.style !== undefined &&
      (!sheet.style ||
        typeof sheet.style !== 'object' ||
        Array.isArray(sheet.style) ||
        Object.values(sheet.style).some((value) => typeof value !== 'string'))
    )
      throw new TypeError('Invalid cell styles.');
    if (
      sheet.freezeColumns !== undefined &&
      (!Number.isSafeInteger(sheet.freezeColumns) ||
        sheet.freezeColumns < 0 ||
        sheet.freezeColumns > 16_384)
    )
      throw new TypeError('Invalid frozen column count.');
    if (
      sheet.formats !== undefined &&
      (!sheet.formats || typeof sheet.formats !== 'object' || Array.isArray(sheet.formats))
    )
      throw new TypeError('Invalid cell formats.');
    for (const [address, format] of Object.entries(sheet.formats ?? {})) {
      if (!/^[A-Z]+[1-9]\d*$/.test(address)) throw new TypeError('Invalid cell format address.');
      assertCellFormat(format);
    }
    if (sheet.cellProvenance !== undefined) {
      if (
        !sheet.cellProvenance ||
        typeof sheet.cellProvenance !== 'object' ||
        Array.isArray(sheet.cellProvenance)
      )
        throw new TypeError('Invalid cell provenance.');
      for (const [address, entry] of Object.entries(sheet.cellProvenance))
        if (
          !/^[A-Z]+[1-9]\d*$/.test(address) ||
          !entry ||
          !entry.actor ||
          typeof entry.actor.id !== 'string' ||
          !entry.actor.id ||
          entry.actor.id.length > 256 ||
          typeof entry.actor.name !== 'string' ||
          entry.actor.name.length > 512 ||
          !['human', 'ai', 'system'].includes(entry.source) ||
          (entry.at !== undefined &&
            (typeof entry.at !== 'string' || !Number.isFinite(Date.parse(entry.at))))
        )
          throw new TypeError('Invalid cell provenance.');
    }
    if (
      sheet.inputValues !== undefined &&
      (!sheet.inputValues ||
        typeof sheet.inputValues !== 'object' ||
        Array.isArray(sheet.inputValues))
    )
      throw new TypeError('Invalid canonical input values.');
    for (const [address, canonical] of Object.entries(sheet.inputValues ?? {})) {
      const position = /^[A-Z]+[1-9]\d*$/.test(address) ? parseSpreadsheetAddress(address) : null;
      const input = position && sheet.data[position.y]?.[position.x];
      if (
        !position ||
        input === undefined ||
        (typeof input === 'string' && input.trimStart().startsWith('='))
      )
        throw new TypeError('Canonical input requires a corresponding nonformula data entry.');
      assertNumericValue(canonical, bounds.maxCellLength);
    }
    for (const map of [
      sheet.style,
      sheet.formats,
      sheet.cellProvenance,
      sheet.inputValues,
      sheet.mergeCells,
    ])
      for (const address of Object.keys(map ?? {})) {
        const position = /^[A-Z]+[1-9]\d*$/.test(address) ? parseSpreadsheetAddress(address) : null;
        if (!position || position.x > 16383 || position.y > 1048575)
          throw new TypeError('Invalid cell metadata address.');
      }
    const extent = worksheetExtent(sheet, workbook as Workbook);
    if (
      extent.columns > 16384 ||
      extent.rows > 1048576 ||
      extent.columns * extent.rows > bounds.maxCells
    )
      throw new RangeError('Worksheet extent exceeds the cell limit.');
    if (
      (sheet.rowIds && sheet.rowIds.length !== extent.rows) ||
      (sheet.columnIds && sheet.columnIds.length !== extent.columns)
    )
      throw new TypeError('Worksheet line identities must cover the effective extent.');
    for (const [address] of Object.entries(sheet.style ?? {})) {
      const position = parseSpreadsheetAddress(address)!;
      if (position.x >= (sheet.data[position.y]?.length ?? 0)) cells++;
    }
    const counted = new Set(Object.keys(sheet.style ?? {}));
    for (const map of [sheet.formats, sheet.cellProvenance, sheet.inputValues])
      for (const address of Object.keys(map ?? {})) {
        const position = parseSpreadsheetAddress(address)!;
        if (!counted.has(address) && position.x >= (sheet.data[position.y]?.length ?? 0)) cells++;
        counted.add(address);
      }
    for (const row of sheet.data) {
      if (!Array.isArray(row)) throw new TypeError('Expected a worksheet row.');
      if (row.length > 16_384) throw new RangeError('Worksheet exceeds the column limit.');
      cells += row.length;
      if (cells > bounds.maxCells) throw new RangeError('Workbook exceeds the cell limit.');
      for (const cell of row) {
        if (
          cell !== null &&
          typeof cell !== 'string' &&
          typeof cell !== 'boolean' &&
          !(typeof cell === 'number' && Number.isFinite(cell))
        )
          throw new TypeError('Invalid cell value.');
        if (typeof cell === 'string' && cell.length > bounds.maxCellLength)
          throw new RangeError('Cell exceeds the text limit.');
      }
    }
  }
}

export function appendWorksheet(workbook: Workbook): Workbook {
  const next = copyWorkbook(workbook);
  let index = next.worksheets.length + 1;
  while (
    next.worksheets.some(
      (sheet) => sheet.id === `sheet-${index}` || worksheetNameKey(sheet.name) === `SHEET${index}`,
    )
  )
    index++;
  next.worksheets.push({ id: `sheet-${index}`, name: `Sheet${index}`, data: [['']] });
  assertWorkbook(next);
  return next;
}

export function copyWorkbook(workbook: Workbook): Workbook {
  assertWorkbook(workbook);
  return structuredClone(workbook);
}
