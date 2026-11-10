import { copyWorkbook, assertWorkbook, type CellInput, type Workbook } from './workbook.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { CellFormat } from './review.js';

export interface CellSelection {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}
export const INITIAL_SELECTION: CellSelection = { x1: 0, y1: 0, x2: 0, y2: 0 };

export function normalizeSelection(range: CellSelection): CellSelection {
  const values = [range.x1, range.y1, range.x2, range.y2];
  if (
    values.some((value) => !Number.isSafeInteger(value) || value < 0) ||
    Math.max(range.x1, range.x2) > 16_383 ||
    Math.max(range.y1, range.y2) > 1_048_575
  )
    throw new RangeError('Invalid selection.');
  return {
    x1: Math.min(range.x1, range.x2),
    y1: Math.min(range.y1, range.y2),
    x2: Math.max(range.x1, range.x2),
    y2: Math.max(range.y1, range.y2),
  };
}
export function selectionAddress(range: CellSelection): string {
  const r = normalizeSelection(range),
    first = formatSpreadsheetAddress(r.x1, r.y1),
    last = formatSpreadsheetAddress(r.x2, r.y2);
  return first === last ? first : `${first}:${last}`;
}
export function selectionCells(range: CellSelection): { x: number; y: number; address: string }[] {
  const r = normalizeSelection(range);
  if ((r.x2 - r.x1 + 1) * (r.y2 - r.y1 + 1) > 50_000)
    throw new RangeError('Selection exceeds the cell limit.');
  const cells = [];
  for (let y = r.y1; y <= r.y2; y++)
    for (let x = r.x1; x <= r.x2; x++)
      cells.push({ x, y, address: formatSpreadsheetAddress(x, y) });
  return cells;
}

export function writeCells(
  workbook: Workbook,
  sheetId: string,
  x: number,
  y: number,
  values: CellInput[][],
): Workbook {
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  const width = Math.max(0, ...values.map((row) => row.length));
  normalizeSelection({
    x1: x,
    y1: y,
    x2: x + Math.max(0, width - 1),
    y2: y + Math.max(0, values.length - 1),
  });
  if (y + values.length > 50_000 || width * values.length > 50_000)
    throw new RangeError('Paste exceeds the cell limit.');
  values.forEach((row, dy) => {
    while (sheet.data.length <= y + dy) sheet.data.push([]);
    row.forEach((value, dx) => {
      while (sheet.data[y + dy].length <= x + dx) sheet.data[y + dy].push(null);
      sheet.data[y + dy][x + dx] = value;
    });
  });
  assertWorkbook(next);
  return next;
}

export function readSelection(
  workbook: Workbook,
  sheetId: string,
  range: CellSelection,
): CellInput[][] {
  const r = normalizeSelection(range),
    sheet = workbook.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  selectionCells(r);
  return Array.from({ length: r.y2 - r.y1 + 1 }, (_, y) =>
    Array.from({ length: r.x2 - r.x1 + 1 }, (_, x) => sheet.data[r.y1 + y]?.[r.x1 + x] ?? null),
  );
}

export function parseCellStyle(style: string | undefined): Record<string, string> {
  const properties: Record<string, string> = {};
  for (const declaration of style?.split(';') ?? []) {
    const colon = declaration.indexOf(':');
    if (colon >= 0)
      properties[declaration.slice(0, colon).trim()] = declaration.slice(colon + 1).trim();
  }
  return properties;
}
export function formatCells(
  workbook: Workbook,
  sheetId: string,
  range: CellSelection,
  properties: Record<string, string | null>,
  format?: CellFormat | null,
): Workbook {
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  sheet.style = { ...sheet.style };
  sheet.formats = { ...sheet.formats };
  for (const { address } of selectionCells(range)) {
    const style = parseCellStyle(sheet.style[address]);
    for (const [property, value] of Object.entries(properties))
      if (value === null) delete style[property];
      else style[property] = value;
    const serialized = Object.entries(style)
      .map(([key, value]) => `${key}: ${value}`)
      .join('; ');
    if (serialized) sheet.style[address] = serialized;
    else delete sheet.style[address];
    if (format === null) delete sheet.formats[address];
    else if (format)
      sheet.formats[address] = {
        ...format,
        ...(format.kind === 'unit'
          ? {
              inputUnit:
                sheet.formats[address]?.inputUnit ?? sheet.formats[address]?.unit ?? format.unit,
            }
          : {}),
      };
  }
  return next;
}
export function clearSelection(
  workbook: Workbook,
  sheetId: string,
  range: CellSelection,
  formatting = false,
): Workbook {
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId)!;
  for (const { x, y, address } of selectionCells(range)) {
    if (formatting) {
      if (sheet.style) delete sheet.style[address];
      if (sheet.formats) delete sheet.formats[address];
    } else if (sheet.data[y]?.[x] !== undefined) sheet.data[y][x] = null;
  }
  return next;
}

export function borderCells(
  workbook: Workbook,
  sheetId: string,
  range: CellSelection,
  kind: 'all' | 'outside' | 'none' | 'top' | 'right' | 'bottom' | 'left',
): Workbook {
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId)!;
  const r = normalizeSelection(range);
  sheet.style = { ...sheet.style };
  const cells = selectionCells(r);
  const targeted = (cell: (typeof cells)[number], side: string) =>
    kind === 'all' ||
    kind === side ||
    (kind === 'outside' &&
      ((side === 'top' && cell.y === r.y1) ||
        (side === 'bottom' && cell.y === r.y2) ||
        (side === 'left' && cell.x === r.x1) ||
        (side === 'right' && cell.x === r.x2)));
  const sides = ['top', 'right', 'bottom', 'left'];
  const remove = cells.every((cell) => {
    const style = parseCellStyle(sheet.style![cell.address]);
    return sides.filter((side) => targeted(cell, side)).every((side) => !!style[`border-${side}`]);
  });
  for (const cell of cells) {
    const style = parseCellStyle(sheet.style[cell.address]);
    for (const side of sides) {
      if (kind === 'none') {
        delete style[`border-${side}`];
        delete style.border;
      } else if (targeted(cell, side)) {
        if (remove) delete style[`border-${side}`];
        else style[`border-${side}`] = '1px solid currentColor';
      }
    }
    sheet.style[cell.address] = Object.entries(style)
      .map(([key, value]) => `${key}:${value}`)
      .join(';');
  }
  return next;
}

/** RFC 4180-style quoting, also used for tab-delimited clipboard text. */
export function parseDelimited(text: string, delimiter = '\t'): string[][] {
  const rows: string[][] = [[]];
  let field = '',
    quoted = false;
  for (let index = 0; index < text.length; index++) {
    const character = text[index];
    if (character === '"') {
      if (quoted && text[index + 1] === '"') {
        field += '"';
        index++;
      } else if (!field || quoted) quoted = !quoted;
      else field += character;
    } else if (!quoted && (character === delimiter || character === '\n' || character === '\r')) {
      rows.at(-1)!.push(field);
      field = '';
      if (character !== delimiter) {
        if (character === '\r' && text[index + 1] === '\n') index++;
        rows.push([]);
      }
    } else field += character;
  }
  if (quoted) throw new TypeError('Unclosed quoted field.');
  rows.at(-1)!.push(field);
  if (
    rows.length > 1 &&
    rows.at(-1)!.length === 1 &&
    rows.at(-1)![0] === '' &&
    /[\r\n]$/.test(text)
  )
    rows.pop();
  return rows;
}
export function serializeDelimited(rows: CellInput[][], delimiter = '\t'): string {
  return rows
    .map((row) =>
      row
        .map((value) => {
          const text = value == null ? '' : String(value);
          return text.includes(delimiter) || /["\r\n]/.test(text)
            ? `"${text.replace(/"/g, '""')}"`
            : text;
        })
        .join(delimiter),
    )
    .join('\n');
}

export function mergeSelection(
  workbook: Workbook,
  sheetId: string,
  range: CellSelection,
  unmerge = false,
): Workbook {
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId)!;
  const r = normalizeSelection(range);
  sheet.mergeCells = { ...sheet.mergeCells };
  for (const [address, [width, height]] of Object.entries(sheet.mergeCells)) {
    const anchor = parseSpreadsheetAddress(address)!;
    if (
      anchor.x <= r.x2 &&
      anchor.x + width - 1 >= r.x1 &&
      anchor.y <= r.y2 &&
      anchor.y + height - 1 >= r.y1
    )
      delete sheet.mergeCells[address];
  }
  if (!unmerge && (r.x2 > r.x1 || r.y2 > r.y1))
    sheet.mergeCells[formatSpreadsheetAddress(r.x1, r.y1)] = [r.x2 - r.x1 + 1, r.y2 - r.y1 + 1];
  return next;
}
