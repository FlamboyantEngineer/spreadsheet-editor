import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { readXlsxXml } from './xlsx-xml.js';

export interface WorksheetXml {
  rows: number[];
  cells: Map<string, string | undefined>;
  features: Set<string>;
}

interface CellXml {
  address: string;
  numeric: boolean;
  value?: string;
}

function readCell(
  address: string,
  type: string | undefined,
  cells: WorksheetXml['cells'],
): CellXml {
  const position =
    /^[A-Z]{1,3}[1-9]\d{0,6}$/.test(address ?? '') && parseSpreadsheetAddress(address);
  if (!position || position.x >= 16384 || position.y >= 1048576 || cells.has(address))
    throw new TypeError('XLSX contains an invalid or duplicate cell address.');
  return { address, numeric: type === undefined || type === 'n' };
}

function rowWithGeometry(attributes: Readonly<Record<string, string>>): number | undefined {
  const row = Number(attributes.r);
  if (!Number.isInteger(row) || row < 1 || row > 1048576)
    throw new TypeError('XLSX contains an invalid row number.');
  return attributes.ht !== undefined || attributes.hidden !== undefined ? row : undefined;
}

/** Read only authored numeric values and the explicitly stored cells and rows. */
export function readWorksheetXml(bytes: Uint8Array, cellLimit: number): WorksheetXml {
  const rows: number[] = [],
    cells = new Map<string, string | undefined>(),
    features = new Set<string>();
  let data = false;
  let cell: CellXml | undefined;
  let numericValue = false;
  readXlsxXml(bytes, {
    open(name, attributes, depth) {
      if (depth === 2) {
        features.add(name);
        data = name === 'sheetData';
      }
      if (!data) return;
      if (depth === 3 && name === 'row') {
        const row = rowWithGeometry(attributes);
        if (row !== undefined) rows.push(row);
      }
      if (depth === 4 && name === 'c') {
        if (cells.size >= cellLimit) throw new RangeError('XLSX exceeds the stored cell limit.');
        cell = readCell(attributes.r, attributes.t, cells);
      }
      if (depth !== 5 || !cell) return;
      if (name === 'f') cell.numeric = false;
      if (name === 'v' && cell.numeric) {
        cell.value = '';
        numericValue = true;
      }
    },
    close(depth) {
      if (depth === 5) numericValue = false;
      if (depth === 4 && cell) {
        cells.set(cell.address, cell.numeric ? cell.value?.trim() || undefined : undefined);
        cell = undefined;
      }
    },
    text(value) {
      if (numericValue && cell) cell.value = (cell.value ?? '') + value;
    },
  });
  return { rows, cells, features };
}
