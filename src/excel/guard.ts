import { unzipSync } from 'fflate';

export interface XlsxLimits {
  maxBytes?: number;
  maxExpandedBytes?: number;
  maxEntries?: number;
  maxWorksheets?: number;
  maxCells?: number;
}
export const XLSX_LIMITS = Object.freeze({
  maxBytes: 20 * 1024 * 1024,
  maxExpandedBytes: 128 * 1024 * 1024,
  maxEntries: 8192,
  maxWorksheets: 200,
  maxCells: 250000,
});
const decoder = new TextDecoder('utf-8', { fatal: true });
const reject = (): never => {
  throw new TypeError(
    'Expected a bounded, unencrypted .xlsx workbook without macros or XML entities.',
  );
};

/** Inspect the central directory and authored XML before ExcelJS receives any workbook. */
export function inspectXlsx(input: Uint8Array, limits: XlsxLimits = {}) {
  const bounds = { ...XLSX_LIMITS, ...limits };
  if (input.length < 22 || input.length > bounds.maxBytes) reject();
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  const uint16 = (offset: number) => view.getUint16(offset, true),
    uint32 = (offset: number) => view.getUint32(offset, true);
  if (uint32(0) !== 0x04034b50) reject();
  let end = -1;
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65557); i--)
    if (uint32(i) === 0x06054b50 && i + 22 + uint16(i + 20) === input.length) {
      end = i;
      break;
    }
  if (end < 0 || uint16(end + 4) !== 0 || uint16(end + 6) !== 0) reject();
  const count = uint16(end + 10),
    centralSize = uint32(end + 12),
    centralOffset = uint32(end + 16);
  if (
    !count ||
    count > bounds.maxEntries ||
    uint16(end + 8) !== count ||
    centralOffset + centralSize !== end
  )
    reject();
  const entries = new Map<string, number>();
  let expanded = 0,
    offset = centralOffset;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > end || uint32(offset) !== 0x02014b50) reject();
    const flags = uint16(offset + 8),
      method = uint16(offset + 10),
      size = uint32(offset + 24),
      nameLength = uint16(offset + 28),
      extra = uint16(offset + 30),
      comment = uint16(offset + 32),
      local = uint32(offset + 42);
    if (
      flags & 1 ||
      ![0, 8].includes(method) ||
      size === 0xffffffff ||
      local >= centralOffset ||
      offset + 46 + nameLength + extra + comment > end
    )
      reject();
    const name = decoder.decode(input.subarray(offset + 46, offset + 46 + nameLength));
    if (
      !name ||
      name.startsWith('/') ||
      name.includes('\\') ||
      name.includes('\0') ||
      name.split('/').some((part) => part === '..' || part === '.') ||
      entries.has(name) ||
      /(?:vbaProject\.bin|\.xlsb$)/i.test(name)
    )
      reject();
    expanded += size;
    if (expanded > bounds.maxExpandedBytes || size > 32 * 1024 * 1024) reject();
    entries.set(name, size);
    offset += 46 + nameLength + extra + comment;
  }
  if (
    offset !== end ||
    !entries.has('[Content_Types].xml') ||
    !entries.has('xl/workbook.xml') ||
    !entries.has('_rels/.rels')
  )
    reject();
  const worksheets = [...entries.keys()].filter((name) =>
    /^xl\/worksheets\/sheet[1-9]\d*\.xml$/.test(name),
  );
  if (!worksheets.length || worksheets.length > bounds.maxWorksheets) reject();
  const parts = unzipSync(input);
  let cells = 0;
  for (const [name, bytes] of Object.entries(parts)) {
    if (entries.get(name) !== bytes.length) reject();
    if (!/\.(xml|rels)$/.test(name)) continue;
    const xml = decoder.decode(bytes);
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) reject();
    let depth = 0;
    for (const token of xml.matchAll(/<([^>]+)>/g)) {
      const value = token[1];
      if (value.startsWith('?') || value.startsWith('!')) continue;
      if (value.startsWith('/')) depth--;
      else if (!value.endsWith('/')) depth++;
      if (depth < 0 || depth > 128) reject();
    }
    if (depth !== 0) reject();
    if (worksheets.includes(name)) {
      cells += [...xml.matchAll(/<(?:\w+:)?c(?:\s|>)/g)].length;
      if (cells > bounds.maxCells) reject();
    }
    if (
      name === '[Content_Types].xml' &&
      (!xml.includes(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      ) ||
        /macroEnabled|vbaProject|spreadsheetml\.sheet\.binary/i.test(xml))
    )
      reject();
  }
  return { parts, worksheets: worksheets.length, cells };
}
