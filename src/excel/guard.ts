import { unzipSync } from 'fflate';
import { readWorksheetXml, type WorksheetXml } from './worksheet-xml.js';
import { readXlsxXml } from './xlsx-xml.js';

export interface XlsxLimits {
  maxBytes?: number;
  maxExpandedBytes?: number;
  maxEntries?: number;
  maxWorksheets?: number;
  maxCells?: number;
}
const XLSX_LIMITS = Object.freeze({
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

type Bounds = Required<XlsxLimits>;

/** Little-endian reads over the archive bytes. */
function archive(input: Uint8Array) {
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  return {
    input,
    uint16: (offset: number) => view.getUint16(offset, true),
    uint32: (offset: number) => view.getUint32(offset, true),
  };
}
type Archive = ReturnType<typeof archive>;

/** The offset of the end-of-central-directory record, whose comment reaches the end. */
function directoryEnd({ input, uint16, uint32 }: Archive): number {
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65557); i--)
    if (uint32(i) === 0x06054b50 && i + 22 + uint16(i + 20) === input.length) return i;
  return -1;
}

/** Absolute, normalized names only; no macro project and no binary workbook. */
function unsafeName(name: string, entries: Map<string, number>): boolean {
  return (
    !name ||
    name.startsWith('/') ||
    name.includes('\\') ||
    name.includes('\0') ||
    name.split('/').some((part) => part === '..' || part === '.') ||
    entries.has(name) ||
    /(?:vbaProject\.bin|\.xlsb$)/i.test(name)
  );
}

/** One central directory record: its name and expanded size, and where the next one starts. */
function directoryEntry(
  zip: Archive,
  offset: number,
  { end, centralOffset }: Record<string, number>,
) {
  const { uint16, uint32 } = zip;
  if (offset + 46 > end || uint32(offset) !== 0x02014b50) reject();
  const flags = uint16(offset + 8),
    method = uint16(offset + 10),
    size = uint32(offset + 24),
    nameLength = uint16(offset + 28),
    next = offset + 46 + nameLength + uint16(offset + 30) + uint16(offset + 32);
  if (
    flags & 1 ||
    ![0, 8].includes(method) ||
    size === 0xffffffff ||
    uint32(offset + 42) >= centralOffset ||
    next > end
  )
    reject();
  const name = decoder.decode(zip.input.subarray(offset + 46, offset + 46 + nameLength));
  return { name, size, next };
}

/** The archive's entries and their expanded sizes, read from a bounded central directory. */
function directory(zip: Archive, bounds: Bounds): Map<string, number> {
  const { uint16, uint32 } = zip;
  if (uint32(0) !== 0x04034b50) reject();
  const end = directoryEnd(zip);
  if (end < 0 || uint16(end + 4) !== 0 || uint16(end + 6) !== 0) reject();
  const count = uint16(end + 10),
    centralSize = uint32(end + 12),
    centralOffset = uint32(end + 16);
  if (!count || count > bounds.maxEntries || uint16(end + 8) !== count) reject();
  if (centralOffset + centralSize !== end) reject();
  const entries = new Map<string, number>();
  let expanded = 0,
    offset = centralOffset;
  for (let i = 0; i < count; i++) {
    const { name, size, next } = directoryEntry(zip, offset, { end, centralOffset });
    if (unsafeName(name, entries)) reject();
    expanded += size;
    if (expanded > bounds.maxExpandedBytes || size > 32 * 1024 * 1024) reject();
    entries.set(name, size);
    offset = next;
  }
  if (offset !== end) reject();
  return entries;
}

/** Rejects a content-types part that is not an ordinary, macro-free workbook. */
function checkContentTypes(bytes: Uint8Array) {
  const xml = decoder.decode(bytes);
  if (
    !xml.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml') ||
    /macroEnabled|vbaProject|spreadsheetml\.sheet\.binary/i.test(xml)
  )
    reject();
}

/** Inspect the central directory and authored XML before ExcelJS receives any workbook. */
export function inspectXlsx(input: Uint8Array, limits: XlsxLimits = {}) {
  const bounds = { ...XLSX_LIMITS, ...limits };
  if (input.length < 22 || input.length > bounds.maxBytes) reject();
  const entries = directory(archive(input), bounds);
  if (['[Content_Types].xml', 'xl/workbook.xml', '_rels/.rels'].some((name) => !entries.has(name)))
    reject();
  const worksheetNames = [...entries.keys()].filter((name) =>
    /^xl\/worksheets\/sheet[1-9]\d*\.xml$/.test(name),
  );
  if (!worksheetNames.length || worksheetNames.length > bounds.maxWorksheets) reject();
  const parts = unzipSync(input),
    worksheetPaths = new Set(worksheetNames);
  const worksheets = new Map<string, WorksheetXml>();
  let cells = 0;
  for (const [name, bytes] of Object.entries(parts)) {
    if (entries.get(name) !== bytes.length) reject();
    if (!/\.(xml|rels)$/.test(name)) continue;
    if (name === '[Content_Types].xml') checkContentTypes(bytes);
    if (worksheetPaths.has(name)) {
      const worksheet = readWorksheetXml(bytes, bounds.maxCells - cells);
      worksheets.set(name, worksheet);
      cells += worksheet.cells.size;
    } else readXlsxXml(bytes);
  }
  return { parts, worksheets };
}
