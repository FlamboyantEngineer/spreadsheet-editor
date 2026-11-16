import { requireDefined } from './require-defined.js';
import { type CellInput, DEFAULT_WORKBOOK_LIMITS } from './workbook.js';

/** Fields read so far, the field being read, and whether it is inside quotes. */
interface Reading {
  rows: string[][];
  field: string;
  quoted: boolean;
  cells: number;
}

/** Ends the current field, within the clipboard cell and column limits. */
function endField(reading: Reading) {
  const row = requireDefined(reading.rows.at(-1));
  if (++reading.cells > DEFAULT_WORKBOOK_LIMITS.maxCells || row.length >= 16_384)
    throw new RangeError('Clipboard exceeds the cell limit.');
  row.push(reading.field);
  reading.field = '';
}

/** A quote doubles inside quotes, opens or closes a quoted field, or is literal mid-field. */
function quote(reading: Reading, text: string, index: number): number {
  if (reading.quoted && text[index + 1] === '"') {
    reading.field += '"';
    return index + 1;
  }
  if (!reading.field || reading.quoted) reading.quoted = !reading.quoted;
  else reading.field += '"';
  return index;
}

/** Reads the character at `index`, returning the index of the last character it consumed. */
function step(
  reading: Reading,
  { text, delimiter }: { text: string; delimiter: string },
  index: number,
): number {
  const character = text[index];
  if (character === '"') return quote(reading, text, index);
  if (reading.quoted || (character !== delimiter && character !== '\n' && character !== '\r')) {
    reading.field += character;
    return index;
  }
  endField(reading);
  if (character === delimiter) return index;
  reading.rows.push([]);
  return character === '\r' && text[index + 1] === '\n' ? index + 1 : index;
}

/** RFC 4180-style quoting, also used for tab-delimited clipboard text. */
export function parseDelimited(text: string, delimiter = '\t'): string[][] {
  if (
    text.length > 16 * 1024 * 1024 ||
    new TextEncoder().encode(text).byteLength > 16 * 1024 * 1024
  )
    throw new RangeError('Clipboard text exceeds the byte limit.');
  const reading: Reading = { rows: [[]], field: '', quoted: false, cells: 0 };
  for (let index = 0; index < text.length; index++) {
    index = step(reading, { text, delimiter }, index);
    if (reading.field.length > DEFAULT_WORKBOOK_LIMITS.maxCellLength)
      throw new RangeError('Clipboard cell exceeds the text limit.');
  }
  if (reading.quoted) throw new TypeError('Unclosed quoted field.');
  endField(reading);
  const { rows } = reading,
    last = requireDefined(rows.at(-1));
  if (rows.length > 1 && last.length === 1 && last[0] === '' && /[\r\n]$/.test(text)) rows.pop();
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
