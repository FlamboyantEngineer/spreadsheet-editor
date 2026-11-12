import { validateFormulaCapabilities } from './capabilities.js';
import type { ExcelFormulaCapabilities } from './contracts.js';

const HEADER_LIMIT = 65_536;
/** Length-prefixed JSON header followed by bounded binary data avoids base64 expansion of imports. */
export function encodeImportRequest(
  bytes: Uint8Array,
  capabilities?: ExcelFormulaCapabilities,
): Buffer {
  if (bytes.byteLength > 20 * 1024 * 1024)
    throw new RangeError('Excel process input exceeds its byte limit.');
  const header = Buffer.from(
    JSON.stringify({ formulaCapabilities: validateFormulaCapabilities(capabilities) }),
  );
  if (header.length > HEADER_LIMIT)
    throw new RangeError('Excel capability header exceeds its byte limit.');
  const prefix = Buffer.alloc(4);
  prefix.writeUInt32LE(header.length);
  return Buffer.concat([prefix, header, bytes]);
}
export function decodeImportRequest(request: Buffer) {
  if (request.length < 4) throw new TypeError('Missing Excel import header.');
  const length = request.readUInt32LE(0);
  if (length > HEADER_LIMIT || length + 4 > request.length)
    throw new RangeError('Invalid Excel capability header length.');
  const header = JSON.parse(request.subarray(4, length + 4).toString()) as Record<string, unknown>;
  if (
    !header ||
    typeof header !== 'object' ||
    Array.isArray(header) ||
    Object.keys(header).some((key) => key !== 'formulaCapabilities')
  )
    throw new TypeError('Invalid Excel import header.');
  const formulaCapabilities = validateFormulaCapabilities(header.formulaCapabilities);
  const bytes = request.subarray(length + 4);
  if (bytes.length > 20 * 1024 * 1024)
    throw new RangeError('Excel process input exceeds its byte limit.');
  return { bytes, formulaCapabilities };
}
