import { validateFormulaCapabilities } from './capabilities.js';
import type { ExcelFormulaCapabilities, XlsxExport } from './contracts.js';

const HEADER_LIMIT = 65_536;
const importLimit = 20 * 1024 * 1024;
export const excelOutputLimit = 100 * 1024 * 1024;
export type ExcelOperation = 'import' | 'export';

export function excelInputLimit(operation: ExcelOperation): number {
  return operation === 'import' ? importLimit + HEADER_LIMIT + 4 : 100 * 1024 * 1024;
}

/** Length-prefixed JSON metadata, followed by the original binary bytes. */
export function encodeExcelHeader(value: object): Buffer {
  const header = Buffer.from(JSON.stringify(value));
  if (header.length > HEADER_LIMIT) throw new RangeError('Excel metadata exceeds its byte limit.');
  const prefix = Buffer.alloc(4);
  prefix.writeUInt32LE(header.length);
  return Buffer.concat([prefix, header]);
}

function decodeExcelFrame(frame: Buffer) {
  if (frame.length < 4) throw new TypeError('Missing Excel metadata header.');
  const length = frame.readUInt32LE(0);
  if (length > HEADER_LIMIT || length + 4 > frame.length)
    throw new RangeError('Invalid Excel metadata header length.');
  const header = JSON.parse(frame.subarray(4, length + 4).toString()) as Record<string, unknown>;
  if (!header || typeof header !== 'object' || Array.isArray(header))
    throw new TypeError('Invalid Excel metadata header.');
  return { header, bytes: frame.subarray(length + 4) };
}

/** Length-prefixed JSON header followed by bounded binary data avoids base64 expansion of imports. */
export function encodeImportRequest(
  bytes: Uint8Array,
  capabilities?: ExcelFormulaCapabilities,
): Buffer {
  if (bytes.byteLength > importLimit)
    throw new RangeError('Excel process input exceeds its byte limit.');
  const header = encodeExcelHeader({
    formulaCapabilities: validateFormulaCapabilities(capabilities),
  });
  return Buffer.concat([header, bytes]);
}
export function decodeImportRequest(request: Buffer) {
  const { header, bytes } = decodeExcelFrame(request);
  if (Object.keys(header).some((key) => key !== 'formulaCapabilities'))
    throw new TypeError('Invalid Excel import header.');
  const formulaCapabilities = validateFormulaCapabilities(header.formulaCapabilities);
  if (bytes.length > importLimit)
    throw new RangeError('Excel process input exceeds its byte limit.');
  return { bytes, formulaCapabilities };
}

export function decodeExportResponse(response: Buffer): XlsxExport {
  const { header, bytes } = decodeExcelFrame(response);
  return {
    bytes: new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength),
    report: header.report as XlsxExport['report'],
  };
}
