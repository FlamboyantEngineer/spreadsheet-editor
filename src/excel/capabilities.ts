import type { ExcelFormulaCapabilities, ExcelHostFunction } from './contracts.js';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';

const MAX_FUNCTIONS = 128;
const MAX_ALIASES = 16;
const NAME = /^[A-Za-z_][A-Za-z0-9_.]{0,127}$/;
const plainObject = (value: object) =>
  [Object.prototype, null].includes(Object.getPrototypeOf(value));

/** Capabilities cross a process boundary as data; callbacks and implicit host catalogs are forbidden. */
export function validateFormulaCapabilities(value: unknown): ExcelFormulaCapabilities {
  if (value === undefined) return { hostFunctions: [] };
  if (!value || typeof value !== 'object' || Array.isArray(value) || !plainObject(value))
    throw new TypeError('Excel formula capabilities must be an object.');
  const candidate = value as ExcelFormulaCapabilities;
  if (
    Object.keys(candidate).some((key) => key !== 'hostFunctions') ||
    !Array.isArray(candidate.hostFunctions) ||
    candidate.hostFunctions.length > MAX_FUNCTIONS
  )
    throw new TypeError('Excel formula capabilities need at most 128 host functions.');
  const names = new Set<string>();
  const name = (input: unknown) => {
    if (typeof input !== 'string' || !NAME.test(input))
      throw new TypeError('Excel host function names must be bounded formula identifiers.');
    const normalized = input.toUpperCase();
    if (SUPPORTED_FUNCTIONS.includes(normalized))
      throw new TypeError(`Excel host capabilities cannot override built-in ${normalized}.`);
    if (names.has(normalized)) throw new TypeError(`Duplicate Excel host function: ${normalized}.`);
    names.add(normalized);
    return normalized;
  };
  const hostFunctions = candidate.hostFunctions.map((entry): ExcelHostFunction => {
    if (
      !entry ||
      typeof entry !== 'object' ||
      Array.isArray(entry) ||
      !plainObject(entry) ||
      Object.keys(entry).some((key) => !['name', 'aliases', 'capturedText'].includes(key)) ||
      (entry.aliases !== undefined &&
        (!Array.isArray(entry.aliases) || entry.aliases.length > MAX_ALIASES)) ||
      (entry.capturedText !== undefined && !['plain', 'html-to-text'].includes(entry.capturedText))
    )
      throw new TypeError('Invalid Excel host function descriptor.');
    return {
      name: name(entry.name),
      ...(entry.aliases ? { aliases: entry.aliases.map(name) } : {}),
      capturedText: entry.capturedText ?? 'plain',
    };
  });
  return { hostFunctions };
}

export function hostFunctionMap(capabilities: ExcelFormulaCapabilities) {
  const functions = new Map<string, ExcelHostFunction>();
  for (const descriptor of capabilities.hostFunctions)
    for (const name of [descriptor.name, ...(descriptor.aliases ?? [])])
      functions.set(name, descriptor);
  return functions;
}
