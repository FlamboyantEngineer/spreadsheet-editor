import {
  arity,
  checked,
  fail,
  flat,
  type FunctionCall,
  type FunctionImpl,
  integer,
  matrix,
  number,
  scalar,
  truth,
  type Value,
  zeroOf,
} from './function-call.js';
import type { CellValue } from './values.js';
import { compileWildcard, hasWildcard } from './wildcard.js';

function transpose(call: FunctionCall) {
  arity(call, 1);
  const rows = matrix(call.args[0]);
  return Array.from({ length: rows[0]?.length ?? 0 }, (_, x) => rows.map((row) => row[x] ?? null));
}

function position(call: FunctionCall, value: Value): number {
  const numeric = number(call, value);
  if (numeric.dimension || numeric.quantityKind) fail();
  const whole = BigInt(numeric.numerator) / BigInt(numeric.denominator);
  if (whole < 0n || whole > 1048576n) fail('#REF!');
  return Number(whole);
}

function index(call: FunctionCall) {
  arity(call, 2, 3);
  const { args } = call;
  const rows = matrix(args[0]),
    height = rows.length,
    width = rows[0]?.length ?? 0;
  let row = position(call, args[1]),
    column = args[2] === undefined ? 1 : position(call, args[2]);
  if (args[2] === undefined && height === 1) {
    column = row;
    row = 1;
  }
  if (row > height || column > width || !height || !width) fail('#REF!');
  if (row === 0 && column === 0) return rows;
  if (row === 0) return rows.map((values) => [values[column - 1] ?? null]);
  if (column === 0) return [rows[row - 1]];
  return rows[row - 1][column - 1] ?? null;
}

const LOOKUP_ARITY: Readonly<Record<string, readonly [number, number]>> = {
  MATCH: [2, 3],
  LOOKUP: [2, 3],
  VLOOKUP: [3, 4],
  HLOOKUP: [3, 4],
};

function lookupEntries(call: FunctionCall, rows: CellValue[][]): CellValue[] {
  if (call.name === 'VLOOKUP') return rows.map((row) => row[0] ?? null);
  if (call.name === 'HLOOKUP') return rows[0] ?? [];
  return flat(call.args[1]);
}

/** 0 for an exact match, 1 for the last entry not above the target, -1 for not below it. */
function lookupMode(call: FunctionCall): number {
  const { name, args } = call;
  if (name === 'MATCH') return args[2] === undefined ? 1 : integer(call, args[2], [-1, 1]);
  return name === 'LOOKUP' || args[3] === undefined || truth(args[3]) ? 1 : 0;
}

function lookupMatcher(call: FunctionCall, target: CellValue, mode: number) {
  if (mode === 1) return (entry: CellValue) => call.compare(entry, target) <= 0;
  if (mode === -1) return (entry: CellValue) => call.compare(entry, target) >= 0;
  if (typeof target !== 'string') return (entry: CellValue) => call.compare(entry, target) === 0;
  const pattern = hasWildcard(target) ? compileWildcard(target) : null;
  return (entry: CellValue) => {
    const value = checked(entry);
    return (
      typeof value === 'string' &&
      (pattern ? pattern(value) >= 0 : call.compare(value, target) === 0)
    );
  };
}

function lookupOutput(call: FunctionCall, rows: CellValue[][], found: number): CellValue {
  const { name, args } = call;
  if (name === 'MATCH') return call.math.parse(String(found + 1));
  if (name === 'LOOKUP') {
    const result = flat(args[2] ?? args[1])[found];
    if (result === undefined) fail('#N/A');
    return result ?? zeroOf(call);
  }
  const output = integer(call, args[2], [1, 16384]) - 1;
  const result = name === 'VLOOKUP' ? rows[found]?.[output] : rows[output]?.[found];
  if (result === undefined) fail('#REF!');
  return result ?? zeroOf(call);
}

function lookup(call: FunctionCall) {
  const { name, args } = call;
  arity(call, ...LOOKUP_ARITY[name]);
  const target = checked(scalar(args[0])),
    rows = matrix(args[1]);
  const entries = lookupEntries(call, rows),
    mode = lookupMode(call);
  const matches = lookupMatcher(call, target, mode);
  let found = -1;
  for (let i = 0; i < entries.length; i++) {
    if (!matches(entries[i])) continue;
    found = i;
    if (mode === 0) break;
  }
  if (found < 0) fail('#N/A');
  return lookupOutput(call, rows, found);
}

export const LOOKUP_FUNCTIONS: Readonly<Record<string, FunctionImpl>> = {
  TRANSPOSE: transpose,
  INDEX: index,
  MATCH: lookup,
  LOOKUP: lookup,
  VLOOKUP: lookup,
  HLOOKUP: lookup,
};
