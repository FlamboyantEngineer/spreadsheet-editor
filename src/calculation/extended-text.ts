import {
  arity,
  checked,
  fail,
  flat,
  type FunctionCall,
  type FunctionImpl,
  integer,
  number,
  oneOf,
  scalar,
  text,
  truth,
  zeroOf,
} from './function-call.js';
import { char1252, code1252, formatTextMask } from './text-format.js';
import { isNumeric } from './values.js';
import { wildcard } from './wildcard.js';

/** A function of one text argument. */
const unaryText =
  (transform: (value: string) => string): FunctionImpl =>
  (call) => {
    arity(call, 1);
    return transform(text(call.args[0]));
  };

const clean = (value: string) =>
  [...value].filter((character) => character.charCodeAt(0) >= 32).join('');
const proper = (value: string) =>
  value
    .toLowerCase()
    .replace(
      /(^|[^\p{L}])(\p{L})/gu,
      (_, separator: string, letter: string) => separator + letter.toUpperCase(),
    );

function exact(call: FunctionCall) {
  arity(call, 2);
  return text(call.args[0]) === text(call.args[1]);
}

function textOnly(call: FunctionCall) {
  arity(call, 1);
  const value = checked(scalar(call.args[0]));
  return typeof value === 'string' ? value : '';
}

function repeat(call: FunctionCall) {
  arity(call, 2);
  const value = text(call.args[0]),
    times = integer(call, call.args[1], [0, 32767]);
  if (value.length * times > 32767) fail();
  return value.repeat(times);
}

function find(call: FunctionCall) {
  arity(call, 2, 3);
  const { args } = call;
  const needle = text(args[0]),
    haystack = text(args[1]),
    start = (args[2] === undefined ? 1 : integer(call, args[2], [1, haystack.length || 1])) - 1;
  const index =
    call.name === 'FIND'
      ? haystack.indexOf(needle, start)
      : wildcard(needle, haystack.slice(start), false);
  if (index < 0) fail();
  return call.math.parse(String((call.name === 'FIND' ? index : start + index) + 1));
}

function character(call: FunctionCall) {
  arity(call, 1);
  const code = integer(call, call.args[0], [1, call.name === 'CHAR' ? 255 : 1114111]);
  if (code >= 55296 && code <= 57343) fail();
  return call.name === 'CHAR' ? char1252(code) : String.fromCodePoint(code);
}

function code(call: FunctionCall) {
  arity(call, 1);
  const value = text(call.args[0]);
  if (!value) fail();
  return call.math.parse(String(call.name === 'CODE' ? code1252(value) : value.codePointAt(0)));
}

function replace(call: FunctionCall) {
  arity(call, 4);
  const { args } = call;
  const value = text(args[0]),
    start = integer(call, args[1], [1, 1000000]) - 1,
    count = integer(call, args[2], [0, 1000000]);
  return value.slice(0, start) + text(args[3]) + value.slice(start + count);
}

function substitute(call: FunctionCall) {
  arity(call, 3, 4);
  const { args } = call;
  const value = text(args[0]),
    old = text(args[1]),
    next = text(args[2]);
  if (!old) return value;
  const pieces = value.split(old);
  if (args[3] === undefined) return pieces.join(next);
  const instance = integer(call, args[3], [1, 1000000]);
  return pieces
    .map((piece, index) => (index ? (index === instance ? next : old) + piece : piece))
    .join('');
}

function textJoin(call: FunctionCall) {
  arity(call, 3, 255);
  const { args } = call;
  const delimiter = text(args[0]),
    skip = truth(args[1]),
    parts = args
      .slice(2)
      .flatMap(flat)
      .filter((value) => !skip || (value !== null && value !== ''));
  const result = parts.map((value) => text(value)).join(delimiter);
  if (result.length > 32767) fail();
  return result;
}

function textMask(call: FunctionCall) {
  arity(call, 2);
  return formatTextMask(number(call, call.args[0]), text(call.args[1]));
}

function information(call: FunctionCall) {
  const { name, args } = call;
  arity(call, 1, name === 'XOR' ? 255 : 1);
  const value = checked(scalar(args[0]));
  if (name === 'ISLOGICAL') return typeof value === 'boolean';
  if (name === 'ISNONTEXT') return typeof value !== 'string';
  if (name === 'N') {
    if (isNumeric(value)) return value;
    return value === true ? oneOf(call) : zeroOf(call);
  }
  return args.flatMap(flat).filter((item) => truth(item)).length % 2 === 1;
}

export const TEXT_FUNCTIONS: Readonly<Record<string, FunctionImpl>> = {
  CLEAN: unaryText(clean),
  PROPER: unaryText(proper),
  EXACT: exact,
  T: textOnly,
  REPT: repeat,
  FIND: find,
  SEARCH: find,
  CHAR: character,
  UNICHAR: character,
  CODE: code,
  UNICODE: code,
  REPLACE: replace,
  SUBSTITUTE: substitute,
  TEXTJOIN: textJoin,
  TEXT: textMask,
  ISLOGICAL: information,
  ISNONTEXT: information,
  XOR: information,
  N: information,
};
