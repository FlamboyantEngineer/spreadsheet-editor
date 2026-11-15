import { Decimal } from 'decimal.js';
import { calendarParts, calendarSerial } from './calendar.js';
import { formatCellValue } from './format.js';
import { char1252, code1252, formatTextMask } from './text-format.js';
import { type createQuantityArithmetic, makeQuantity, quantityInUnit } from './units.js';
import {
  CalculationError,
  type CellValue,
  isError,
  isNumeric,
  mergeNumericProvenance,
  type NumericValue,
} from './values.js';

type Value = CellValue | CellValue[][];
type Math = ReturnType<typeof createQuantityArithmetic>;
export const FUNCTION_ALIASES: Readonly<Record<string, string>> = Object.freeze({
  AVG: 'AVERAGE',
  STDEV: 'STDEV.S',
  STDEVS: 'STDEV.S',
  STDEVP: 'STDEV.P',
  VAR: 'VAR.S',
  VARS: 'VAR.S',
  VARP: 'VAR.P',
  'CEILING.MATH': 'CEILINGMATH',
  'CEILING.PRECISE': 'CEILINGPRECISE',
  'FLOOR.MATH': 'FLOORMATH',
  'FLOOR.PRECISE': 'FLOORPRECISE',
  'RANK.AVG': 'RANKAVG',
  'RANK.EQ': 'RANKEQ',
});
export const EXTENDED_FUNCTIONS = Object.freeze([
  'COUNTIF',
  'COUNTIFS',
  'SUMIF',
  'SUMIFS',
  'AVERAGEIF',
  'AVERAGEIFS',
  'MAXIFS',
  'MINIFS',
  'SUMPRODUCT',
  'MEDIAN',
  'LARGE',
  'SMALL',
  'RANK',
  'RANKAVG',
  'RANKEQ',
  'INDEX',
  'MATCH',
  'LOOKUP',
  'VLOOKUP',
  'HLOOKUP',
  'TRANSPOSE',
  'CLEAN',
  'PROPER',
  'EXACT',
  'FIND',
  'SEARCH',
  'REPLACE',
  'SUBSTITUTE',
  'REPT',
  'CHAR',
  'CODE',
  'UNICHAR',
  'UNICODE',
  'TEXTJOIN',
  'TEXT',
  'T',
  'N',
  'ISLOGICAL',
  'ISNONTEXT',
  'XOR',
  'CEILINGMATH',
  'CEILINGPRECISE',
  'FLOORMATH',
  'FLOORPRECISE',
  'MROUND',
  'EVEN',
  'ODD',
  'CONVERT',
  'DEGREES',
  'RADIANS',
  'ATAN2',
  'EDATE',
  'EOMONTH',
  'WEEKDAY',
]);
const scalar = (value: Value): CellValue =>
  Array.isArray(value) ? (value[0]?.[0] ?? null) : value;
const flat = (value: Value): CellValue[] => (Array.isArray(value) ? value.flat() : [value]);
const matrix = (value: Value): CellValue[][] => (Array.isArray(value) ? value : [[value]]);
const fail = (code: ConstructorParameters<typeof CalculationError>[0] = '#VALUE!'): never => {
  throw new CalculationError(code);
};
function checked(value: CellValue): CellValue {
  if (isError(value)) throw new CalculationError(value.code);
  return value;
}

/** Bounded dynamic programming for authored wildcards, never a backtracking regex. */
function wildcard(pattern: string, text: string, full = true): number {
  if (pattern.length * (text.length + 1) > 1000000) fail('#LIMIT!');
  const tokens: { character: string; wildcard: boolean }[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const character = pattern[i];
    if (character === '~' && /[~*?]/.test(pattern[i + 1] ?? ''))
      tokens.push({ character: pattern[++i], wildcard: false });
    else tokens.push({ character, wildcard: character === '*' || character === '?' });
  }
  let next = Array<boolean>(text.length + 1).fill(!full);
  next[text.length] = true;
  for (let index = tokens.length - 1; index >= 0; index--) {
    const token = tokens[index],
      current = Array<boolean>(text.length + 1).fill(false);
    for (let i = text.length; i >= 0; i--)
      current[i] =
        token.wildcard && token.character === '*'
          ? next[i] || (i < text.length && current[i + 1])
          : i < text.length &&
            (token.wildcard || token.character.toUpperCase() === text[i].toUpperCase()) &&
            next[i + 1];
    next = current;
  }
  return full ? (next[0] ? 0 : -1) : next.findIndex(Boolean);
}

export function evaluateExtendedFunction(
  name: string,
  args: Value[],
  math: Math,
  compare: (a: CellValue, b: CellValue) => number,
): Value | undefined {
  if (!EXTENDED_FUNCTIONS.includes(name)) return undefined;
  const zero = math.parse('0'),
    one = math.parse('1');
  const arity = (min: number, max = min) => {
    if (args.length < min || args.length > max) fail();
  };
  const number = (value: Value): NumericValue => {
    const item = checked(scalar(value));
    return isNumeric(item)
      ? item
      : item === null
        ? zero
        : typeof item === 'boolean'
          ? item
            ? one
            : zero
          : math.parse(String(item));
  };
  const integer = (value: Value, min = -1000000, max = 1000000) =>
    math.integer(number(value), min, max);
  const text = (value: Value) =>
    formatCellValue(checked(scalar(value)), { approximationMarker: false });
  const truth = (value: Value) => {
    const item = checked(scalar(value));
    return typeof item === 'boolean'
      ? item
      : item === null
        ? false
        : isNumeric(item)
          ? item.numerator !== '0'
          : String(item).toUpperCase() === 'TRUE';
  };
  const criteria = (value: CellValue, condition: CellValue) => {
    checked(value);
    checked(condition);
    if (typeof condition !== 'string') return compare(value, condition) === 0;
    const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(condition),
      operator = match?.[1] ?? '=',
      operand = match?.[2] ?? condition;
    let target: CellValue = operand;
    try {
      target = math.parse(operand);
    } catch {
      if (operand.toUpperCase() === 'TRUE' || operand.toUpperCase() === 'FALSE')
        target = operand.toUpperCase() === 'TRUE';
    }
    if (typeof target === 'string' && /[*?~]/.test(target) && ['=', '<>'].includes(operator))
      return wildcard(target, formatCellValue(value)) >= 0 === (operator === '=');
    const c = compare(value, target);
    return operator === '='
      ? c === 0
      : operator === '<>'
        ? c !== 0
        : operator === '<'
          ? c < 0
          : operator === '>'
            ? c > 0
            : operator === '<='
              ? c <= 0
              : c >= 0;
  };
  if (
    [
      'COUNTIF',
      'SUMIF',
      'AVERAGEIF',
      'COUNTIFS',
      'SUMIFS',
      'AVERAGEIFS',
      'MAXIFS',
      'MINIFS',
    ].includes(name)
  ) {
    const single = !name.endsWith('IFS'),
      count = name.startsWith('COUNT');
    arity(single ? 2 : count ? 2 : 3, single ? (count ? 2 : 3) : 255);
    const outputs = flat(single ? (args[2] ?? args[0]) : args[0]),
      pairs = single ? args.slice(0, 2) : args.slice(count ? 0 : 1);
    if (pairs.length % 2) fail();
    const tests: { values: CellValue[]; condition: CellValue }[] = [];
    for (let i = 0; i < pairs.length; i += 2) {
      const values = flat(pairs[i]);
      if (values.length !== outputs.length) fail();
      tests.push({ values, condition: scalar(pairs[i + 1]) });
    }
    const selected = outputs.filter((_, index) =>
      tests.every(({ values, condition }) => criteria(values[index], condition)),
    );
    if (count) return math.parse(String(selected.length));
    selected.forEach(checked);
    const numbers = selected.filter(isNumeric);
    if (name.startsWith('MAX') || name.startsWith('MIN'))
      return numbers.reduce(
        (a, b) => (compare(a, b) > 0 === name.startsWith('MAX') ? a : b),
        numbers[0] ?? zero,
      );
    return name.startsWith('AVERAGE') ? math.average(numbers) : numbers.reduce(math.add, zero);
  }
  if (name === 'SUMPRODUCT') {
    arity(1, 255);
    const ranges = args.map(flat),
      size = ranges[0].length;
    if (ranges.some((range) => range.length !== size)) fail();
    let sum = zero;
    for (let i = 0; i < size; i++)
      sum = math.add(
        sum,
        ranges.reduce((product, range) => {
          const value = checked(range[i]);
          return math.multiply(product, isNumeric(value) ? value : zero);
        }, one),
      );
    return sum;
  }
  if (['MEDIAN', 'LARGE', 'SMALL', 'RANK', 'RANKAVG', 'RANKEQ'].includes(name)) {
    const rank = name.startsWith('RANK');
    arity(name === 'MEDIAN' ? 1 : 2, name === 'MEDIAN' ? 255 : rank ? 3 : 2);
    const values = (name === 'MEDIAN' ? args.flatMap(flat) : flat(args[rank ? 1 : 0]))
      .map(checked)
      .filter(isNumeric)
      .sort(compare);
    if (!values.length) fail('#NUM!');
    if (name === 'MEDIAN') {
      const mid = Math.floor(values.length / 2);
      return values.length % 2 ? values[mid] : math.average([values[mid - 1], values[mid]]);
    }
    if (!rank) {
      const k = integer(args[1], 1, values.length);
      return values[name === 'SMALL' ? k - 1 : values.length - k];
    }
    const target = number(args[0]),
      ascending = args[2] !== undefined && number(args[2]).numerator !== '0';
    const positions = values
      .map((value, index) =>
        compare(value, target) === 0 ? (ascending ? index + 1 : values.length - index) : 0,
      )
      .filter(Boolean);
    if (!positions.length) fail('#N/A');
    return name === 'RANKAVG'
      ? math.divide(
          math.parse(String(positions.reduce((a, b) => a + b, 0))),
          math.parse(String(positions.length)),
        )
      : math.parse(String(Math.min(...positions)));
  }
  if (name === 'TRANSPOSE') {
    arity(1);
    const rows = matrix(args[0]);
    return Array.from({ length: rows[0]?.length ?? 0 }, (_, x) =>
      rows.map((row) => row[x] ?? null),
    );
  }
  if (name === 'INDEX') {
    arity(2, 3);
    const rows = matrix(args[0]),
      height = rows.length,
      width = rows[0]?.length ?? 0;
    const index = (value: Value) => {
      const numeric = number(value);
      if (numeric.dimension || numeric.quantityKind) fail();
      const integer = BigInt(numeric.numerator) / BigInt(numeric.denominator);
      if (integer < 0n || integer > 1048576n) fail('#REF!');
      return Number(integer);
    };
    let row = index(args[1]),
      column = args[2] === undefined ? 1 : index(args[2]);
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
  if (['MATCH', 'LOOKUP', 'VLOOKUP', 'HLOOKUP'].includes(name)) {
    arity(
      name === 'LOOKUP' ? 2 : name === 'MATCH' ? 2 : 3,
      name === 'VLOOKUP' || name === 'HLOOKUP' ? 4 : 3,
    );
    const target = checked(scalar(args[0])),
      rows = matrix(args[1]);
    const entries =
      name === 'VLOOKUP'
        ? rows.map((row) => row[0] ?? null)
        : name === 'HLOOKUP'
          ? (rows[0] ?? [])
          : flat(args[1]);
    const mode =
      name === 'MATCH'
        ? args[2] === undefined
          ? 1
          : integer(args[2], -1, 1)
        : name === 'LOOKUP' || args[3] === undefined || truth(args[3])
          ? 1
          : 0;
    let found = -1;
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i],
        c = compare(entry, target);
      if (
        mode === 0
          ? typeof target === 'string'
            ? wildcard(target, text(entry)) === 0
            : c === 0
          : mode === 1
            ? c <= 0
            : c >= 0
      ) {
        found = i;
        if (mode === 0) break;
      }
    }
    if (found < 0) fail('#N/A');
    if (name === 'MATCH') return math.parse(String(found + 1));
    if (name === 'LOOKUP') return flat(args[2] ?? args[1])[found] ?? fail('#N/A');
    const output = integer(args[2], 1, 16384) - 1;
    return name === 'VLOOKUP'
      ? (rows[found]?.[output] ?? fail('#REF!'))
      : (rows[output]?.[found] ?? fail('#REF!'));
  }
  if (
    [
      'CLEAN',
      'PROPER',
      'EXACT',
      'FIND',
      'SEARCH',
      'REPLACE',
      'SUBSTITUTE',
      'REPT',
      'CHAR',
      'CODE',
      'UNICHAR',
      'UNICODE',
      'TEXTJOIN',
      'TEXT',
      'T',
    ].includes(name)
  ) {
    if (name === 'CLEAN') {
      arity(1);
      return [...text(args[0])].filter((character) => character.charCodeAt(0) >= 32).join('');
    }
    if (name === 'PROPER') {
      arity(1);
      return text(args[0])
        .toLowerCase()
        .replace(
          /(^|[^\p{L}])(\p{L})/gu,
          (_, separator: string, letter: string) => separator + letter.toUpperCase(),
        );
    }
    if (name === 'EXACT') {
      arity(2);
      return text(args[0]) === text(args[1]);
    }
    if (name === 'T') {
      arity(1);
      const value = checked(scalar(args[0]));
      return typeof value === 'string' ? value : '';
    }
    if (name === 'REPT') {
      arity(2);
      const value = text(args[0]),
        times = integer(args[1], 0, 32767);
      if (value.length * times > 32767) fail();
      return value.repeat(times);
    }
    if (name === 'FIND' || name === 'SEARCH') {
      arity(2, 3);
      const needle = text(args[0]),
        haystack = text(args[1]),
        start = (args[2] === undefined ? 1 : integer(args[2], 1, haystack.length || 1)) - 1;
      const index =
        name === 'FIND'
          ? haystack.indexOf(needle, start)
          : wildcard(needle, haystack.slice(start), false);
      if (index < 0) fail();
      return math.parse(String((name === 'FIND' ? index : start + index) + 1));
    }
    if (name === 'CHAR' || name === 'UNICHAR') {
      arity(1);
      const code = integer(args[0], 1, name === 'CHAR' ? 255 : 1114111);
      if (code >= 55296 && code <= 57343) fail();
      return name === 'CHAR' ? char1252(code) : String.fromCodePoint(code);
    }
    if (name === 'CODE' || name === 'UNICODE') {
      arity(1);
      const value = text(args[0]);
      if (!value) fail();
      return math.parse(String(name === 'CODE' ? code1252(value) : value.codePointAt(0)));
    }
    if (name === 'REPLACE') {
      arity(4);
      const value = text(args[0]),
        start = integer(args[1], 1) - 1,
        count = integer(args[2], 0);
      return value.slice(0, start) + text(args[3]) + value.slice(start + count);
    }
    if (name === 'SUBSTITUTE') {
      arity(3, 4);
      const value = text(args[0]),
        old = text(args[1]),
        next = text(args[2]);
      if (!old) return value;
      const pieces = value.split(old);
      if (args[3] === undefined) return pieces.join(next);
      const instance = integer(args[3], 1);
      return pieces
        .map((piece, index) => (index ? (index === instance ? next : old) + piece : piece))
        .join('');
    }
    if (name === 'TEXTJOIN') {
      arity(3, 255);
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
    if (name === 'TEXT') {
      arity(2);
      return formatTextMask(number(args[0]), text(args[1]));
    }
  }
  if (['ISLOGICAL', 'ISNONTEXT', 'XOR', 'N'].includes(name)) {
    arity(1, name === 'XOR' ? 255 : 1);
    const value = checked(scalar(args[0]));
    if (name === 'ISLOGICAL') return typeof value === 'boolean';
    if (name === 'ISNONTEXT') return typeof value !== 'string';
    if (name === 'N')
      return isNumeric(value) ? value : typeof value === 'boolean' ? (value ? one : zero) : zero;
    return args.flatMap(flat).filter((value) => truth(value)).length % 2 === 1;
  }
  if (name === 'CONVERT') {
    arity(3);
    const codes: Record<string, string> = {
      C: '°C',
      F: '°F',
      K: 'K',
      hr: 'h',
      mn: 'min',
      sec: 's',
      lbm: 'lb',
      ozm: 'oz',
      meter: 'm',
      ang: 'Å',
      l: 'L',
      lt: 'L',
      cel: '°C',
      fah: '°F',
      kel: 'K',
    };
    const from = text(args[1]),
      to = text(args[2]);
    const magnitude = number(args[0]);
    if (magnitude.dimension || magnitude.quantityKind) fail();
    const value = makeQuantity(magnitude, codes[from] ?? from);
    return quantityInUnit(value, codes[to] ?? to);
  }
  if (name === 'RADIANS' || name === 'DEGREES') {
    arity(1);
    const value = number(args[0]);
    if (value.dimension || value.quantityKind) fail();
    return name === 'RADIANS'
      ? math.divide(math.multiply(value, math.pi()), math.parse('180'))
      : math.divide(math.multiply(value, math.parse('180')), math.pi());
  }
  if (name === 'ATAN2') {
    arity(2);
    const D = Decimal.clone({ precision: math.precision }),
      a = number(args[0]),
      b = number(args[1]);
    math.compare(a, b);
    if (a.numerator === '0' && b.numerator === '0') fail('#DIV/0!');
    const provenance = mergeNumericProvenance(a, b);
    return {
      ...math.parse(
        D.atan2(
          new D(b.numerator).div(b.denominator),
          new D(a.numerator).div(a.denominator),
        ).toString(),
        true,
      ),
      ...(provenance ? { provenance } : {}),
    };
  }
  if (
    [
      'CEILINGMATH',
      'CEILINGPRECISE',
      'FLOORMATH',
      'FLOORPRECISE',
      'MROUND',
      'EVEN',
      'ODD',
    ].includes(name)
  ) {
    arity(
      1,
      name === 'EVEN' || name === 'ODD' ? 1 : name === 'MROUND' || name.endsWith('PRECISE') ? 2 : 3,
    );
    const value = number(args[0]),
      unit = args[1] === undefined ? one : number(args[1]);
    if (name === 'EVEN' || name === 'ODD') {
      if (value.dimension) fail();
      const sign = math.compare(value, zero) < 0 ? -1 : 1,
        abs = sign < 0 ? math.negate(value) : value;
      let rounded = math.round(abs, 0, 'up');
      const parity = BigInt(rounded.numerator) % 2n;
      if ((name === 'EVEN' && parity !== 0n) || (name === 'ODD' && parity === 0n))
        rounded = math.add(rounded, one);
      return sign < 0 ? math.negate(rounded) : rounded;
    }
    if (unit.numerator === '0') return zero;
    if (name === 'MROUND') {
      arity(2);
      if (math.compare(value, zero) * math.compare(unit, zero) < 0) fail('#NUM!');
      return math.multiply(math.round(math.divide(value, unit), 0), unit);
    }
    const step = math.compare(unit, zero) < 0 ? math.negate(unit) : unit;
    const negative = math.compare(value, zero) < 0,
      mode = args[2] !== undefined && number(args[2]).numerator !== '0';
    const rounding = name.startsWith('CEILING')
      ? negative && mode
        ? 'floor'
        : 'ceil'
      : negative && mode
        ? 'ceil'
        : 'floor';
    return math.multiply(math.round(math.divide(value, step), 0, rounding), step);
  }
  if (['EDATE', 'EOMONTH', 'WEEKDAY'].includes(name)) {
    arity(1, name === 'WEEKDAY' ? 2 : 2);
    const value = number(args[0]),
      parts = calendarParts(value);
    if (name === 'WEEKDAY') {
      const mode = args[1] === undefined ? 1 : integer(args[1], 1, 17);
      if (![1, 2, 3, 11, 12, 13, 14, 15, 16, 17].includes(mode)) fail('#NUM!');
      const serial = Number(BigInt(value.numerator) / BigInt(value.denominator)),
        day = (serial + 6) % 7;
      const first = mode === 1 ? 0 : mode === 2 || mode === 3 || mode === 11 ? 1 : mode - 10;
      return math.parse(String(((day - first + 7) % 7) + (mode === 3 ? 0 : 1)));
    }
    arity(2);
    const months = integer(args[1], -120000, 120000);
    const monthIndex = parts.year * 12 + parts.month - 1 + months;
    const year = Math.floor(monthIndex / 12),
      month = (((monthIndex % 12) + 12) % 12) + 1;
    // DATE's public year shorthand must not reinterpret an internal 1899 field
    // as 3799. Determine the month's final day without requiring serial 10000-01-01.
    const lastDay =
      year === 1900 && month === 2 ? 29 : new Date(Date.UTC(year, month, 0)).getUTCDate();
    const day = name === 'EOMONTH' ? lastDay : Math.min(parts.day, lastDay);
    if (year === 1899 && month === 12 && day === 31) return zero;
    if (year < 1900 || year > 9999) fail('#NUM!');
    return calendarSerial(year, month, day);
  }
  return undefined;
}
