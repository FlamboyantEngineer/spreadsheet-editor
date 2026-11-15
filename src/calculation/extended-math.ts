import { Decimal } from 'decimal.js';
import { calendarParts, calendarSerial } from './calendar.js';
import {
  arity,
  fail,
  flagged,
  type FunctionCall,
  type FunctionImpl,
  integer,
  number,
  oneOf,
  text,
  zeroOf,
} from './function-call.js';
import { makeQuantity, quantityInUnit } from './units.js';
import { mergeNumericProvenance, type NumericValue } from './values.js';

const CONVERT_CODES: Readonly<Record<string, string>> = {
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

function scalarNumber(call: FunctionCall, index: number): NumericValue {
  const value = number(call, call.args[index]);
  if (value.dimension || value.quantityKind) fail();
  return value;
}

function convert(call: FunctionCall) {
  arity(call, 3);
  const from = text(call.args[1]),
    to = text(call.args[2]);
  const value = makeQuantity(scalarNumber(call, 0), CONVERT_CODES[from] ?? from);
  return quantityInUnit(value, CONVERT_CODES[to] ?? to);
}

function angle(call: FunctionCall) {
  arity(call, 1);
  const { math } = call;
  const value = scalarNumber(call, 0);
  return call.name === 'RADIANS'
    ? math.divide(math.multiply(value, math.pi()), math.parse('180'))
    : math.divide(math.multiply(value, math.parse('180')), math.pi());
}

function atan2(call: FunctionCall) {
  arity(call, 2);
  const { math } = call;
  const D = Decimal.clone({ precision: math.precision }),
    a = number(call, call.args[0]),
    b = number(call, call.args[1]);
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

function evenOdd(call: FunctionCall) {
  arity(call, 1);
  const { math } = call;
  const value = number(call, call.args[0]);
  if (value.dimension) fail();
  const negative = math.compare(value, zeroOf(call)) < 0,
    abs = negative ? math.negate(value) : value;
  let rounded = math.round(abs, 0, 'up');
  const even = BigInt(rounded.numerator) % 2n === 0n;
  if (even !== (call.name === 'EVEN')) rounded = math.add(rounded, oneOf(call));
  return negative ? math.negate(rounded) : rounded;
}

function multipleRound(call: FunctionCall) {
  arity(call, 1, 2);
  const { math } = call;
  const value = number(call, call.args[0]),
    unit = call.args[1] === undefined ? oneOf(call) : number(call, call.args[1]),
    zero = zeroOf(call);
  if (unit.numerator === '0') return zero;
  arity(call, 2);
  if (math.compare(value, zero) * math.compare(unit, zero) < 0) fail('#NUM!');
  return math.multiply(math.round(math.divide(value, unit), 0), unit);
}

/** CEILING.MATH, CEILING.PRECISE, FLOOR.MATH and FLOOR.PRECISE. */
function stepRound(call: FunctionCall) {
  const { name, args, math } = call;
  arity(call, 1, name.endsWith('PRECISE') ? 2 : 3);
  const value = number(call, args[0]),
    unit = args[1] === undefined ? oneOf(call) : number(call, args[1]),
    zero = zeroOf(call);
  if (unit.numerator === '0') return zero;
  const step = math.compare(unit, zero) < 0 ? math.negate(unit) : unit;
  const mode = flagged(call, args[2]),
    away = math.compare(value, zero) < 0 && mode;
  const ceiling = name.startsWith('CEILING') !== away;
  return math.multiply(math.round(math.divide(value, step), 0, ceiling ? 'ceil' : 'floor'), step);
}

const WEEKDAY_MODES = [1, 2, 3, 11, 12, 13, 14, 15, 16, 17];

function weekday(call: FunctionCall) {
  arity(call, 1, 2);
  const { args, math } = call;
  const value = number(call, args[0]);
  calendarParts(value);
  const mode = args[1] === undefined ? 1 : integer(call, args[1], [1, 17]);
  if (!WEEKDAY_MODES.includes(mode)) fail('#NUM!');
  const serial = Number(BigInt(value.numerator) / BigInt(value.denominator)),
    day = (serial + 6) % 7;
  const first = mode === 1 ? 0 : mode === 2 || mode === 3 || mode === 11 ? 1 : mode - 10;
  return math.parse(String(((day - first + 7) % 7) + (mode === 3 ? 0 : 1)));
}

/** EDATE and EOMONTH: a date some months away, or the end of that month. */
function monthShift(call: FunctionCall) {
  arity(call, 1, 2);
  const parts = calendarParts(number(call, call.args[0]));
  arity(call, 2);
  const months = integer(call, call.args[1], [-120000, 120000]);
  const monthIndex = parts.year * 12 + parts.month - 1 + months;
  const year = Math.floor(monthIndex / 12),
    month = (((monthIndex % 12) + 12) % 12) + 1;
  // DATE's public year shorthand must not reinterpret an internal 1899 field
  // as 3799. Determine the month's final day without requiring serial 10000-01-01.
  const lastDay =
    year === 1900 && month === 2 ? 29 : new Date(Date.UTC(year, month, 0)).getUTCDate();
  const day = call.name === 'EOMONTH' ? lastDay : Math.min(parts.day, lastDay);
  if (year === 1899 && month === 12 && day === 31) return zeroOf(call);
  if (year < 1900 || year > 9999) fail('#NUM!');
  return calendarSerial(year, month, day);
}

export const MATH_FUNCTIONS: Readonly<Record<string, FunctionImpl>> = {
  CONVERT: convert,
  RADIANS: angle,
  DEGREES: angle,
  ATAN2: atan2,
  EVEN: evenOdd,
  ODD: evenOdd,
  MROUND: multipleRound,
  CEILINGMATH: stepRound,
  CEILINGPRECISE: stepRound,
  FLOORMATH: stepRound,
  FLOORPRECISE: stepRound,
  EDATE: monthShift,
  EOMONTH: monthShift,
  WEEKDAY: weekday,
};
