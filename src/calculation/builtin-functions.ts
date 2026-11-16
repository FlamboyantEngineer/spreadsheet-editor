import { AGGREGATE_FUNCTIONS } from './aggregate-functions.js';
import { type EagerCall, type EagerFunction, requireArity } from './builtin-call.js';
import { CALENDAR_FUNCTIONS } from './calendar-functions.js';
import { CalculationError, isNumeric, type NumericValue } from './values.js';

const constant =
  (value: boolean): EagerFunction =>
  (call) => {
    requireArity(call, 0);
    return value;
  };

const notAvailable: EagerFunction = (call) => {
  requireArity(call, 0);
  throw new CalculationError('#N/A');
};

const logical: EagerFunction = (call) => {
  requireArity(call, 1, 255);
  const { truth } = call.evaluation.values;
  return call.name === 'AND' ? call.flat.every(truth) : call.flat.some(truth);
};

const not: EagerFunction = (call) => {
  requireArity(call, 1);
  return !call.evaluation.values.truth(call.args[0]);
};

const valueKind: EagerFunction = (call) => {
  requireArity(call, 1);
  const { values } = call.evaluation;
  const value = values.checked(values.scalar(call.args[0]));
  if (call.name === 'ISNUMBER') return isNumeric(value);
  return call.name === 'ISTEXT' ? typeof value === 'string' : value === null;
};

const size: EagerFunction = (call) => {
  requireArity(call, 1);
  const range = call.args[0];
  let length = 1;
  if (Array.isArray(range)) length = call.name === 'ROWS' ? range.length : (range[0]?.length ?? 0);
  return call.evaluation.math.parse(String(length));
};

const concat: EagerFunction = (call) => {
  requireArity(call, 1, 255);
  return call.flat.map(call.evaluation.values.text).join('');
};

/** A function of one text argument. */
const textOf =
  (transform: (call: EagerCall, source: string) => ReturnType<EagerFunction>): EagerFunction =>
  (call) => {
    requireArity(call, 1);
    return transform(call, call.evaluation.values.text(call.args[0]));
  };

/** A count of characters. */
const countAt = (call: EagerCall, index: number) =>
  call.evaluation.math.integer(call.evaluation.values.numeric(call.args[index]), 0, 1_000_000);

const edge: EagerFunction = (call) => {
  requireArity(call, 1, 2);
  const source = call.evaluation.values.text(call.args[0]),
    count = call.args[1] === undefined ? 1 : countAt(call, 1);
  if (call.name === 'LEFT') return source.slice(0, count);
  return count === 0 ? '' : source.slice(-count);
};

const mid: EagerFunction = (call) => {
  requireArity(call, 3);
  const source = call.evaluation.values.text(call.args[0]),
    start = countAt(call, 1);
  if (start < 1) throw new CalculationError('#VALUE!');
  return source.slice(start - 1, start - 1 + countAt(call, 2));
};

const pi: EagerFunction = (call) => {
  requireArity(call, 0);
  return call.evaluation.math.pi();
};

const ROUNDING = {
  INT: 'floor',
  ROUNDUP: 'up',
  ROUND: 'nearest',
  ROUNDDOWN: 'down',
  TRUNC: 'down',
} as const;

const round: EagerFunction = (call) => {
  const single = call.name === 'TRUNC' || call.name === 'INT';
  requireArity(call, single ? 1 : 2, call.name === 'INT' ? 1 : 2);
  const { math, values } = call.evaluation;
  const places = call.args[1] === undefined ? 0 : math.integer(values.numeric(call.args[1]));
  return math.round(
    values.numeric(call.args[0]),
    places,
    ROUNDING[call.name as keyof typeof ROUNDING],
  );
};

/** LOG of an exact power is exact: an integer candidate is verified rationally. */
function logarithm(call: EagerCall, a: NumericValue, b: NumericValue): NumericValue {
  const { math } = call.evaluation;
  const ratio = math.divide(math.unary('LN', a), math.unary('LN', b));
  const candidate = math.round(ratio, 0);
  try {
    math.integer(candidate, -10000, 10000);
    if (math.compare(math.power(b, candidate), a) !== 0) return ratio;
    const exact = { ...candidate, approximate: a.approximate || b.approximate };
    if (!exact.approximate) delete exact.precision;
    return exact;
  } catch (error) {
    if (!(error instanceof CalculationError)) throw error;
    return ratio;
  }
}

const QUOTIENT_ROUNDING = {
  MOD: 'floor',
  FLOOR: 'floor',
  CEILING: 'ceil',
  QUOTIENT: 'down',
} as const;

const binaryMath: EagerFunction = (call) => {
  const { name, args } = call;
  requireArity(call, name === 'LOG' ? 1 : 2, 2);
  const { math, values } = call.evaluation;
  const a = values.numeric(args[0]),
    b = args[1] === undefined ? math.parse('10') : values.numeric(args[1]);
  if (name === 'POWER') return math.power(a, b);
  if (name === 'LOG') return logarithm(call, a, b);
  const quotient = math.round(
    math.divide(a, b),
    0,
    QUOTIENT_ROUNDING[name as keyof typeof QUOTIENT_ROUNDING],
  );
  if (name === 'MOD') return math.subtract(a, math.multiply(b, quotient));
  return name === 'QUOTIENT' ? quotient : math.multiply(b, quotient);
};

const unaryMath: EagerFunction = (call) => {
  requireArity(call, 1);
  const { math, values } = call.evaluation;
  const value = values.numeric(call.args[0]);
  if (call.name === 'ABS') return math.compare(value, values.zero) < 0 ? math.negate(value) : value;
  if (call.name === 'SIGN') return math.parse(String(math.compare(value, values.zero)));
  return call.name === 'SQRT' ? math.sqrt(value) : math.unary(call.name, value);
};

const UNARY_MATH = [
  'ABS',
  'SIGN',
  'SQRT',
  'SIN',
  'COS',
  'TAN',
  'ASIN',
  'ACOS',
  'ATAN',
  'ACOSH',
  'ASINH',
  'ATANH',
  'COSH',
  'SINH',
  'TANH',
  'EXPM1',
  'LOG1P',
  'EXP',
  'LN',
  'LOG10',
];

/** The original catalog's functions that take evaluated arguments, by canonical name. */
export const BUILTIN_FUNCTIONS: Readonly<Record<string, EagerFunction>> = {
  ...CALENDAR_FUNCTIONS,
  ...AGGREGATE_FUNCTIONS,
  TRUE: constant(true),
  FALSE: constant(false),
  NA: notAvailable,
  AND: logical,
  OR: logical,
  NOT: not,
  ISNUMBER: valueKind,
  ISTEXT: valueKind,
  ISBLANK: valueKind,
  ROWS: size,
  COLUMNS: size,
  CONCAT: concat,
  CONCATENATE: concat,
  VALUE: textOf((call, source) => call.evaluation.math.parse(source)),
  LEN: textOf((call, source) => call.evaluation.math.parse(String(source.length))),
  LOWER: textOf((_, source) => source.toLowerCase()),
  UPPER: textOf((_, source) => source.toUpperCase()),
  TRIM: textOf((_, source) => source.trim().replace(/ +/g, ' ')),
  LEFT: edge,
  RIGHT: edge,
  MID: mid,
  PI: pi,
  ROUND: round,
  ROUNDUP: round,
  ROUNDDOWN: round,
  TRUNC: round,
  INT: round,
  POWER: binaryMath,
  MOD: binaryMath,
  QUOTIENT: binaryMath,
  CEILING: binaryMath,
  FLOOR: binaryMath,
  LOG: binaryMath,
  ...Object.fromEntries(UNARY_MATH.map((name) => [name, unaryMath])),
};
