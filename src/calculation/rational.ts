import { Decimal } from 'decimal.js';
import { CalculationError, mergeNumericProvenance, type NumericValue } from './values.js';

export interface NumericOptions {
  /** Significant digits used only for non-rational functions. */
  precision?: number;
  /** Resource limit, not an arithmetic rounding precision. */
  maxIntegerDigits?: number;
  /** Architecture context; never assume a machine word size. */
  wordSizeBits?: string;
}

type RoundingMode = 'nearest' | 'down' | 'up' | 'floor' | 'ceil';
type DecimalType = InstanceType<typeof Decimal>;

const abs = (n: bigint) => (n < 0n ? -n : n);
function gcd(a: bigint, b: bigint): bigint {
  a = abs(a);
  b = abs(b);
  while (b) {
    const remainder = a % b;
    a = b;
    b = remainder;
  }
  return a || 1n;
}
function integerRoot(n: bigint): bigint {
  if (n < 0n) throw new CalculationError('#NUM!');
  if (n < 2n) return n;
  let root = 1n << BigInt(Math.ceil(n.toString(2).length / 2));
  for (;;) {
    const next = (root + n / root) / 2n;
    if (next >= root) return root;
    root = next;
  }
}

/** e^x - 1, by its series near zero, where subtracting from e^x would lose digits. */
function expm1(input: DecimalType): DecimalType {
  if (input.abs().gte('0.1')) return input.exp().minus(1);
  let term = input,
    sum = input;
  for (let k = 2; k <= 20000; k++) {
    term = term.mul(input).div(k);
    const next = sum.plus(term);
    if (next.eq(sum)) return next;
    sum = next;
  }
  throw new CalculationError('#LIMIT!');
}
/** ln(1 + x), by the series of atanh near zero. */
function log1p(input: DecimalType): DecimalType {
  if (input.abs().gte('0.1')) return input.plus(1).ln();
  const u = input.div(input.plus(2)),
    squared = u.mul(u);
  let power = u,
    sum = u;
  for (let k = 3; k <= 40001; k += 2) {
    power = power.mul(squared);
    const next = sum.plus(power.div(k));
    if (next.eq(sum)) return next.mul(2);
    sum = next;
  }
  throw new CalculationError('#LIMIT!');
}
const DECIMAL_FUNCTIONS: Readonly<Record<string, (x: DecimalType) => DecimalType>> = {
  SIN: (x) => x.sin(),
  COS: (x) => x.cos(),
  TAN: (x) => x.tan(),
  ASIN: (x) => x.asin(),
  ACOS: (x) => x.acos(),
  ATAN: (x) => x.atan(),
  SINH: (x) => x.sinh(),
  COSH: (x) => x.cosh(),
  TANH: (x) => x.tanh(),
  ASINH: (x) => x.asinh(),
  ACOSH: (x) => x.acosh(),
  ATANH: (x) => x.atanh(),
  EXPM1: expm1,
  LOG1P: log1p,
  EXP: (x) => x.exp(),
  LN: (x) => x.ln(),
  LOG10: (x) => x.log(10),
};

/** Whether rounding moves the truncated quotient one step away from zero. */
function stepsAway(
  mode: RoundingMode,
  { n, d, remainder }: { n: bigint; d: bigint; remainder: bigint },
) {
  if (remainder === 0n) return false;
  if (mode === 'up') return true;
  if (mode === 'floor') return n < 0n;
  if (mode === 'ceil') return n > 0n;
  return mode === 'nearest' && abs(remainder) * 2n >= d;
}

function numericLimits(options: NumericOptions) {
  const precision = options.precision ?? 80;
  const maxDigits = options.maxIntegerDigits ?? 10_000;
  if (
    !Number.isInteger(precision) ||
    precision < 16 ||
    precision > 1000 ||
    !Number.isInteger(maxDigits) ||
    maxDigits < 32 ||
    maxDigits > 100_000
  )
    throw new RangeError('Invalid numeric limits.');
  return { precision, maxDigits };
}

/** Exact rational operations on reduced numerator/denominator strings. */
function createRationalCore({ precision, maxDigits }: { precision: number; maxDigits: number }) {
  function value(n: bigint, d = 1n, approximate = false): NumericValue {
    if (d === 0n) throw new CalculationError('#DIV/0!');
    if (d < 0n) {
      n = -n;
      d = -d;
    }
    const common = gcd(n, d);
    const numerator = String(n / common),
      denominator = String(d / common);
    if (numerator.length > maxDigits + 1 || denominator.length > maxDigits)
      throw new CalculationError('#LIMIT!');
    return approximate
      ? { kind: 'number', numerator, denominator, approximate, precision }
      : { kind: 'number', numerator, denominator, approximate };
  }
  function parse(text: string, approximate = false): NumericValue {
    const match = /^([+-]?)(\d+(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?$/.exec(text.trim());
    if (!match) throw new CalculationError('#VALUE!');
    const exponent = Number(match[3] ?? '0');
    const [integer, fraction = ''] = match[2].split('.');
    const digits = integer + fraction;
    const scale = fraction.length - exponent;
    if (!Number.isSafeInteger(exponent) || Math.abs(scale) > maxDigits || digits.length > maxDigits)
      throw new CalculationError('#LIMIT!');
    const coefficient = BigInt((match[1] === '-' ? '-' : '') + digits);
    return scale >= 0
      ? value(coefficient, 10n ** BigInt(scale), approximate)
      : value(coefficient * 10n ** BigInt(-scale), 1n, approximate);
  }
  const add = (a: NumericValue, b: NumericValue): NumericValue =>
    value(
      BigInt(a.numerator) * BigInt(b.denominator) + BigInt(b.numerator) * BigInt(a.denominator),
      BigInt(a.denominator) * BigInt(b.denominator),
      a.approximate || b.approximate,
    );
  const negate = (a: NumericValue): NumericValue =>
    value(-BigInt(a.numerator), BigInt(a.denominator), a.approximate);
  const multiply = (a: NumericValue, b: NumericValue): NumericValue =>
    value(
      BigInt(a.numerator) * BigInt(b.numerator),
      BigInt(a.denominator) * BigInt(b.denominator),
      a.approximate || b.approximate,
    );
  const divide = (a: NumericValue, b: NumericValue): NumericValue =>
    value(
      BigInt(a.numerator) * BigInt(b.denominator),
      BigInt(a.denominator) * BigInt(b.numerator),
      a.approximate || b.approximate,
    );
  function compare(a: NumericValue, b: NumericValue): number {
    mergeNumericProvenance(a, b);
    const difference =
      BigInt(a.numerator) * BigInt(b.denominator) - BigInt(b.numerator) * BigInt(a.denominator);
    return difference === 0n ? 0 : difference < 0n ? -1 : 1;
  }
  function integer(a: NumericValue, min = -maxDigits, max = maxDigits): number {
    if (
      a.denominator !== '1' ||
      BigInt(a.numerator) < BigInt(min) ||
      BigInt(a.numerator) > BigInt(max)
    )
      throw new CalculationError('#NUM!');
    return Number(a.numerator);
  }
  const subtract = (a: NumericValue, b: NumericValue) => add(a, negate(b));
  return { value, parse, add, negate, subtract, multiply, divide, compare, integer };
}
type RationalCore = ReturnType<typeof createRationalCore>;

/** Powers, roots, rounding and the functions that are approximate outside exact cases. */
function createApproximations(
  core: RationalCore,
  { precision, maxDigits }: { precision: number; maxDigits: number },
) {
  const D = Decimal.clone({ precision, rounding: Decimal.ROUND_HALF_EVEN });
  const decimal = (a: NumericValue) => new D(a.numerator).div(a.denominator);
  function approximate(result: DecimalType): NumericValue {
    if (!result.isFinite()) throw new CalculationError('#NUM!');
    return core.parse(result.toString(), true);
  }
  function power(a: NumericValue, b: NumericValue): NumericValue {
    if (a.numerator === '0' && b.numerator === '0') throw new CalculationError('#NUM!');
    if (b.denominator !== '1') return approximate(decimal(a).pow(decimal(b)));
    const exponent = core.integer(b);
    if (Math.abs(exponent) * Math.max(a.numerator.length - 1, a.denominator.length - 1) > maxDigits)
      throw new CalculationError('#LIMIT!');
    const [base, divisor] =
      exponent >= 0 ? [a.numerator, a.denominator] : [a.denominator, a.numerator];
    const e = BigInt(Math.abs(exponent));
    return core.value(BigInt(base) ** e, BigInt(divisor) ** e, a.approximate || b.approximate);
  }
  function sqrt(a: NumericValue): NumericValue {
    const n = integerRoot(BigInt(a.numerator)),
      d = integerRoot(BigInt(a.denominator));
    return n * n === BigInt(a.numerator) && d * d === BigInt(a.denominator)
      ? core.value(n, d, a.approximate)
      : approximate(decimal(a).sqrt());
  }
  /** Explicit ROUND-family operations change a value; display formatting never does. */
  function round(a: NumericValue, places: number, mode: RoundingMode = 'nearest'): NumericValue {
    if (!Number.isSafeInteger(places) || Math.abs(places) > maxDigits)
      throw new CalculationError('#NUM!');
    const scale = 10n ** BigInt(Math.abs(places));
    const n = BigInt(a.numerator) * (places >= 0 ? scale : 1n);
    const d = BigInt(a.denominator) * (places >= 0 ? 1n : scale);
    let quotient = n / d;
    if (stepsAway(mode, { n, d, remainder: n % d })) quotient += n < 0n ? -1n : 1n;
    return places >= 0
      ? core.value(quotient, scale, a.approximate)
      : core.value(quotient * scale, 1n, a.approximate);
  }
  function unary(name: string, a: NumericValue): NumericValue {
    const operation = DECIMAL_FUNCTIONS[name];
    if (!operation) throw new CalculationError('#NAME?');
    return approximate(operation(decimal(a)));
  }
  return { power, sqrt, round, unary, pi: () => approximate(D.acos(-1)) };
}

/** Wraps an operation so its result carries the merged provenance of its numeric inputs. */
const carry =
  <Args extends unknown[]>(operation: (...args: Args) => NumericValue) =>
  (...args: Args): NumericValue => {
    const inputs = args.filter((value): value is Args[number] & NumericValue =>
      Boolean(value && typeof value === 'object' && 'kind' in value && value.kind === 'number'),
    );
    const provenance = mergeNumericProvenance(...inputs);
    const result = operation(...args);
    return provenance ? { ...result, provenance } : result;
  };

export function createArithmetic(options: NumericOptions = {}) {
  const limits = numericLimits(options);
  const core = createRationalCore(limits),
    approximations = createApproximations(core, limits);
  return {
    parse: core.parse,
    value: core.value,
    add: carry(core.add),
    subtract: carry(core.subtract),
    multiply: carry(core.multiply),
    divide: carry(core.divide),
    compare: core.compare,
    negate: carry(core.negate),
    integer: core.integer,
    power: carry(approximations.power),
    sqrt: carry(approximations.sqrt),
    round: carry(approximations.round),
    unary: carry(approximations.unary),
    pi: approximations.pi,
    precision: limits.precision,
  };
}

export type Arithmetic = ReturnType<typeof createArithmetic>;
