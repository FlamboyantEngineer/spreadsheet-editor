import { Decimal } from 'decimal.js';
import { CalculationError, type NumericValue } from './values.js';

export interface NumericOptions {
  /** Significant digits used only for non-rational functions. */
  precision?: number;
  /** Resource limit, not an arithmetic rounding precision. */
  maxIntegerDigits?: number;
}

export function createArithmetic(options: NumericOptions = {}) {
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
  const D = Decimal.clone({ precision, rounding: Decimal.ROUND_HALF_EVEN });
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
  function add(a: NumericValue, b: NumericValue): NumericValue {
    return value(
      BigInt(a.numerator) * BigInt(b.denominator) + BigInt(b.numerator) * BigInt(a.denominator),
      BigInt(a.denominator) * BigInt(b.denominator),
      a.approximate || b.approximate,
    );
  }
  function negate(a: NumericValue): NumericValue {
    return value(-BigInt(a.numerator), BigInt(a.denominator), a.approximate);
  }
  function subtract(a: NumericValue, b: NumericValue): NumericValue {
    return add(a, negate(b));
  }
  function multiply(a: NumericValue, b: NumericValue): NumericValue {
    return value(
      BigInt(a.numerator) * BigInt(b.numerator),
      BigInt(a.denominator) * BigInt(b.denominator),
      a.approximate || b.approximate,
    );
  }
  function divide(a: NumericValue, b: NumericValue): NumericValue {
    return value(
      BigInt(a.numerator) * BigInt(b.denominator),
      BigInt(a.denominator) * BigInt(b.numerator),
      a.approximate || b.approximate,
    );
  }
  function compare(a: NumericValue, b: NumericValue): number {
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
  const decimal = (a: NumericValue) => new D(a.numerator).div(a.denominator);
  function approximate(result: Decimal): NumericValue {
    if (!result.isFinite()) throw new CalculationError('#NUM!');
    return parse(result.toString(), true);
  }
  function power(a: NumericValue, b: NumericValue): NumericValue {
    if (a.numerator === '0' && b.numerator === '0') throw new CalculationError('#NUM!');
    if (b.denominator !== '1') return approximate(decimal(a).pow(decimal(b)));
    const exponent = integer(b);
    if (Math.abs(exponent) * Math.max(a.numerator.length - 1, a.denominator.length - 1) > maxDigits)
      throw new CalculationError('#LIMIT!');
    return exponent >= 0
      ? value(
          BigInt(a.numerator) ** BigInt(exponent),
          BigInt(a.denominator) ** BigInt(exponent),
          a.approximate || b.approximate,
        )
      : value(
          BigInt(a.denominator) ** BigInt(-exponent),
          BigInt(a.numerator) ** BigInt(-exponent),
          a.approximate || b.approximate,
        );
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
  function sqrt(a: NumericValue): NumericValue {
    const n = integerRoot(BigInt(a.numerator)),
      d = integerRoot(BigInt(a.denominator));
    return n * n === BigInt(a.numerator) && d * d === BigInt(a.denominator)
      ? value(n, d, a.approximate)
      : approximate(decimal(a).sqrt());
  }
  /** Explicit ROUND-family operations change a value; display formatting never does. */
  function round(
    a: NumericValue,
    places: number,
    mode: 'nearest' | 'down' | 'up' | 'floor' | 'ceil' = 'nearest',
  ): NumericValue {
    if (!Number.isSafeInteger(places) || Math.abs(places) > maxDigits)
      throw new CalculationError('#NUM!');
    const scale = 10n ** BigInt(Math.abs(places));
    const n = BigInt(a.numerator) * (places >= 0 ? scale : 1n);
    const d = BigInt(a.denominator) * (places >= 0 ? 1n : scale);
    let quotient = n / d;
    const remainder = n % d,
      sign = n < 0n ? -1n : 1n;
    if (
      remainder !== 0n &&
      (mode === 'up' ||
        (mode === 'floor' && n < 0n) ||
        (mode === 'ceil' && n > 0n) ||
        (mode === 'nearest' && abs(remainder) * 2n >= d))
    )
      quotient += sign;
    return places >= 0
      ? value(quotient, scale, a.approximate)
      : value(quotient * scale, 1n, a.approximate);
  }
  function unary(name: string, a: NumericValue): NumericValue {
    const input = decimal(a);
    switch (name) {
      case 'SIN':
        return approximate(input.sin());
      case 'COS':
        return approximate(input.cos());
      case 'TAN':
        return approximate(input.tan());
      case 'ASIN':
        return approximate(input.asin());
      case 'ACOS':
        return approximate(input.acos());
      case 'ATAN':
        return approximate(input.atan());
      case 'SINH':
        return approximate(input.sinh());
      case 'COSH':
        return approximate(input.cosh());
      case 'TANH':
        return approximate(input.tanh());
      case 'ASINH':
        return approximate(input.asinh());
      case 'ACOSH':
        return approximate(input.acosh());
      case 'ATANH':
        return approximate(input.atanh());
      case 'EXPM1': {
        if (input.abs().gte('0.1')) return approximate(input.exp().minus(1));
        let term = input,
          sum = input;
        for (let k = 2; k <= 20000; k++) {
          term = term.mul(input).div(k);
          const next = sum.plus(term);
          if (next.eq(sum)) return approximate(next);
          sum = next;
        }
        throw new CalculationError('#LIMIT!');
      }
      case 'LOG1P': {
        if (input.abs().gte('0.1')) return approximate(input.plus(1).ln());
        const u = input.div(input.plus(2)),
          squared = u.mul(u);
        let power = u,
          sum = u;
        for (let k = 3; k <= 40001; k += 2) {
          power = power.mul(squared);
          const next = sum.plus(power.div(k));
          if (next.eq(sum)) return approximate(next.mul(2));
          sum = next;
        }
        throw new CalculationError('#LIMIT!');
      }
      case 'EXP':
        return approximate(input.exp());
      case 'LN':
        return approximate(input.ln());
      case 'LOG10':
        return approximate(input.log(10));
      default:
        throw new CalculationError('#NAME?');
    }
  }
  return {
    parse,
    value,
    add,
    subtract,
    multiply,
    divide,
    compare,
    negate,
    integer,
    power,
    sqrt,
    round,
    unary,
    pi: () => approximate(D.acos(-1)),
    precision,
  };
}

export type Arithmetic = ReturnType<typeof createArithmetic>;
