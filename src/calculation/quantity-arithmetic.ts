import {
  attach,
  combineKinds,
  type Dimension,
  dimensionOf,
  keepProvenance,
  kindsFromPowers,
  nonSIKindPowers,
  raw,
  same,
  scalar,
} from './quantity-core.js';
import { type Arithmetic, createArithmetic, type NumericOptions } from './rational.js';
import {
  CalculationError,
  isPureTemperature,
  mergeNumericProvenance,
  type NumericValue,
} from './values.js';

type Temperature = NumericValue['temperature'];

const dimensionless = (v: NumericValue) => {
  if (!scalar(v)) throw new CalculationError('#VALUE!');
};
const isScalarZero = (v: NumericValue) => v.numerator === '0' && scalar(v);
const hasNonSIKind = (v: NumericValue) => Object.keys(nonSIKindPowers(v.quantityKind)).length > 0;

/** The shared dimension of two operands; a scalar zero takes the other operand's dimension. */
function compatible(a: NumericValue, b: NumericValue): Dimension {
  const ad = dimensionOf(a),
    bd = dimensionOf(b);
  if (!same(ad, bd)) {
    if (isScalarZero(a)) return bd;
    if (isScalarZero(b)) return ad;
    throw new CalculationError('#VALUE!');
  }
  if (
    (a.quantityKind && b.quantityKind && a.quantityKind !== b.quantityKind) ||
    (hasNonSIKind(a) !== hasNonSIKind(b) && !isScalarZero(a) && !isScalarZero(b))
  )
    throw new CalculationError('#VALUE!');
  mergeNumericProvenance(a, b);
  return ad;
}
const temp = (v: NumericValue): Temperature =>
  isPureTemperature(v) ? (v.temperature ?? 'absolute') : undefined;

/** Absolute minus absolute is a difference; anything with an absolute side stays absolute. */
function sumTemperature(at: Temperature, bt: Temperature, subtract: boolean): Temperature {
  if (!at && !bt) return undefined;
  if (subtract && at === 'absolute' && bt === 'absolute') return 'difference';
  return at === 'absolute' || bt === 'absolute' ? 'absolute' : 'difference';
}

const addition =
  (math: Arithmetic, subtract: boolean) =>
  (a: NumericValue, b: NumericValue): NumericValue => {
    const dimension = compatible(a, b),
      at = temp(a),
      bt = temp(b);
    if (
      (!subtract && at === 'absolute' && bt === 'absolute') ||
      (subtract && at === 'difference' && bt === 'absolute')
    )
      throw new CalculationError('#VALUE!');
    return attach(subtract ? math.subtract(raw(a), raw(b)) : math.add(raw(a), raw(b)), {
      dimension,
      quantityKind: a.quantityKind ?? b.quantityKind,
      temperature: sumTemperature(at, bt, subtract),
      provenance: mergeNumericProvenance(a, b),
    });
  };

const averageOf =
  (math: Arithmetic) =>
  (values: readonly NumericValue[]): NumericValue => {
    const first = values[0];
    if (!first) throw new CalculationError('#DIV/0!');
    values.forEach((v) => {
      // A synthetic zero can seed a sum, but an authored scalar zero is not
      // a measurement in the first value's unit (and vice versa).
      if (!same(dimensionOf(first), dimensionOf(v)) || scalar(first) !== scalar(v))
        throw new CalculationError('#VALUE!');
      compatible(first, v);
      if (temp(first) !== temp(v)) throw new CalculationError('#VALUE!');
    });
    const total = values.reduce((sum, v) => math.add(sum, raw(v)), math.parse('0'));
    return attach(math.divide(total, math.parse(String(values.length))), {
      dimension: dimensionOf(first),
      quantityKind: first.quantityKind,
      temperature: temp(first),
      provenance: mergeNumericProvenance(...values),
    });
  };

const product =
  (math: Arithmetic) =>
  (a: NumericValue, b: NumericValue): NumericValue => {
    const kept = scalar(b) ? a : scalar(a) ? b : undefined;
    return attach(math.multiply(raw(a), raw(b)), {
      dimension: dimensionOf(a).map((n, i) => n + dimensionOf(b)[i]) as unknown as Dimension,
      quantityKind: kept ? kept.quantityKind : combineKinds(a, b),
      temperature: kept ? temp(kept) : undefined,
      provenance: mergeNumericProvenance(a, b),
    });
  };

const quotient =
  (math: Arithmetic) =>
  (a: NumericValue, b: NumericValue): NumericValue =>
    attach(math.divide(raw(a), raw(b)), {
      dimension: dimensionOf(a).map((n, i) => n - dimensionOf(b)[i]) as unknown as Dimension,
      quantityKind: scalar(b) ? a.quantityKind : combineKinds(a, b, true),
      temperature: scalar(b) ? temp(a) : undefined,
      provenance: mergeNumericProvenance(a, b),
    });

/** A dimension power times a rational exponent; it must stay a whole number. */
function scaledPower(power: number, b: NumericValue): number {
  const result = BigInt(power) * BigInt(b.numerator);
  if (result % BigInt(b.denominator)) throw new CalculationError('#VALUE!');
  return Number(result / BigInt(b.denominator));
}

const powerOf =
  (math: Arithmetic) =>
  (a: NumericValue, b: NumericValue): NumericValue => {
    dimensionless(b);
    const dimension = dimensionOf(a).map((n) => {
      const exponent = scaledPower(n, b);
      if (Math.abs(exponent) > 32) throw new CalculationError('#LIMIT!');
      return exponent;
    }) as unknown as Dimension;
    const result = math.power(raw(a), raw(b)),
      unit = b.numerator === b.denominator;
    const kinds = () =>
      Object.entries(nonSIKindPowers(a.quantityKind)).map(([kind, power]) => [
        kind,
        scaledPower(power, b),
      ]);
    return attach(result, {
      dimension,
      quantityKind: unit ? a.quantityKind : kindsFromPowers(Object.fromEntries(kinds())),
      temperature: unit ? temp(a) : undefined,
      provenance: mergeNumericProvenance(a, b),
    });
  };

const squareRoot =
  (math: Arithmetic) =>
  (a: NumericValue): NumericValue => {
    if (dimensionOf(a).some((n) => n % 2)) throw new CalculationError('#VALUE!');
    const result = math.sqrt(raw(a));
    const kinds = Object.entries(nonSIKindPowers(a.quantityKind)).map(([kind, power]) => {
      if (power % 2) throw new CalculationError('#VALUE!');
      return [kind, power / 2];
    });
    return attach(result, {
      dimension: dimensionOf(a).map((n) => n / 2) as unknown as Dimension,
      quantityKind: kindsFromPowers(Object.fromEntries(kinds)),
      provenance: mergeNumericProvenance(a),
    });
  };

/** Trigonometry accepts a plain angle; every other function needs a dimensionless argument. */
const unaryOf =
  (math: Arithmetic) =>
  (...args: Parameters<Arithmetic['unary']>): NumericValue => {
    const [name, value] = args;
    if (
      ['SIN', 'COS', 'TAN'].includes(name) &&
      value.quantityKind === 'angle' &&
      !dimensionOf(value).some(Boolean)
    )
      return keepProvenance(math.unary(name, raw(value)), value);
    dimensionless(value);
    return keepProvenance(math.unary(...args), value);
  };

/** Dimensional and physical-kind checks wrap the exact rational kernel. */
export function createQuantityArithmetic(options: NumericOptions = {}) {
  const math = createArithmetic(options);
  return {
    ...math,
    integer: (...args: Parameters<typeof math.integer>) => {
      dimensionless(args[0]);
      return math.integer(...args);
    },
    average: averageOf(math),
    add: addition(math, false),
    subtract: addition(math, true),
    compare: (a: NumericValue, b: NumericValue) => {
      compatible(a, b);
      if (temp(a) && temp(b) && temp(a) !== temp(b)) throw new CalculationError('#VALUE!');
      return math.compare(a, b);
    },
    multiply: product(math),
    divide: quotient(math),
    negate: (a: NumericValue) => {
      if (temp(a) === 'absolute') throw new CalculationError('#VALUE!');
      return attach(math.negate(raw(a)), {
        dimension: dimensionOf(a),
        quantityKind: a.quantityKind,
        temperature: temp(a),
        provenance: mergeNumericProvenance(a),
      });
    },
    power: powerOf(math),
    sqrt: squareRoot(math),
    round: (...args: Parameters<typeof math.round>) => {
      dimensionless(args[0]);
      return keepProvenance(math.round(...args), args[0]);
    },
    unary: unaryOf(math),
  };
}
