import { createArithmetic, type NumericOptions } from './rational.js';
import { CalculationError, type NumericValue } from './values.js';

export type Dimension = readonly [number, number, number, number, number, number, number];
const none: Dimension = [0, 0, 0, 0, 0, 0, 0];
const dimensions: Record<string, Dimension> = {
  length: [1, 0, 0, 0, 0, 0, 0],
  mass: [0, 1, 0, 0, 0, 0, 0],
  time: [0, 0, 1, 0, 0, 0, 0],
  temperature: [0, 0, 0, 0, 1, 0, 0],
  area: [2, 0, 0, 0, 0, 0, 0],
  volume: [3, 0, 0, 0, 0, 0, 0],
  velocity: [1, 0, -1, 0, 0, 0, 0],
  acceleration: [1, 0, -2, 0, 0, 0, 0],
  force: [1, 1, -2, 0, 0, 0, 0],
  pressure: [-1, 1, -2, 0, 0, 0, 0],
  energy: [2, 1, -2, 0, 0, 0, 0],
  power: [2, 1, -3, 0, 0, 0, 0],
  frequency: [0, 0, -1, 0, 0, 0, 0],
};
const definitions: Record<string, { dimension: Dimension; factor: string; offset?: string }> = {};
for (const [dimension, units] of Object.entries({
  length: { m: '1', mm: '0.001', cm: '0.01', km: '1000', in: '0.0254', ft: '0.3048', yd: '0.9144' },
  mass: { kg: '1', g: '0.001', mg: '0.000001', lb: '0.45359237' },
  time: { s: '1', ms: '0.001', min: '60', h: '3600' },
  force: { N: '1', kN: '1000' },
  pressure: { Pa: '1', kPa: '1000', MPa: '1000000', GPa: '1000000000' },
  energy: { J: '1', kJ: '1000' },
  power: { W: '1', kW: '1000' },
  frequency: { Hz: '1', kHz: '1000' },
  temperature: { K: '1' },
})) {
  for (const [unit, factor] of Object.entries(units))
    definitions[unit] = { dimension: dimensions[dimension], factor };
}
definitions['°C'] = { dimension: dimensions.temperature, factor: '1', offset: '273.15' };
export const UNIT_PRESETS = Object.freeze([
  'mm',
  'cm',
  'm',
  'km',
  'in',
  'ft',
  'm^2',
  'm^3',
  'kg',
  'g',
  'lb',
  's',
  'ms',
  'min',
  'h',
  'm/s',
  'km/h',
  'm/s^2',
  'N',
  'kN',
  'Pa',
  'kPa',
  'MPa',
  'GPa',
  'J',
  'kJ',
  'W',
  'kW',
  'Hz',
  'K',
  '°C',
]);
const same = (a: Dimension, b: Dimension) => a.every((value, index) => value === b[index]);
const dimensionOf = (value: NumericValue): Dimension => value.dimension ?? none;
const attach = (
  value: NumericValue,
  dimension: Dimension,
  quantityKind?: string,
): NumericValue => ({
  ...value,
  ...(dimension.some(Boolean) ? { dimension } : {}),
  ...(quantityKind ? { quantityKind } : {}),
});

export function resolveUnit(unit: string, options: NumericOptions = {}) {
  const math = createArithmetic(options);
  let dimension: Dimension = none,
    factor = math.parse('1');
  const direct = Object.hasOwn(definitions, unit) ? definitions[unit] : undefined;
  if (direct)
    return {
      ...direct,
      factor: math.parse(direct.factor),
      offset: math.parse(direct.offset ?? '0'),
    };
  const normalized = unit
    .replaceAll('²', '^2')
    .replaceAll('³', '^3')
    .replace(/\s*([*/])\s*/g, '$1')
    .replace(/\s+/g, '*');
  const tokens = normalized.split(/([*/])/);
  let sign = 1;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (i % 2) {
      sign = token === '/' ? -1 : 1;
      continue;
    }
    const match = /^([A-Za-z°]+)(?:\^(-?\d+))?$/.exec(token),
      definition =
        match && Object.hasOwn(definitions, match[1]) ? definitions[match[1]] : undefined;
    if (!match || !definition || definition.offset) throw new CalculationError('#VALUE!');
    const exponent = Number(match[2] ?? 1) * sign;
    if (!Number.isInteger(exponent) || Math.abs(exponent) > 16)
      throw new CalculationError('#VALUE!');
    dimension = dimension.map(
      (n, index) => n + definition.dimension[index] * exponent,
    ) as unknown as Dimension;
    factor = math.multiply(
      factor,
      math.power(math.parse(definition.factor), math.parse(String(exponent))),
    );
  }
  return { dimension, factor, offset: math.parse('0') };
}

export function makeQuantity(
  value: NumericValue,
  unit: string,
  options: NumericOptions = {},
  quantityKind?: string,
): NumericValue {
  const math = createArithmetic(options),
    definition = resolveUnit(unit, options);
  return attach(
    math.add(math.multiply(value, definition.factor), definition.offset),
    definition.dimension,
    quantityKind,
  );
}
export function quantityInUnit(
  value: NumericValue,
  unit: string,
  options: NumericOptions = {},
): NumericValue {
  const math = createArithmetic(options),
    definition = resolveUnit(unit, options);
  if (!same(dimensionOf(value), definition.dimension)) throw new CalculationError('#VALUE!');
  const number = { ...value };
  delete number.dimension;
  delete number.quantityKind;
  return math.divide(math.subtract(number, definition.offset), definition.factor);
}
export function validateUnitDimension(unit: string, expected?: string) {
  const definition = resolveUnit(unit);
  if (
    expected &&
    (!Object.hasOwn(dimensions, expected) || !same(definition.dimension, dimensions[expected]))
  )
    throw new CalculationError('#VALUE!');
}
export function dimensionLabel(value: NumericValue): string {
  return (value.dimension ?? [])
    .map((power, index) =>
      power
        ? `${['m', 'kg', 's', 'A', 'K', 'mol', 'cd'][index]}${power === 1 ? '' : `^${power}`}`
        : '',
    )
    .filter(Boolean)
    .join('·');
}

/** The same rational kernel, with dimensional checks and no conversion through binary floats. */
export function createQuantityArithmetic(options: NumericOptions = {}) {
  const math = createArithmetic(options);
  const dimensionless = (value: NumericValue) => {
    if (dimensionOf(value).some(Boolean)) throw new CalculationError('#VALUE!');
  };
  const compatible = (a: NumericValue, b: NumericValue): Dimension => {
    const ad = dimensionOf(a),
      bd = dimensionOf(b);
    if (!same(ad, bd)) {
      if (a.numerator === '0' && !ad.some(Boolean)) return bd;
      if (b.numerator === '0' && !bd.some(Boolean)) return ad;
      throw new CalculationError('#VALUE!');
    }
    if (a.quantityKind && b.quantityKind && a.quantityKind !== b.quantityKind)
      throw new CalculationError('#VALUE!');
    return ad;
  };
  return {
    ...math,
    add: (a: NumericValue, b: NumericValue) =>
      attach(math.add(a, b), compatible(a, b), a.quantityKind ?? b.quantityKind),
    subtract: (a: NumericValue, b: NumericValue) =>
      attach(math.subtract(a, b), compatible(a, b), a.quantityKind ?? b.quantityKind),
    compare: (a: NumericValue, b: NumericValue) => {
      compatible(a, b);
      return math.compare(a, b);
    },
    multiply: (a: NumericValue, b: NumericValue) =>
      attach(
        math.multiply(a, b),
        dimensionOf(a).map((n, i) => n + dimensionOf(b)[i]) as unknown as Dimension,
      ),
    divide: (a: NumericValue, b: NumericValue) =>
      attach(
        math.divide(a, b),
        dimensionOf(a).map((n, i) => n - dimensionOf(b)[i]) as unknown as Dimension,
      ),
    negate: (a: NumericValue) => attach(math.negate(a), dimensionOf(a), a.quantityKind),
    power: (a: NumericValue, b: NumericValue) => {
      dimensionless(b);
      const dimension = dimensionOf(a).map((n) => {
        const result = BigInt(n) * BigInt(b.numerator);
        if (result % BigInt(b.denominator)) throw new CalculationError('#VALUE!');
        const exponent = Number(result / BigInt(b.denominator));
        if (Math.abs(exponent) > 32) throw new CalculationError('#LIMIT!');
        return exponent;
      }) as unknown as Dimension;
      return attach(math.power(a, b), dimension);
    },
    sqrt: (a: NumericValue) => {
      if (dimensionOf(a).some((n) => n % 2)) throw new CalculationError('#VALUE!');
      return attach(math.sqrt(a), dimensionOf(a).map((n) => n / 2) as unknown as Dimension);
    },
    round: (...args: Parameters<typeof math.round>) => {
      dimensionless(args[0]);
      return math.round(...args);
    },
    unary: (...args: Parameters<typeof math.unary>) => {
      dimensionless(args[1]);
      return math.unary(...args);
    },
  };
}
