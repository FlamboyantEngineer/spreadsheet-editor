import { createArithmetic, type NumericOptions } from './rational.js';
import {
  BINARY_PREFIXES,
  SI_DECIMAL_PREFIXES,
  UNIT_DEFINITIONS,
  type UnitConversionDefinition,
} from './unit-catalog.js';
import {
  CalculationError,
  isPureTemperature,
  mergeNumericProvenance,
  normalizeNumericValue,
  type NumericValue,
} from './values.js';
export { BINARY_PREFIXES, SI_DECIMAL_PREFIXES, UNIT_DEFINITIONS } from './unit-catalog.js';

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
  speed: [1, 0, -1, 0, 0, 0, 0],
  acceleration: [1, 0, -2, 0, 0, 0, 0],
  force: [1, 1, -2, 0, 0, 0, 0],
  pressure: [-1, 1, -2, 0, 0, 0, 0],
  energy: [2, 1, -2, 0, 0, 0, 0],
  power: [2, 1, -3, 0, 0, 0, 0],
  frequency: [0, 0, -1, 0, 0, 0, 0],
  electrical: [0, 0, 0, 1, 0, 0, 0],
  light: [0, 0, 0, 0, 0, 0, 1],
  amountofsubstance: [0, 0, 0, 0, 0, 1, 0],
  angle: none,
  datastorage: none,
  information: none,
  dimensionless: none,
  ratio: none,
};
const baseDimensions: Record<string, Dimension> = {
  m: dimensions.length,
  kg: dimensions.mass,
  s: dimensions.time,
  A: dimensions.electrical,
  K: dimensions.temperature,
  mol: dimensions.amountofsubstance,
  cd: dimensions.light,
  rad: none,
  bit: none,
  N: dimensions.force,
  Pa: dimensions.pressure,
  J: dimensions.energy,
  W: dimensions.power,
  Hz: dimensions.frequency,
  C: [0, 0, 1, 1, 0, 0, 0],
  V: [2, 1, -3, -1, 0, 0, 0],
  ohm: [2, 1, -3, -2, 0, 0, 0],
  F: [-2, -1, 4, 2, 0, 0, 0],
  H: [2, 1, -2, -2, 0, 0, 0],
};
interface Definition {
  conversion: Readonly<UnitConversionDefinition>;
  quantityKind?: string;
}
const definitions: Record<string, Definition> = Object.create(null) as Record<string, Definition>;
for (const [category, data] of Object.entries(UNIT_DEFINITIONS)) {
  for (const definition of data.units)
    for (const alias of new Set([...definition.input, definition.display])) {
      if (category === 'temperature' && ['C', 'F'].includes(alias)) continue;
      definitions[alias] = {
        conversion: definition.conversion,
        ...(['energy', 'frequency', 'angle'].includes(category)
          ? { quantityKind: category }
          : category === 'dataStorage'
            ? { quantityKind: 'information' }
            : {}),
      };
    }
}
for (const [alias, canonical] of [
  ['degC', '°C'],
  ['degF', '°F'],
  ['ohm', 'Ω'],
  ['bit', 'b'],
  ['bits', 'b'],
])
  definitions[alias] = definitions[canonical];
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
  '°F',
]);
const same = (a: Dimension, b: Dimension) => a.every((v, i) => v === b[i]);
const dimensionOf = (v: NumericValue): Dimension => v.dimension ?? none;
const scalar = (v: NumericValue) => !dimensionOf(v).some(Boolean) && !v.quantityKind;
const raw = (v: NumericValue): NumericValue => ({
  kind: 'number',
  numerator: v.numerator,
  denominator: v.denominator,
  approximate: v.approximate,
  ...(v.precision === undefined ? {} : { precision: v.precision }),
});
function nonSIKindPowers(kind?: string): Record<string, number> {
  const powers: Record<string, number> = {};
  if (
    !kind ||
    !/^(?:angle|information)(?:\^-?\d+)?(?:\*(?:angle|information)(?:\^-?\d+)?)*$/.test(kind)
  )
    return powers;
  for (const token of kind.split('*')) {
    const [name, power] = token.split('^');
    powers[name] = (powers[name] ?? 0) + Number(power ?? 1);
  }
  return powers;
}
function kindsFromPowers(powers: Record<string, number>): string | undefined {
  return (
    Object.entries(powers)
      .filter(([, n]) => n)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([kind, n]) => (n === 1 ? kind : `${kind}^${n}`))
      .join('*') || undefined
  );
}
function combineKinds(a: NumericValue, b: NumericValue, subtract = false): string | undefined {
  const powers = nonSIKindPowers(a.quantityKind);
  for (const [kind, power] of Object.entries(nonSIKindPowers(b.quantityKind)))
    powers[kind] = (powers[kind] ?? 0) + (subtract ? -power : power);
  return kindsFromPowers(powers);
}
function attach(
  value: NumericValue,
  dimension: Dimension,
  quantityKind?: string,
  temperature?: NumericValue['temperature'],
  provenance?: Record<string, string>,
): NumericValue {
  if (dimension.some((n) => Math.abs(n) > 32)) throw new CalculationError('#LIMIT!');
  if (provenance && Object.keys(provenance).length > 128) throw new CalculationError('#LIMIT!');
  return {
    ...raw(value),
    ...(dimension.some(Boolean) ? { dimension } : {}),
    ...(quantityKind ? { quantityKind } : {}),
    ...(temperature
      ? { temperature }
      : same(dimension, dimensions.temperature)
        ? { temperature: 'difference' as const }
        : {}),
    ...(provenance ? { provenance } : {}),
  };
}
function canonicalDimension(expression: string): Dimension {
  let result: Dimension = none,
    sign = 1;
  for (const token of expression
    .replaceAll('²', '^2')
    .replaceAll('³', '^3')
    .split(/([*/])|\s+/)
    .filter(Boolean)) {
    if (token === '*' || token === '/') {
      sign = token === '/' ? -1 : 1;
      continue;
    }
    const match = /^([\p{L}]+)(?:\^(-?\d+))?$/u.exec(token);
    const base = match && baseDimensions[match[1]];
    if (!match || !base) throw new CalculationError('#VALUE!');
    const exponent = Number(match[2] ?? 1) * sign;
    result = result.map((n, i) => n + base[i] * exponent) as unknown as Dimension;
  }
  return result;
}
function atom(token: string, options: NumericOptions) {
  const math = createArithmetic(options);
  let definition = definitions[token],
    prefixFactor = math.parse('1');
  if (!definition)
    for (const prefix of [...SI_DECIMAL_PREFIXES, ...BINARY_PREFIXES]) {
      for (const alias of prefix.input) {
        if (!token.startsWith(alias)) continue;
        const base = token.slice(alias.length),
          candidate = definitions[base];
        if (
          !candidate ||
          candidate.conversion.offset ||
          candidate.conversion.requiresWordSize ||
          (prefix.base === 2 && !['b', 'B', 'o'].includes(base))
        )
          continue;
        definition = candidate;
        prefixFactor = math.power(
          math.parse(String(prefix.base)),
          math.parse(String(prefix.exponent)),
        );
        break;
      }
      if (definition) break;
    }
  if (!definition) {
    const prefix = SI_DECIMAL_PREFIXES.find((prefix) => prefix.input.includes(token));
    if (prefix)
      return {
        dimension: none,
        factor: math.power(math.parse('10'), math.parse(String(prefix.exponent))),
        offset: math.parse('0'),
        quantityKind: undefined,
      };
    throw new CalculationError('#VALUE!');
  }
  const conversion = definition.conversion;
  let factor: NumericValue;
  if (conversion.requiresWordSize) {
    if (
      typeof options.wordSizeBits !== 'string' ||
      !options.wordSizeBits ||
      !/^[1-9]\d{0,5}$/.test(options.wordSizeBits) ||
      BigInt(options.wordSizeBits) > 100000n
    )
      throw new CalculationError('#VALUE!');
    factor = math.parse(options.wordSizeBits);
  } else
    factor = math.divide(
      math.parse(conversion.numerator),
      math.parse(conversion.denominator ?? '1'),
    );
  if (conversion.pi) factor = math.multiply(factor, math.pi());
  factor = math.multiply(factor, prefixFactor);
  return {
    dimension: canonicalDimension(conversion.unit),
    factor,
    offset: math.multiply(math.parse(conversion.offset ?? '0'), factor),
    quantityKind:
      definition.quantityKind ?? (conversion.unit.includes('rad^2') ? 'angle^2' : undefined),
  };
}
export function resolveUnit(unit: string, options: NumericOptions = {}) {
  if (typeof unit !== 'string' || !unit || unit.length > 256) throw new CalculationError('#VALUE!');
  const math = createArithmetic(options),
    direct = definitions[unit];
  if (direct) return atom(unit, options);
  const normalized = unit
    .replaceAll('²', '^2')
    .replaceAll('³', '^3')
    .replace(/\s*([*/])\s*/g, '$1')
    .replace(/fl oz/g, 'fl_oz')
    .replace(/\s+/g, '*');
  const tokens = normalized.split(/([*/])/);
  let dimension: Dimension = none,
    factor = math.parse('1'),
    sign = 1;
  const kindPowers: Record<string, number> = {};
  for (let i = 0; i < tokens.length; i++) {
    if (i % 2) {
      sign = tokens[i] === '/' ? -1 : 1;
      continue;
    }
    const match = /^([\p{L}µΩ°′″'"_]+)(?:\^([+-]?\d+))?$/u.exec(tokens[i]);
    if (!match) throw new CalculationError('#VALUE!');
    const definition = atom(match[1].replace('fl_oz', 'fl oz'), options),
      exponent = Number(match[2] ?? 1) * sign;
    if (!Number.isInteger(exponent) || Math.abs(exponent) > 16)
      throw new CalculationError('#VALUE!');
    dimension = dimension.map(
      (n, index) => n + definition.dimension[index] * exponent,
    ) as unknown as Dimension;
    if (dimension.some((n) => Math.abs(n) > 32)) throw new CalculationError('#LIMIT!');
    factor = math.multiply(factor, math.power(definition.factor, math.parse(String(exponent))));
    if (definition.quantityKind === 'angle' || definition.quantityKind === 'information')
      kindPowers[definition.quantityKind] = (kindPowers[definition.quantityKind] ?? 0) + exponent;
  }
  const quantityKind =
    Object.entries(kindPowers)
      .filter(([, n]) => n)
      .map(([kind, n]) => (n === 1 ? kind : `${kind}^${n}`))
      .join('*') || undefined;
  return { dimension, factor, offset: math.parse('0'), quantityKind };
}
export function makeQuantity(
  value: NumericValue,
  unit: string,
  options: NumericOptions = {},
  quantityKind?: string,
): NumericValue {
  const math = createArithmetic(options),
    definition = resolveUnit(unit, options);
  if (quantityKind && definition.quantityKind && quantityKind !== definition.quantityKind)
    throw new CalculationError('#VALUE!');
  const temperature = same(definition.dimension, dimensions.temperature) ? 'absolute' : undefined;
  return attach(
    math.add(math.multiply(raw(value), definition.factor), definition.offset),
    definition.dimension,
    quantityKind ?? definition.quantityKind,
    temperature,
    mergeNumericProvenance(value),
  );
}
export function quantityInUnit(
  value: NumericValue,
  unit: string,
  options: NumericOptions = {},
): NumericValue {
  const math = createArithmetic(options),
    definition = resolveUnit(unit, options),
    normalized = normalizeNumericValue(value);
  if (
    !same(dimensionOf(value), definition.dimension) ||
    (value.quantityKind &&
      definition.quantityKind &&
      value.quantityKind !== definition.quantityKind)
  )
    throw new CalculationError('#VALUE!');
  const offset = normalized.temperature === 'difference' ? math.parse('0') : definition.offset;
  return {
    ...math.divide(math.subtract(raw(value), offset), definition.factor),
    ...(value.provenance ? { provenance: { ...value.provenance } } : {}),
  };
}
export function validateUnitDimension(
  unit: string,
  expected?: string,
  options: NumericOptions = {},
) {
  const definition = resolveUnit(unit, options);
  if (!expected) return;
  const key = expected.toLowerCase().replace(/\s+/g, '');
  const dimension = dimensions[key] ?? resolveUnit(expected, options).dimension;
  if (
    !same(definition.dimension, dimension) ||
    (['dimensionless', 'ratio'].includes(key) && definition.quantityKind) ||
    (key === 'angle' && definition.quantityKind !== 'angle') ||
    (['datastorage', 'information'].includes(key) && definition.quantityKind !== 'information') ||
    (!dimensions[key] &&
      resolveUnit(expected, options).quantityKind !== definition.quantityKind &&
      definition.dimension.every((n) => n === 0))
  )
    throw new CalculationError('#VALUE!');
}
export function dimensionLabel(value: NumericValue): string {
  const label = (value.dimension ?? [])
    .map((power, index) =>
      power
        ? `${['m', 'kg', 's', 'A', 'K', 'mol', 'cd'][index]}${power === 1 ? '' : `^${power}`}`
        : '',
    )
    .filter(Boolean)
    .join('·');
  const nonSI = Object.entries(nonSIKindPowers(value.quantityKind))
    .filter(([, power]) => power)
    .map(([kind, power]) => `${kind === 'angle' ? 'rad' : 'bit'}${power === 1 ? '' : `^${power}`}`)
    .join('·');
  return [label, nonSI].filter(Boolean).join('·');
}

/** Dimensional and physical-kind checks wrap the exact rational kernel. */
export function createQuantityArithmetic(options: NumericOptions = {}) {
  const math = createArithmetic(options);
  const dimensionless = (v: NumericValue) => {
    if (!scalar(v)) throw new CalculationError('#VALUE!');
  };
  const compatible = (a: NumericValue, b: NumericValue): Dimension => {
    const ad = dimensionOf(a),
      bd = dimensionOf(b);
    if (!same(ad, bd)) {
      if (a.numerator === '0' && scalar(a)) return bd;
      if (b.numerator === '0' && scalar(b)) return ad;
      throw new CalculationError('#VALUE!');
    }
    if (
      (a.quantityKind && b.quantityKind && a.quantityKind !== b.quantityKind) ||
      (Object.keys(nonSIKindPowers(a.quantityKind)).length > 0 !==
        Object.keys(nonSIKindPowers(b.quantityKind)).length > 0 &&
        !(a.numerator === '0' && scalar(a)) &&
        !(b.numerator === '0' && scalar(b)))
    )
      throw new CalculationError('#VALUE!');
    mergeNumericProvenance(a, b);
    return ad;
  };
  const temp = (v: NumericValue) =>
    isPureTemperature(v) ? (v.temperature ?? 'absolute') : undefined;
  const addition = (a: NumericValue, b: NumericValue, subtract: boolean) => {
    const dimension = compatible(a, b),
      at = temp(a),
      bt = temp(b);
    if (
      (!subtract && at === 'absolute' && bt === 'absolute') ||
      (subtract && at === 'difference' && bt === 'absolute')
    )
      throw new CalculationError('#VALUE!');
    const temperature =
      at || bt
        ? subtract && at === 'absolute' && bt === 'absolute'
          ? 'difference'
          : at === 'absolute' || bt === 'absolute'
            ? 'absolute'
            : 'difference'
        : undefined;
    return attach(
      subtract ? math.subtract(raw(a), raw(b)) : math.add(raw(a), raw(b)),
      dimension,
      a.quantityKind ?? b.quantityKind,
      temperature,
      mergeNumericProvenance(a, b),
    );
  };
  const average = (values: readonly NumericValue[]): NumericValue => {
    const first = values[0];
    if (!first) throw new CalculationError('#DIV/0!');
    values.forEach((v) => {
      compatible(first, v);
      if (temp(first) !== temp(v)) throw new CalculationError('#VALUE!');
    });
    return attach(
      math.divide(
        values.reduce((sum, v) => math.add(sum, raw(v)), math.parse('0')),
        math.parse(String(values.length)),
      ),
      dimensionOf(first),
      first.quantityKind,
      temp(first),
      mergeNumericProvenance(...values),
    );
  };
  const unary = (...args: Parameters<typeof math.unary>) => {
    if (
      ['SIN', 'COS', 'TAN'].includes(args[0]) &&
      args[1].quantityKind === 'angle' &&
      !dimensionOf(args[1]).some(Boolean)
    )
      return {
        ...math.unary(args[0], raw(args[1])),
        ...(args[1].provenance ? { provenance: { ...args[1].provenance } } : {}),
      };
    dimensionless(args[1]);
    return {
      ...math.unary(...args),
      ...(args[1].provenance ? { provenance: { ...args[1].provenance } } : {}),
    };
  };
  return {
    ...math,
    average,
    add: (a: NumericValue, b: NumericValue) => addition(a, b, false),
    subtract: (a: NumericValue, b: NumericValue) => addition(a, b, true),
    compare: (a: NumericValue, b: NumericValue) => {
      compatible(a, b);
      if (temp(a) && temp(b) && temp(a) !== temp(b)) throw new CalculationError('#VALUE!');
      return math.compare(a, b);
    },
    multiply: (a: NumericValue, b: NumericValue) =>
      attach(
        math.multiply(raw(a), raw(b)),
        dimensionOf(a).map((n, i) => n + dimensionOf(b)[i]) as unknown as Dimension,
        scalar(b) ? a.quantityKind : scalar(a) ? b.quantityKind : combineKinds(a, b),
        scalar(b) ? temp(a) : scalar(a) ? temp(b) : undefined,
        mergeNumericProvenance(a, b),
      ),
    divide: (a: NumericValue, b: NumericValue) =>
      attach(
        math.divide(raw(a), raw(b)),
        dimensionOf(a).map((n, i) => n - dimensionOf(b)[i]) as unknown as Dimension,
        scalar(b) ? a.quantityKind : combineKinds(a, b, true),
        scalar(b) ? temp(a) : undefined,
        mergeNumericProvenance(a, b),
      ),
    negate: (a: NumericValue) => {
      if (temp(a) === 'absolute') throw new CalculationError('#VALUE!');
      return attach(
        math.negate(raw(a)),
        dimensionOf(a),
        a.quantityKind,
        temp(a),
        mergeNumericProvenance(a),
      );
    },
    power: (a: NumericValue, b: NumericValue) => {
      dimensionless(b);
      const dimension = dimensionOf(a).map((n) => {
        const result = BigInt(n) * BigInt(b.numerator);
        if (result % BigInt(b.denominator)) throw new CalculationError('#VALUE!');
        const exponent = Number(result / BigInt(b.denominator));
        if (Math.abs(exponent) > 32) throw new CalculationError('#LIMIT!');
        return exponent;
      }) as unknown as Dimension;
      return attach(
        math.power(raw(a), raw(b)),
        dimension,
        b.numerator === b.denominator
          ? a.quantityKind
          : kindsFromPowers(
              Object.fromEntries(
                Object.entries(nonSIKindPowers(a.quantityKind)).map(([kind, power]) => {
                  const exponent = BigInt(power) * BigInt(b.numerator);
                  if (exponent % BigInt(b.denominator)) throw new CalculationError('#VALUE!');
                  return [kind, Number(exponent / BigInt(b.denominator))];
                }),
              ),
            ),
        b.numerator === b.denominator ? temp(a) : undefined,
        mergeNumericProvenance(a, b),
      );
    },
    sqrt: (a: NumericValue) => {
      if (dimensionOf(a).some((n) => n % 2)) throw new CalculationError('#VALUE!');
      return attach(
        math.sqrt(raw(a)),
        dimensionOf(a).map((n) => n / 2) as unknown as Dimension,
        kindsFromPowers(
          Object.fromEntries(
            Object.entries(nonSIKindPowers(a.quantityKind)).map(([kind, power]) => {
              if (power % 2) throw new CalculationError('#VALUE!');
              return [kind, power / 2];
            }),
          ),
        ),
        undefined,
        mergeNumericProvenance(a),
      );
    },
    round: (...args: Parameters<typeof math.round>) => {
      dimensionless(args[0]);
      return {
        ...math.round(...args),
        ...(args[0].provenance ? { provenance: { ...args[0].provenance } } : {}),
      };
    },
    unary,
  };
}
