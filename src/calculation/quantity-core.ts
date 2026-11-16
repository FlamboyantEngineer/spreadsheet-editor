import { UNIT_DEFINITIONS, type UnitConversionDefinition } from './unit-catalog.js';
import { CalculationError, type NumericValue, TEXT_COLLATOR } from './values.js';

export type Dimension = readonly [number, number, number, number, number, number, number];
export const none: Dimension = [0, 0, 0, 0, 0, 0, 0];
export const dimensions: Record<string, Dimension> = {
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

export interface Definition {
  conversion: Readonly<UnitConversionDefinition>;
  quantityKind?: string;
}
function categoryKind(category: string): { quantityKind?: string } {
  if (['energy', 'frequency', 'angle'].includes(category)) return { quantityKind: category };
  return category === 'dataStorage' ? { quantityKind: 'information' } : {};
}
/** Every spelling of a catalog unit; bare C and F stay coulomb and farad, not temperatures. */
const catalogEntries = Object.entries(UNIT_DEFINITIONS).flatMap(([category, data]) =>
  data.units.flatMap((unit) =>
    [...new Set([...unit.input, unit.display])]
      .filter((alias) => category !== 'temperature' || !['C', 'F'].includes(alias))
      .map((alias): [string, Definition] => [
        alias,
        { conversion: unit.conversion, ...categoryKind(category) },
      ]),
  ),
);
export const definitions: Record<string, Definition> = Object.create(null) as Record<
  string,
  Definition
>;
for (const [alias, definition] of catalogEntries) definitions[alias] = definition;
for (const [alias, canonical] of [
  ['degC', '°C'],
  ['degF', '°F'],
  ['ohm', 'Ω'],
  ['bit', 'b'],
  ['bits', 'b'],
])
  definitions[alias] = definitions[canonical];

export const same = (a: Dimension, b: Dimension) => a.every((v, i) => v === b[i]);
export const dimensionOf = (v: NumericValue): Dimension => v.dimension ?? none;
export const scalar = (v: NumericValue) => !dimensionOf(v).some(Boolean) && !v.quantityKind;
export const raw = (v: NumericValue): NumericValue => ({
  kind: 'number',
  numerator: v.numerator,
  denominator: v.denominator,
  approximate: v.approximate,
  ...(v.precision === undefined ? {} : { precision: v.precision }),
});
export function nonSIKindPowers(kind?: string): Record<string, number> {
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
export function kindsFromPowers(powers: Record<string, number>): string | undefined {
  return (
    Object.entries(powers)
      .filter(([, n]) => n)
      .sort(([a], [b]) => TEXT_COLLATOR.compare(a, b))
      .map(([kind, n]) => (n === 1 ? kind : `${kind}^${n}`))
      .join('*') || undefined
  );
}
export function combineKinds(
  a: NumericValue,
  b: NumericValue,
  subtract = false,
): string | undefined {
  const powers = nonSIKindPowers(a.quantityKind);
  for (const [kind, power] of Object.entries(nonSIKindPowers(b.quantityKind)))
    powers[kind] = (powers[kind] ?? 0) + (subtract ? -power : power);
  return kindsFromPowers(powers);
}

/** What a quantity carries besides its magnitude. */
interface QuantityParts {
  dimension: Dimension;
  quantityKind?: string;
  temperature?: NumericValue['temperature'];
  provenance?: Record<string, string>;
}
/** A magnitude with the given dimension, kind, temperature role and provenance. */
export function attach(
  value: NumericValue,
  { dimension, quantityKind, temperature, provenance }: QuantityParts,
): NumericValue {
  if (dimension.some((n) => Math.abs(n) > 32)) throw new CalculationError('#LIMIT!');
  if (provenance && Object.keys(provenance).length > 128) throw new CalculationError('#LIMIT!');
  const role = temperature || (same(dimension, dimensions.temperature) ? 'difference' : undefined);
  return {
    ...raw(value),
    ...(dimension.some(Boolean) ? { dimension } : {}),
    ...(quantityKind ? { quantityKind } : {}),
    ...(role ? { temperature: role } : {}),
    ...(provenance ? { provenance } : {}),
  };
}
/** A copy of an operation's result that keeps the provenance of its source. */
export const keepProvenance = (result: NumericValue, source: NumericValue): NumericValue => ({
  ...result,
  ...(source.provenance ? { provenance: { ...source.provenance } } : {}),
});
export function canonicalDimension(expression: string): Dimension {
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
