import { attach, dimensionOf, dimensions, nonSIKindPowers, raw, same } from './quantity-core.js';
import { createArithmetic, type NumericOptions } from './rational.js';
import { resolveUnit } from './unit-resolution.js';
import {
  CalculationError,
  mergeNumericProvenance,
  normalizeNumericValue,
  type NumericValue,
} from './values.js';

export { createQuantityArithmetic } from './quantity-arithmetic.js';
export type { Dimension } from './quantity-core.js';
export { BINARY_PREFIXES, SI_DECIMAL_PREFIXES, UNIT_DEFINITIONS } from './unit-catalog.js';
export { type ResolvedUnit, resolveUnit } from './unit-resolution.js';

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
/** Options for `makeQuantity`: the arithmetic context and the quantity's physical kind. */
export interface QuantityOptions extends NumericOptions {
  quantityKind?: string;
}
export function makeQuantity(
  value: NumericValue,
  unit: string,
  options: QuantityOptions = {},
): NumericValue {
  const { quantityKind } = options;
  const math = createArithmetic(options),
    definition = resolveUnit(unit, options);
  if (quantityKind && definition.quantityKind && quantityKind !== definition.quantityKind)
    throw new CalculationError('#VALUE!');
  const temperature = same(definition.dimension, dimensions.temperature) ? 'absolute' : undefined;
  return attach(math.add(math.multiply(raw(value), definition.factor), definition.offset), {
    dimension: definition.dimension,
    quantityKind: quantityKind ?? definition.quantityKind,
    temperature,
    provenance: mergeNumericProvenance(value),
  });
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
    ((Object.keys(nonSIKindPowers(value.quantityKind)).length > 0 ||
      Object.keys(nonSIKindPowers(definition.quantityKind)).length > 0) &&
      value.quantityKind !== definition.quantityKind) ||
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
