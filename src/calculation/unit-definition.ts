// Unit catalog derived from platform-sdk at 43f4f64773c9eca230aabc390ea7b03be3c290da.
// See NOTICE.md for source attribution. Numeric conversion strings are intentionally retained.
export interface UnitPrefixDefinition {
  name: string;
  input: readonly string[];
  display: string;
  base: 2 | 10;
  exponent: number;
  factorText: string;
}

export interface UnitConversionDefinition {
  readonly unit: string;
  readonly numerator: string;
  readonly denominator?: string;
  readonly offset?: string;
  readonly pi?: boolean;
  readonly requiresWordSize?: boolean;
}

export interface UnitDefinition {
  name: string;
  input: readonly string[];
  display: string;
  conversionText: string;
  conversion: Readonly<UnitConversionDefinition>;
  description: string;
}

export interface UnitCategoryDefinition {
  name: string;
  canonicalUnit: string;
  units: readonly Readonly<UnitDefinition>[];
}

export const freezePrefix = (prefix: UnitPrefixDefinition): Readonly<UnitPrefixDefinition> =>
  Object.freeze({
    ...prefix,
    input: Object.freeze([...prefix.input]),
  });

export const freezeUnit = (unit: UnitDefinition): Readonly<UnitDefinition> =>
  Object.freeze({
    ...unit,
    conversion: Object.freeze({ ...unit.conversion }),
    input: Object.freeze([...unit.input]),
  });

export const freezeCategory = (
  category: UnitCategoryDefinition,
): Readonly<UnitCategoryDefinition> =>
  Object.freeze({
    ...category,
    units: Object.freeze(category.units.map((unit) => freezeUnit(unit))),
  });
