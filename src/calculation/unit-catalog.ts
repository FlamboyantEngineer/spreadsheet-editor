// Unit catalog derived from platform-sdk at 43f4f64773c9eca230aabc390ea7b03be3c290da.
// See NOTICE.md for source attribution. Numeric conversion strings are intentionally retained.
import { BASE_CATEGORIES } from './unit-categories-base.js';
import { DERIVED_CATEGORIES } from './unit-categories-derived.js';
import { OTHER_CATEGORIES } from './unit-categories-other.js';
import { freezePrefix, type UnitPrefixDefinition } from './unit-definition.js';

export type { UnitConversionDefinition } from './unit-definition.js';

/** A prefix whose display symbol is the last of its input spellings. */
const prefix =
  (base: 2 | 10) => (name: string, symbol: string | readonly string[], exponent: number) => {
    const input = typeof symbol === 'string' ? [symbol] : symbol;
    return freezePrefix({
      name,
      input,
      display: input[input.length - 1],
      base,
      exponent,
      factorText: `${base}^${exponent}`,
    });
  };
const decimal = prefix(10),
  binary = prefix(2);

export const SI_DECIMAL_PREFIXES: readonly Readonly<UnitPrefixDefinition>[] = Object.freeze([
  decimal('quecto', 'q', -30),
  decimal('ronto', 'r', -27),
  decimal('yocto', 'y', -24),
  decimal('zepto', 'z', -21),
  decimal('atto', 'a', -18),
  decimal('femto', 'f', -15),
  decimal('pico', 'p', -12),
  decimal('nano', 'n', -9),
  decimal('micro', ['u', '\u00b5'], -6),
  decimal('milli', 'm', -3),
  decimal('centi', 'c', -2),
  decimal('deci', 'd', -1),
  decimal('deka', 'da', 1),
  decimal('hecto', 'h', 2),
  decimal('kilo', 'k', 3),
  decimal('mega', 'M', 6),
  decimal('giga', 'G', 9),
  decimal('tera', 'T', 12),
  decimal('peta', 'P', 15),
  decimal('exa', 'E', 18),
  decimal('zetta', 'Z', 21),
  decimal('yotta', 'Y', 24),
  decimal('ronna', 'R', 27),
  decimal('quetta', 'Q', 30),
]);

export const BINARY_PREFIXES: readonly Readonly<UnitPrefixDefinition>[] = Object.freeze([
  binary('kibi', 'Ki', 10),
  binary('mebi', 'Mi', 20),
  binary('gibi', 'Gi', 30),
  binary('tebi', 'Ti', 40),
  binary('pebi', 'Pi', 50),
  binary('exbi', 'Ei', 60),
  binary('zebi', 'Zi', 70),
  binary('yobi', 'Yi', 80),
  binary('robi', 'Ri', 90),
  binary('quebi', 'Qi', 100),
]);

export const UNIT_DEFINITIONS = Object.freeze({
  ...BASE_CATEGORIES,
  ...DERIVED_CATEGORIES,
  ...OTHER_CATEGORIES,
});
