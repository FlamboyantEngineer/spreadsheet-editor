import {
  canonicalDimension,
  type Definition,
  definitions,
  type Dimension,
  none,
} from './quantity-core.js';
import { type Arithmetic, createArithmetic, type NumericOptions } from './rational.js';
import {
  BINARY_PREFIXES,
  SI_DECIMAL_PREFIXES,
  type UnitConversionDefinition,
} from './unit-catalog.js';
import { CalculationError, type NumericValue } from './values.js';

/** A unit resolved to its canonical dimension, scale, offset and non-SI kind. */
export interface ResolvedUnit {
  dimension: Dimension;
  factor: NumericValue;
  offset: NumericValue;
  quantityKind: string | undefined;
}
type Prefix = (typeof SI_DECIMAL_PREFIXES)[number];
const PREFIXES = [...SI_DECIMAL_PREFIXES, ...BINARY_PREFIXES];

/** The unit a token names after a prefix spelling; binary prefixes apply only to b, B and o. */
function prefixable(token: string, prefix: Prefix, alias: string): Definition | undefined {
  if (!token.startsWith(alias)) return undefined;
  const base = token.slice(alias.length),
    candidate = definitions[base];
  if (!candidate || candidate.conversion.offset || candidate.conversion.requiresWordSize)
    return undefined;
  return prefix.base === 2 && !['b', 'B', 'o'].includes(base) ? undefined : candidate;
}

/** The first prefix and unit a token splits into, with the prefix's factor. */
function prefixed(token: string, math: Arithmetic) {
  for (const prefix of PREFIXES) {
    const definition = prefix.input.map((alias) => prefixable(token, prefix, alias)).find(Boolean);
    if (definition)
      return {
        definition,
        prefixFactor: math.power(
          math.parse(String(prefix.base)),
          math.parse(String(prefix.exponent)),
        ),
      };
  }
  return undefined;
}

/** A bare decimal prefix, such as `k`, as a dimensionless scale. */
function barePrefix(token: string, math: Arithmetic): ResolvedUnit {
  const prefix = SI_DECIMAL_PREFIXES.find((candidate) => candidate.input.includes(token));
  if (!prefix) throw new CalculationError('#VALUE!');
  return {
    dimension: none,
    factor: math.power(math.parse('10'), math.parse(String(prefix.exponent))),
    offset: math.parse('0'),
    quantityKind: undefined,
  };
}

function wordSize(options: NumericOptions): string {
  const bits = options.wordSizeBits;
  if (typeof bits !== 'string' || !/^[1-9]\d{0,5}$/.test(bits) || BigInt(bits) > 100000n)
    throw new CalculationError('#VALUE!');
  return bits;
}

function conversionFactor(
  conversion: Readonly<UnitConversionDefinition>,
  options: NumericOptions,
  math: Arithmetic,
): NumericValue {
  const factor = conversion.requiresWordSize
    ? math.parse(wordSize(options))
    : math.divide(math.parse(conversion.numerator), math.parse(conversion.denominator ?? '1'));
  return conversion.pi ? math.multiply(factor, math.pi()) : factor;
}

/** One unit symbol, possibly with a prefix. */
function atom(token: string, options: NumericOptions): ResolvedUnit {
  const math = createArithmetic(options);
  const direct = definitions[token];
  const found = direct
    ? { definition: direct, prefixFactor: math.parse('1') }
    : prefixed(token, math);
  if (!found) return barePrefix(token, math);
  const { conversion, quantityKind } = found.definition;
  const factor = math.multiply(conversionFactor(conversion, options, math), found.prefixFactor);
  return {
    dimension: canonicalDimension(conversion.unit),
    factor,
    offset: math.multiply(math.parse(conversion.offset ?? '0'), factor),
    quantityKind: quantityKind ?? (conversion.unit.includes('rad^2') ? 'angle^2' : undefined),
  };
}

const UNIT_TERM = /^([\p{L}µΩ°′″'"_]+)(?:\^([+-]?\d+))?$/u;

/** One `symbol^power` term of a compound unit, with the sign of its separator applied. */
function unitTerm(text: string, sign: number, options: NumericOptions) {
  const match = UNIT_TERM.exec(text);
  if (!match) throw new CalculationError('#VALUE!');
  const unit = atom(match[1].replace('fl_oz', 'fl oz'), options),
    exponent = Number(match[2] ?? 1) * sign;
  if (!Number.isInteger(exponent) || Math.abs(exponent) > 16) throw new CalculationError('#VALUE!');
  return { unit, exponent };
}

const compoundTerms = (unit: string) =>
  unit
    .replaceAll('²', '^2')
    .replaceAll('³', '^3')
    .replace(/\s*([*/])\s*/g, '$1')
    .replace(/fl oz/g, 'fl_oz')
    .replace(/\s+/g, '*')
    .split(/([*/])/);

export function resolveUnit(unit: string, options: NumericOptions = {}): ResolvedUnit {
  if (typeof unit !== 'string' || !unit || unit.length > 256) throw new CalculationError('#VALUE!');
  if (definitions[unit]) return atom(unit, options);
  const math = createArithmetic(options),
    tokens = compoundTerms(unit);
  let dimension: Dimension = none,
    factor = math.parse('1');
  const kindPowers: Record<string, number> = {};
  for (let i = 0; i < tokens.length; i += 2) {
    const term = unitTerm(tokens[i], tokens[i - 1] === '/' ? -1 : 1, options),
      { quantityKind } = term.unit;
    dimension = dimension.map(
      (n, index) => n + term.unit.dimension[index] * term.exponent,
    ) as unknown as Dimension;
    if (dimension.some((n) => Math.abs(n) > 32)) throw new CalculationError('#LIMIT!');
    factor = math.multiply(factor, math.power(term.unit.factor, math.parse(String(term.exponent))));
    if (quantityKind === 'angle' || quantityKind === 'information')
      kindPowers[quantityKind] = (kindPowers[quantityKind] ?? 0) + term.exponent;
  }
  const quantityKind =
    Object.entries(kindPowers)
      .filter(([, n]) => n)
      .map(([kind, n]) => (n === 1 ? kind : `${kind}^${n}`))
      .join('*') || undefined;
  return { dimension, factor, offset: math.parse('0'), quantityKind };
}
