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

const freezePrefix = (prefix: UnitPrefixDefinition): Readonly<UnitPrefixDefinition> =>
  Object.freeze({
    ...prefix,
    input: Object.freeze([...prefix.input]),
  });

const freezeUnit = (unit: UnitDefinition): Readonly<UnitDefinition> =>
  Object.freeze({
    ...unit,
    conversion: Object.freeze({ ...unit.conversion }),
    input: Object.freeze([...unit.input]),
  });

const freezeCategory = (category: UnitCategoryDefinition): Readonly<UnitCategoryDefinition> =>
  Object.freeze({
    ...category,
    units: Object.freeze(category.units.map((unit) => freezeUnit(unit))),
  });

export const SI_DECIMAL_PREFIXES: readonly Readonly<UnitPrefixDefinition>[] = Object.freeze([
  freezePrefix({
    name: 'quecto',
    input: ['q'],
    display: 'q',
    base: 10,
    exponent: -30,
    factorText: '10^-30',
  }),
  freezePrefix({
    name: 'ronto',
    input: ['r'],
    display: 'r',
    base: 10,
    exponent: -27,
    factorText: '10^-27',
  }),
  freezePrefix({
    name: 'yocto',
    input: ['y'],
    display: 'y',
    base: 10,
    exponent: -24,
    factorText: '10^-24',
  }),
  freezePrefix({
    name: 'zepto',
    input: ['z'],
    display: 'z',
    base: 10,
    exponent: -21,
    factorText: '10^-21',
  }),
  freezePrefix({
    name: 'atto',
    input: ['a'],
    display: 'a',
    base: 10,
    exponent: -18,
    factorText: '10^-18',
  }),
  freezePrefix({
    name: 'femto',
    input: ['f'],
    display: 'f',
    base: 10,
    exponent: -15,
    factorText: '10^-15',
  }),
  freezePrefix({
    name: 'pico',
    input: ['p'],
    display: 'p',
    base: 10,
    exponent: -12,
    factorText: '10^-12',
  }),
  freezePrefix({
    name: 'nano',
    input: ['n'],
    display: 'n',
    base: 10,
    exponent: -9,
    factorText: '10^-9',
  }),
  freezePrefix({
    name: 'micro',
    input: ['u', '\u00b5'],
    display: '\u00b5',
    base: 10,
    exponent: -6,
    factorText: '10^-6',
  }),
  freezePrefix({
    name: 'milli',
    input: ['m'],
    display: 'm',
    base: 10,
    exponent: -3,
    factorText: '10^-3',
  }),
  freezePrefix({
    name: 'centi',
    input: ['c'],
    display: 'c',
    base: 10,
    exponent: -2,
    factorText: '10^-2',
  }),
  freezePrefix({
    name: 'deci',
    input: ['d'],
    display: 'd',
    base: 10,
    exponent: -1,
    factorText: '10^-1',
  }),
  freezePrefix({
    name: 'deka',
    input: ['da'],
    display: 'da',
    base: 10,
    exponent: 1,
    factorText: '10^1',
  }),
  freezePrefix({
    name: 'hecto',
    input: ['h'],
    display: 'h',
    base: 10,
    exponent: 2,
    factorText: '10^2',
  }),
  freezePrefix({
    name: 'kilo',
    input: ['k'],
    display: 'k',
    base: 10,
    exponent: 3,
    factorText: '10^3',
  }),
  freezePrefix({
    name: 'mega',
    input: ['M'],
    display: 'M',
    base: 10,
    exponent: 6,
    factorText: '10^6',
  }),
  freezePrefix({
    name: 'giga',
    input: ['G'],
    display: 'G',
    base: 10,
    exponent: 9,
    factorText: '10^9',
  }),
  freezePrefix({
    name: 'tera',
    input: ['T'],
    display: 'T',
    base: 10,
    exponent: 12,
    factorText: '10^12',
  }),
  freezePrefix({
    name: 'peta',
    input: ['P'],
    display: 'P',
    base: 10,
    exponent: 15,
    factorText: '10^15',
  }),
  freezePrefix({
    name: 'exa',
    input: ['E'],
    display: 'E',
    base: 10,
    exponent: 18,
    factorText: '10^18',
  }),
  freezePrefix({
    name: 'zetta',
    input: ['Z'],
    display: 'Z',
    base: 10,
    exponent: 21,
    factorText: '10^21',
  }),
  freezePrefix({
    name: 'yotta',
    input: ['Y'],
    display: 'Y',
    base: 10,
    exponent: 24,
    factorText: '10^24',
  }),
  freezePrefix({
    name: 'ronna',
    input: ['R'],
    display: 'R',
    base: 10,
    exponent: 27,
    factorText: '10^27',
  }),
  freezePrefix({
    name: 'quetta',
    input: ['Q'],
    display: 'Q',
    base: 10,
    exponent: 30,
    factorText: '10^30',
  }),
]);

export const BINARY_PREFIXES: readonly Readonly<UnitPrefixDefinition>[] = Object.freeze([
  freezePrefix({
    name: 'kibi',
    input: ['Ki'],
    display: 'Ki',
    base: 2,
    exponent: 10,
    factorText: '2^10',
  }),
  freezePrefix({
    name: 'mebi',
    input: ['Mi'],
    display: 'Mi',
    base: 2,
    exponent: 20,
    factorText: '2^20',
  }),
  freezePrefix({
    name: 'gibi',
    input: ['Gi'],
    display: 'Gi',
    base: 2,
    exponent: 30,
    factorText: '2^30',
  }),
  freezePrefix({
    name: 'tebi',
    input: ['Ti'],
    display: 'Ti',
    base: 2,
    exponent: 40,
    factorText: '2^40',
  }),
  freezePrefix({
    name: 'pebi',
    input: ['Pi'],
    display: 'Pi',
    base: 2,
    exponent: 50,
    factorText: '2^50',
  }),
  freezePrefix({
    name: 'exbi',
    input: ['Ei'],
    display: 'Ei',
    base: 2,
    exponent: 60,
    factorText: '2^60',
  }),
  freezePrefix({
    name: 'zebi',
    input: ['Zi'],
    display: 'Zi',
    base: 2,
    exponent: 70,
    factorText: '2^70',
  }),
  freezePrefix({
    name: 'yobi',
    input: ['Yi'],
    display: 'Yi',
    base: 2,
    exponent: 80,
    factorText: '2^80',
  }),
  freezePrefix({
    name: 'robi',
    input: ['Ri'],
    display: 'Ri',
    base: 2,
    exponent: 90,
    factorText: '2^90',
  }),
  freezePrefix({
    name: 'quebi',
    input: ['Qi'],
    display: 'Qi',
    base: 2,
    exponent: 100,
    factorText: '2^100',
  }),
]);

export const UNIT_DEFINITIONS = Object.freeze({
  time: freezeCategory({
    name: 'Time',
    canonicalUnit: 's',
    units: [
      {
        name: 'second',
        input: ['s'],
        display: 's',
        conversion: { unit: 's', numerator: '1' },
        conversionText: '1 s',
        description: 'SI base unit of time',
      },
      {
        name: 'minute',
        input: ['min'],
        display: 'min',
        conversion: { unit: 's', numerator: '60' },
        conversionText: '60 s',
        description: '1 minute is equal to 60 seconds',
      },
      {
        name: 'hour',
        input: ['h'],
        display: 'h',
        conversion: { unit: 's', numerator: '3600' },
        conversionText: '3600 s',
        description: '1 hour is equal to 3600 seconds',
      },
    ],
  }),
  length: freezeCategory({
    name: 'Length',
    canonicalUnit: 'm',
    units: [
      {
        name: 'meter',
        input: ['m'],
        display: 'm',
        conversion: { unit: 'm', numerator: '1' },
        conversionText: '1 m',
        description: 'SI base unit of length',
      },
      {
        name: 'angstrom',
        input: ['Å', 'Å', 'angstrom'],
        display: 'Å',
        conversion: { unit: 'm', numerator: '1', denominator: '10000000000' },
        conversionText: '1e-10 m',
        description: '1 angstrom is exactly one ten-billionth of a meter',
      },
      {
        name: 'foot',
        input: ['ft'],
        display: 'ft',
        conversion: { unit: 'm', numerator: '0.3048' },
        conversionText: '0.3048 m',
        description: '1 international foot is exactly 0.3048 meters',
      },
      {
        name: 'inch',
        input: ['in'],
        display: 'in',
        conversion: { unit: 'm', numerator: '0.0254' },
        conversionText: '0.0254 m',
        description: '1 inch is exactly 0.0254 meters',
      },
      {
        name: 'yard',
        input: ['yd'],
        display: 'yd',
        conversion: { unit: 'm', numerator: '0.9144' },
        conversionText: '0.9144 m',
        description: '1 yard is exactly 0.9144 meters',
      },
      {
        name: 'mile',
        input: ['mi'],
        display: 'mi',
        conversion: { unit: 'm', numerator: '1609.344' },
        conversionText: '1609.344 m',
        description: '1 international mile is exactly 1609.344 meters',
      },
      {
        name: 'statute mile',
        input: ['SM'],
        display: 'SM',
        conversion: { unit: 'm', numerator: '1609.344' },
        conversionText: '1609.344 m',
        description: '1 statute mile is equal to 1 international mile',
      },
      {
        name: 'nautical mile',
        input: ['NM', 'nmi'],
        display: 'NM',
        conversion: { unit: 'm', numerator: '1852' },
        conversionText: '1852 m',
        description: '1 nautical mile is exactly 1852 meters',
      },
    ],
  }),
  mass: freezeCategory({
    name: 'Mass',
    canonicalUnit: 'kg',
    units: [
      {
        name: 'kilogram',
        input: ['kg'],
        display: 'kg',
        conversion: { unit: 'kg', numerator: '1' },
        conversionText: '1 kg',
        description: 'SI base unit of mass',
      },
      {
        name: 'gram',
        input: ['g'],
        display: 'g',
        conversion: { unit: 'kg', numerator: '0.001' },
        conversionText: '0.001 kg',
        description: '1 gram is exactly 0.001 kilograms',
      },
      {
        name: 'ounce (avoirdupois)',
        input: ['oz'],
        display: 'oz',
        conversion: { unit: 'kg', numerator: '0.028349523125' },
        conversionText: '0.028349523125 kg',
        description: '1 avoirdupois ounce is exactly 28.349523125 grams',
      },
      {
        name: 'pound (avoirdupois)',
        input: ['lb'],
        display: 'lb',
        conversion: { unit: 'kg', numerator: '0.45359237' },
        conversionText: '0.45359237 kg',
        description: '1 avoirdupois pound is exactly 0.45359237 kilograms',
      },
    ],
  }),
  temperature: freezeCategory({
    name: 'Temperature',
    canonicalUnit: 'K',
    units: [
      {
        name: 'kelvin',
        input: ['K'],
        display: 'K',
        conversion: { unit: 'K', numerator: '1' },
        conversionText: '1 K',
        description: 'SI base unit of thermodynamic temperature',
      },
      {
        name: 'Celsius',
        input: ['C', '\u00b0C'],
        display: '\u00b0C',
        conversion: { unit: 'K', numerator: '1', offset: '273.15' },
        conversionText: 'K = \u00b0C + 273.15',
        description:
          'Celsius temperature is offset from thermodynamic temperature by exactly 273.15 K; temperature intervals have the same numerical value in \u00b0C and K',
      },
      {
        name: 'Fahrenheit',
        input: ['F', '\u00b0F'],
        display: '\u00b0F',
        conversion: { unit: 'K', numerator: '5', denominator: '9', offset: '459.67' },
        conversionText: 'K = (\u00b0F + 459.67) \u00d7 5/9',
        description:
          'Fahrenheit temperature is offset and scaled relative to kelvin; 32 \u00b0F is 273.15 K',
      },
    ],
  }),
  volume: freezeCategory({
    name: 'Volume',
    canonicalUnit: 'm\u00b3',
    units: [
      {
        name: 'cubic meter',
        input: ['m\u00b3'],
        display: 'm\u00b3',
        conversion: { unit: 'm^3', numerator: '1' },
        conversionText: '1 m\u00b3',
        description: 'SI derived unit of volume',
      },
      {
        name: 'liter',
        input: ['L', 'l'],
        display: 'L',
        conversion: { unit: 'm^3', numerator: '0.001' },
        conversionText: '0.001 m\u00b3',
        description: '1 liter is exactly 0.001 cubic meters',
      },
      {
        name: 'fluid ounce (US)',
        input: ['fl oz'],
        display: 'fl oz',
        conversion: { unit: 'm^3', numerator: '0.0000295735295625' },
        conversionText: '0.0000295735295625 m\u00b3',
        description: '1 US fluid ounce is 1/128 US gallon',
      },
      {
        name: 'pint (US liquid)',
        input: ['pt'],
        display: 'pt',
        conversion: { unit: 'm^3', numerator: '0.000473176473' },
        conversionText: '0.000473176473 m\u00b3',
        description: '1 US liquid pint is 1/8 US gallon',
      },
      {
        name: 'quart (US liquid)',
        input: ['qt'],
        display: 'qt',
        conversion: { unit: 'm^3', numerator: '0.000946352946' },
        conversionText: '0.000946352946 m\u00b3',
        description: '1 US liquid quart is 1/4 US gallon',
      },
      {
        name: 'gallon (US liquid)',
        input: ['gal'],
        display: 'gal',
        conversion: { unit: 'm^3', numerator: '0.003785411784' },
        conversionText: '0.003785411784 m\u00b3',
        description: '1 US liquid gallon is exactly 231 cubic inches',
      },
      {
        name: 'barrel (petroleum)',
        input: ['bbl'],
        display: 'bbl',
        conversion: { unit: 'm^3', numerator: '0.158987294928' },
        conversionText: '0.158987294928 m\u00b3',
        description: '1 petroleum barrel is exactly 42 US liquid gallons',
      },
    ],
  }),
  electrical: freezeCategory({
    name: 'Electrical',
    canonicalUnit: 'A',
    units: [
      {
        name: 'ampere',
        input: ['A'],
        display: 'A',
        conversion: { unit: 'A', numerator: '1' },
        conversionText: '1 A',
        description: 'SI base unit of electric current',
      },
      {
        name: 'coulomb',
        input: ['C'],
        display: 'C',
        conversion: { unit: 'C', numerator: '1' },
        conversionText: '1 C',
        description: 'SI derived unit of electric charge, equal to A s',
      },
      {
        name: 'volt',
        input: ['V'],
        display: 'V',
        conversion: { unit: 'V', numerator: '1' },
        conversionText: '1 V',
        description: 'SI derived unit of electric potential difference',
      },
      {
        name: 'watt',
        input: ['W'],
        display: 'W',
        conversion: { unit: 'W', numerator: '1' },
        conversionText: '1 W',
        description: 'SI derived unit of power',
      },
      {
        name: 'ohm',
        input: ['\u03a9'],
        display: '\u03a9',
        conversion: { unit: 'ohm', numerator: '1' },
        conversionText: '1 \u03a9',
        description: 'SI derived unit of electrical resistance',
      },
      {
        name: 'farad',
        input: ['F'],
        display: 'F',
        conversion: { unit: 'F', numerator: '1' },
        conversionText: '1 F',
        description: 'SI derived unit of electrical capacitance',
      },
      {
        name: 'henry',
        input: ['H'],
        display: 'H',
        conversion: { unit: 'H', numerator: '1' },
        conversionText: '1 H',
        description: 'SI derived unit of electrical inductance',
      },
    ],
  }),
  dataStorage: freezeCategory({
    name: 'Data Storage',
    canonicalUnit: 'b',
    units: [
      {
        name: 'bit',
        input: ['b'],
        display: 'b',
        conversion: { unit: 'bit', numerator: '1' },
        conversionText: '1 b',
        description: 'Canonical unit of binary information',
      },
      {
        name: 'byte',
        input: ['B'],
        display: 'B',
        conversion: { unit: 'bit', numerator: '8' },
        conversionText: '8 b',
        description: '1 byte is equal to 8 bits',
      },
      {
        name: 'word',
        input: ['w'],
        display: 'w',
        conversion: { unit: 'bit', numerator: '16', requiresWordSize: true },
        conversionText: '16 b',
        description:
          '1 word is equal to 16 bits; this is a common size for a machine word in computing, but the actual size of a word can vary depending on the architecture',
      },
      {
        name: 'nibble',
        input: ['n'],
        display: 'n',
        conversion: { unit: 'bit', numerator: '4' },
        conversionText: '4 b',
        description: '1 nibble is equal to 4 bits; also known as a half-byte',
      },
      {
        name: 'octet',
        input: ['o'],
        display: 'o',
        conversion: { unit: 'bit', numerator: '8' },
        conversionText: '8 b',
        description:
          '1 octet is equal to 8 bits; this term is often used in networking contexts to avoid ambiguity about the size of a byte',
      },
      {
        name: 'short word',
        input: ['sw'],
        display: 'sw',
        conversion: { unit: 'bit', numerator: '32', requiresWordSize: true },
        conversionText: '32 b',
        description:
          '1 short word is equal to 32 bits; this is a common size for a short word in computing, but the actual size can vary depending on the architecture',
      },
      {
        name: 'long word',
        input: ['lw'],
        display: 'lw',
        conversion: { unit: 'bit', numerator: '64', requiresWordSize: true },
        conversionText: '64 b',
        description:
          '1 long word is equal to 64 bits; this is a common size for a long word in computing, but the actual size can vary depending on the architecture',
      },
    ],
  }),
  energy: freezeCategory({
    name: 'Energy',
    canonicalUnit: 'J',
    units: [
      {
        name: 'joule',
        input: ['J'],
        display: 'J',
        conversion: { unit: 'J', numerator: '1' },
        conversionText: '1 J',
        description: 'SI derived unit of energy',
      },
      {
        name: 'calorie (thermochemical)',
        input: ['cal'],
        display: 'cal',
        conversion: { unit: 'J', numerator: '4.184' },
        conversionText: '4.184 J',
        description: '1 thermochemical calorie is exactly 4.184 joules',
      },
    ],
  }),
  pressure: freezeCategory({
    name: 'Pressure',
    canonicalUnit: 'Pa',
    units: [
      {
        name: 'pascal',
        input: ['Pa'],
        display: 'Pa',
        conversion: { unit: 'Pa', numerator: '1' },
        conversionText: '1 Pa',
        description: 'SI derived unit of pressure',
      },
      {
        name: 'bar',
        input: ['bar'],
        display: 'bar',
        conversion: { unit: 'Pa', numerator: '100000' },
        conversionText: '100000 Pa',
        description: '1 bar is exactly 100,000 pascals',
      },
      {
        name: 'standard atmosphere',
        input: ['atm'],
        display: 'atm',
        conversion: { unit: 'Pa', numerator: '101325' },
        conversionText: '101325 Pa',
        description: '1 standard atmosphere is exactly 101,325 pascals',
      },
      {
        name: 'torr',
        input: ['Torr', 'torr'],
        display: 'Torr',
        conversion: { unit: 'Pa', numerator: '101325', denominator: '760' },
        conversionText: '133.322368 Pa',
        description: '1 torr is 1/760 standard atmosphere; value is rounded here',
      },
      {
        name: 'psi',
        input: ['psi'],
        display: 'psi',
        conversion: { unit: 'Pa', numerator: '4.4482216152605', denominator: '0.00064516' },
        conversionText: '6894.757293 Pa',
        description: '1 psi is 1 pound-force per square inch; value is rounded here',
      },
      {
        name: 'mmHg',
        input: ['mmHg'],
        display: 'mmHg',
        conversion: { unit: 'Pa', numerator: '133.322387' },
        conversionText: '133.322387 Pa',
        description: 'Conventional millimeter of mercury; value is rounded here',
      },
      {
        name: 'inHg',
        input: ['inHg'],
        display: 'inHg',
        conversion: { unit: 'Pa', numerator: '3386.389' },
        conversionText: '3386.389 Pa',
        description: 'Conventional inch of mercury; value is rounded here',
      },
    ],
  }),
  speed: freezeCategory({
    name: 'Speed',
    canonicalUnit: 'm/s',
    units: [
      {
        name: 'meter per second',
        input: ['m/s'],
        display: 'm/s',
        conversion: { unit: 'm/s', numerator: '1' },
        conversionText: '1 m/s',
        description: 'Coherent SI unit of speed',
      },
      {
        name: 'kilometer per hour',
        input: ['km/h'],
        display: 'km/h',
        conversion: { unit: 'm/s', numerator: '1000', denominator: '3600' },
        conversionText: '0.2777777778 m/s',
        description: '1 kilometer per hour is 1/3.6 meters per second; value is rounded here',
      },
      {
        name: 'knot',
        input: ['kn', 'kt'],
        display: 'kn',
        conversion: { unit: 'm/s', numerator: '1852', denominator: '3600' },
        conversionText: '0.5144444444 m/s',
        description: '1 knot is exactly 1 nautical mile per hour; value is rounded here',
      },
      {
        name: 'mile per hour',
        input: ['mph'],
        display: 'mph',
        conversion: { unit: 'm/s', numerator: '0.44704' },
        conversionText: '0.44704 m/s',
        description: '1 international mile per hour is exactly 0.44704 meters per second',
      },
    ],
  }),
  angle: freezeCategory({
    name: 'Angle',
    canonicalUnit: 'rad',
    units: [
      {
        name: 'radian',
        input: ['rad'],
        display: 'rad',
        conversion: { unit: 'rad', numerator: '1' },
        conversionText: '1 rad',
        description: 'Coherent SI derived unit for plane angle',
      },
      {
        name: 'degree',
        input: ['deg', '\u00b0'],
        display: '\u00b0',
        conversion: { unit: 'rad', numerator: '1', denominator: '180', pi: true },
        conversionText: '\u03c0/180 rad',
        description: '1 degree is \u03c0/180 radians',
      },
      {
        name: 'arcminute',
        input: ['arcmin', "'"],
        display: '\u2032',
        conversion: { unit: 'rad', numerator: '1', denominator: '10800', pi: true },
        conversionText: '\u03c0/10800 rad',
        description: '1 arcminute is 1/60 degree',
      },
      {
        name: 'arcsecond',
        input: ['arcsec', '"'],
        display: '\u2033',
        conversion: { unit: 'rad', numerator: '1', denominator: '648000', pi: true },
        conversionText: '\u03c0/648000 rad',
        description: '1 arcsecond is 1/3600 degree',
      },
    ],
  }),
  frequency: freezeCategory({
    name: 'Frequency',
    canonicalUnit: 'Hz',
    units: [
      {
        name: 'hertz',
        input: ['Hz'],
        display: 'Hz',
        conversion: { unit: 'Hz', numerator: '1' },
        conversionText: '1 Hz',
        description: 'SI derived unit of frequency, equal to s^-1',
      },
    ],
  }),
  light: freezeCategory({
    name: 'Light',
    canonicalUnit: 'cd',
    units: [
      {
        name: 'candela',
        input: ['cd'],
        display: 'cd',
        conversion: { unit: 'cd', numerator: '1' },
        conversionText: '1 cd',
        description: 'SI base unit of luminous intensity',
      },
      {
        name: 'lumen',
        input: ['lm'],
        display: 'lm',
        conversion: { unit: 'cd rad^2', numerator: '1' },
        conversionText: '1 lm',
        description: 'SI derived unit of luminous flux, equal to cd sr',
      },
      {
        name: 'lux',
        input: ['lx'],
        display: 'lx',
        conversion: { unit: 'cd rad^2/m^2', numerator: '1' },
        conversionText: '1 lx',
        description: 'SI derived unit of illuminance, equal to lm m^-2',
      },
    ],
  }),
  force: freezeCategory({
    name: 'Force',
    canonicalUnit: 'N',
    units: [
      {
        name: 'newton',
        input: ['N'],
        display: 'N',
        conversion: { unit: 'N', numerator: '1' },
        conversionText: '1 N',
        description: 'SI derived unit of force, equal to kg m s^-2',
      },
      {
        name: 'pound-force',
        input: ['lbf'],
        display: 'lbf',
        conversion: { unit: 'N', numerator: '4.4482216152605' },
        conversionText: '4.4482216152605 N',
        description:
          '1 pound-force is exactly the force from 1 avoirdupois pound under standard gravity',
      },
    ],
  }),
  amountOfSubstance: freezeCategory({
    name: 'Amount of Substance',
    canonicalUnit: 'mol',
    units: [
      {
        name: 'mole',
        input: ['mol'],
        display: 'mol',
        conversion: { unit: 'mol', numerator: '1' },
        conversionText: '1 mol',
        description: 'SI base unit of amount of substance',
      },
    ],
  }),
});
