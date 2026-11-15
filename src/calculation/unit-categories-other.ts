// Unit catalog derived from platform-sdk at 43f4f64773c9eca230aabc390ea7b03be3c290da.
// See NOTICE.md for source attribution. Numeric conversion strings are intentionally retained.
import { freezeCategory } from './unit-definition.js';

/** Speed, angle, frequency, light, force and amount of substance. */
export const OTHER_CATEGORIES = {
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
};
