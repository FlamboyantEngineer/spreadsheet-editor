import { LOOKUP_FUNCTIONS } from './extended-lookup.js';
import { MATH_FUNCTIONS } from './extended-math.js';
import { STATISTICS_FUNCTIONS } from './extended-statistics.js';
import { TEXT_FUNCTIONS } from './extended-text.js';
import type { FunctionImpl } from './function-call.js';

export const FUNCTION_ALIASES: Readonly<Record<string, string>> = Object.freeze({
  AVG: 'AVERAGE',
  STDEV: 'STDEV.S',
  STDEVS: 'STDEV.S',
  STDEVP: 'STDEV.P',
  VAR: 'VAR.S',
  VARS: 'VAR.S',
  VARP: 'VAR.P',
  'CEILING.MATH': 'CEILINGMATH',
  'CEILING.PRECISE': 'CEILINGPRECISE',
  'FLOOR.MATH': 'FLOORMATH',
  'FLOOR.PRECISE': 'FLOORPRECISE',
  'RANK.AVG': 'RANKAVG',
  'RANK.EQ': 'RANKEQ',
});
/** The functions beyond the original catalog, by name; each takes its evaluated arguments. */
export const EXTENDED_FUNCTION_TABLE: Readonly<Record<string, FunctionImpl>> = {
  ...STATISTICS_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  ...MATH_FUNCTIONS,
};
