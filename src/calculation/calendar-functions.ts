import { type EagerCall, type EagerFunction, requireArity } from './builtin-call.js';
import { calendarParts, calendarSerial } from './calendar.js';
import type { Value } from './function-call.js';
import { CalculationError, type NumericValue } from './values.js';

const SECONDS_PER_DAY = 86400;

/** A whole number argument, such as a year or a count of seconds. */
function whole(call: EagerCall, value: Value): number {
  const n = call.evaluation.values.numeric(value);
  if (n.dimension || n.quantityKind) throw new CalculationError('#VALUE!');
  const result = Number(BigInt(n.numerator) / BigInt(n.denominator));
  if (!Number.isSafeInteger(result)) throw new CalculationError('#NUM!');
  return result;
}

const dayFraction = (call: EagerCall, seconds: number): NumericValue =>
  call.evaluation.math.divide(
    call.evaluation.math.parse(String(seconds)),
    call.evaluation.math.parse(String(SECONDS_PER_DAY)),
  );

function date(call: EagerCall) {
  requireArity(call, 3);
  const { args } = call;
  return calendarSerial(whole(call, args[0]), whole(call, args[1]), whole(call, args[2]));
}

function time(call: EagerCall) {
  requireArity(call, 3);
  const { args } = call;
  const seconds = whole(call, args[0]) * 3600 + whole(call, args[1]) * 60 + whole(call, args[2]);
  if (seconds < 0 || !Number.isSafeInteger(seconds)) throw new CalculationError('#NUM!');
  return dayFraction(call, seconds % SECONDS_PER_DAY);
}

/** The pass's clock in the calculation's time zone, captured once for every TODAY and NOW. */
function clockFields(call: EagerCall): Record<string, string> {
  const { evaluation } = call;
  evaluation.clock ??= evaluation.host.now?.() ?? new Date();
  return Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: evaluation.options.timeZone ?? 'UTC',
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      hourCycle: 'h23',
    })
      .formatToParts(evaluation.clock)
      .map((part) => [part.type, part.value]),
  );
}

function now(call: EagerCall) {
  requireArity(call, 0);
  const fields = clockFields(call);
  const today = calendarSerial(Number(fields.year), Number(fields.month), Number(fields.day));
  if (call.name === 'TODAY') return today;
  const seconds = Number(fields.hour) * 3600 + Number(fields.minute) * 60 + Number(fields.second);
  return call.evaluation.math.add(today, dayFraction(call, seconds));
}

function days(call: EagerCall) {
  requireArity(call, 2);
  const { math } = call.evaluation;
  return math.subtract(
    math.parse(String(whole(call, call.args[0]))),
    math.parse(String(whole(call, call.args[1]))),
  );
}

const PARTS = {
  YEAR: 'year',
  MONTH: 'month',
  DAY: 'day',
  HOUR: 'hour',
  MINUTE: 'minute',
  SECOND: 'second',
} as const;

function part(call: EagerCall) {
  requireArity(call, 1);
  const parts = calendarParts(call.evaluation.values.numeric(call.args[0]));
  return call.evaluation.math.parse(String(parts[PARTS[call.name as keyof typeof PARTS]]));
}

export const CALENDAR_FUNCTIONS: Readonly<Record<string, EagerFunction>> = {
  DATE: date,
  TIME: time,
  TODAY: now,
  NOW: now,
  DAYS: days,
  YEAR: part,
  MONTH: part,
  DAY: part,
  HOUR: part,
  MINUTE: part,
  SECOND: part,
};
