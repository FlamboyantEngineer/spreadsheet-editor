import { CalculationError, type NumericValue } from './values.js';
import { createArithmetic } from './rational.js';

const epoch = Date.UTC(1899, 11, 31),
  dayMilliseconds = 86400000;
export function calendarSerial(year: number, month: number, day: number): NumericValue {
  if (
    ![year, month, day].every(Number.isInteger) ||
    year < 0 ||
    year > 9999 ||
    Math.abs(month) > 120000 ||
    Math.abs(day) > 3660000
  )
    throw new CalculationError('#NUM!');
  if (year < 1900) year += 1900;
  if (year === 1900 && month === 2 && day === 29) return createArithmetic().parse('60');
  const date = Date.UTC(year, month - 1, day);
  const days =
    Math.floor((date - epoch) / dayMilliseconds) + (date >= Date.UTC(1900, 2, 1) ? 1 : 0);
  if (days < 0 || days > 2958465) throw new CalculationError('#NUM!');
  return createArithmetic().parse(String(days));
}
export function calendarParts(value: NumericValue) {
  if (value.dimension) throw new CalculationError('#VALUE!');
  const n = BigInt(value.numerator),
    d = BigInt(value.denominator);
  if (n < 0n || n / d > 2958465n) throw new CalculationError('#NUM!');
  const day = Number(n / d),
    seconds = Number(((n % d) * 86400n) / d);
  const date = new Date(epoch + (day - (day >= 60 ? 1 : 0)) * dayMilliseconds);
  return {
    year: date.getUTCFullYear(),
    month: day === 60 ? 2 : date.getUTCMonth() + 1,
    day: day === 60 ? 29 : date.getUTCDate(),
    hour: Math.floor(seconds / 3600),
    minute: Math.floor(seconds / 60) % 60,
    second: seconds % 60,
    date,
  };
}
export function formatCalendar(value: NumericValue, kind: 'date' | 'time' | 'datetime'): string {
  const parts = calendarParts(value);
  const date =
    parts.year === 1900 && parts.month === 2 && parts.day === 29
      ? 'Feb 29, 1900'
      : new Intl.DateTimeFormat('en-US', {
          year: 'numeric',
          month: 'short',
          day: 'numeric',
          timeZone: 'UTC',
        }).format(parts.date);
  const time = [parts.hour, parts.minute, parts.second]
    .map((part) => String(part).padStart(2, '0'))
    .join(':');
  return kind === 'date' ? date : kind === 'time' ? time : `${date} ${time}`;
}
