import { Decimal } from 'decimal.js';
import { calendarParts } from './calendar.js';
import { CalculationError, type NumericValue } from './values.js';

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
type Parts = ReturnType<typeof calendarParts>;
const pad = (n: number) => String(n).padStart(2, '0');
const TIME_MASKS = new Set(['hh:mm:ss', 'hh:mm', 'h:mm:ss', 'h:mm']);
const DATE_MASKS: Readonly<Record<string, (p: Parts) => string>> = {
  'yyyy-mm-dd': (p) => `${String(p.year).padStart(4, '0')}-${pad(p.month)}-${pad(p.day)}`,
  'mm/dd/yyyy': (p) => `${pad(p.month)}/${pad(p.day)}/${p.year}`,
  'm/d/yyyy': (p) => `${p.month}/${p.day}/${p.year}`,
  'dd/mm/yyyy': (p) => `${pad(p.day)}/${pad(p.month)}/${p.year}`,
  'd-mmm-yyyy': (p) => `${p.day}-${MONTHS[p.month - 1]}-${p.year}`,
  'mmm d, yyyy': (p) => `${MONTHS[p.month - 1]} ${p.day}, ${p.year}`,
};

function timeText(p: Parts, mask: string): string {
  const hour = mask.startsWith('hh') ? pad(p.hour) : String(p.hour);
  return `${hour}:${pad(p.minute)}${mask.endsWith(':ss') ? `:${pad(p.second)}` : ''}`;
}

/** `0`, `#,##0`, either with fixed decimals, and a `$` prefix or a `%` suffix. */
function numberMaskText(value: NumericValue, match: RegExpExecArray): string {
  const D = Decimal.clone({
    precision: Math.min(100_100, Math.max(value.numerator.length, value.denominator.length) + 110),
    rounding: Decimal.ROUND_HALF_UP,
  });
  let amount = new D(value.numerator).div(value.denominator);
  if (match[4]) amount = amount.mul(100);
  let text = amount.toFixed(match[3]?.length ?? 0);
  if (match[2] === '#,##0') {
    const [integer, fraction] = text.split('.');
    text =
      integer.replace(/\B(?=(\d{3})+(?!\d))/g, ',') +
      (fraction === undefined ? '' : '.' + fraction);
  }
  return `${match[1] ?? ''}${text}${match[4] ?? ''}`;
}

/** Deliberately bounded Excel TEXT subset; unsupported masks fail explicitly. */
export function formatTextMask(value: NumericValue, mask: string): string {
  if (value.dimension || value.quantityKind || mask.length > 128)
    throw new CalculationError('#VALUE!');
  const normalized = mask.toLowerCase();
  if (TIME_MASKS.has(normalized)) return timeText(calendarParts(value), normalized);
  const date = DATE_MASKS[normalized];
  if (date) return date(calendarParts(value));
  const match = /^(\$)?(0|#,##0)(?:\.(0{1,100}))?(%)?$/.exec(mask);
  if (!match || (match[1] && match[4])) throw new CalculationError('#VALUE!');
  return numberMaskText(value, match);
}

const windows1252 = [
  '€',
  '\u0081',
  '‚',
  'ƒ',
  '„',
  '…',
  '†',
  '‡',
  'ˆ',
  '‰',
  'Š',
  '‹',
  'Œ',
  '\u008d',
  'Ž',
  '\u008f',
  '\u0090',
  '‘',
  '’',
  '“',
  '”',
  '•',
  '–',
  '—',
  '˜',
  '™',
  'š',
  '›',
  'œ',
  '\u009d',
  'ž',
  'Ÿ',
];
export function char1252(code: number): string {
  return code >= 128 && code <= 159 ? windows1252[code - 128] : String.fromCharCode(code);
}
export function code1252(text: string): number {
  const character = text[0],
    index = windows1252.indexOf(character);
  if (index >= 0) return index + 128;
  const code = character.charCodeAt(0);
  if (code > 255) throw new CalculationError('#VALUE!');
  return code;
}
