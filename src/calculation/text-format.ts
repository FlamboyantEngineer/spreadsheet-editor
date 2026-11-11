import { Decimal } from 'decimal.js';
import { calendarParts } from './calendar.js';
import { CalculationError, type NumericValue } from './values.js';

/** Deliberately bounded Excel TEXT subset; unsupported masks fail explicitly. */
export function formatTextMask(value: NumericValue, mask: string): string {
  if (value.dimension || value.quantityKind || mask.length > 128)
    throw new CalculationError('#VALUE!');
  const dateMasks = new Set([
    'yyyy-mm-dd',
    'mm/dd/yyyy',
    'm/d/yyyy',
    'dd/mm/yyyy',
    'd-mmm-yyyy',
    'mmm d, yyyy',
    'hh:mm:ss',
    'hh:mm',
    'h:mm:ss',
    'h:mm',
  ]);
  const normalized = mask.toLowerCase();
  if (dateMasks.has(normalized)) {
    const p = calendarParts(value),
      pad = (n: number) => String(n).padStart(2, '0');
    if (normalized.includes('h'))
      return `${normalized.startsWith('hh') ? pad(p.hour) : p.hour}:${pad(p.minute)}${normalized.endsWith(':ss') ? `:${pad(p.second)}` : ''}`;
    const months = [
      'Jan',
      'Feb',
      'Mar',
      'Apr',
      'May',
      'Jun',
      'Jul',
      'Aug',
      'Sep',
      'Oct',
      'Nov',
      'Dec',
    ];
    switch (normalized) {
      case 'yyyy-mm-dd':
        return `${String(p.year).padStart(4, '0')}-${pad(p.month)}-${pad(p.day)}`;
      case 'mm/dd/yyyy':
        return `${pad(p.month)}/${pad(p.day)}/${p.year}`;
      case 'm/d/yyyy':
        return `${p.month}/${p.day}/${p.year}`;
      case 'dd/mm/yyyy':
        return `${pad(p.day)}/${pad(p.month)}/${p.year}`;
      case 'd-mmm-yyyy':
        return `${p.day}-${months[p.month - 1]}-${p.year}`;
      default:
        return `${months[p.month - 1]} ${p.day}, ${p.year}`;
    }
  }
  const match = /^(\$)?(0|#,##0)(?:\.(0{1,100}))?(%)?$/.exec(mask);
  if (!match || (match[1] && match[4])) throw new CalculationError('#VALUE!');
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
