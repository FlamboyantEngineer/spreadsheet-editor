import { Decimal } from 'decimal.js';
import { requireDefined } from '../model/require-defined.js';
import type { CellFormat } from '../model/review.js';
import { formatCalendar } from './calendar.js';
import { type DisplayOptions, formatCellValue } from './format.js';
import { quantityInUnit } from './units.js';
import { type CellValue, isNumeric } from './values.js';

const currencies = new Set(Intl.supportedValuesOf('currency'));
export function validateNumberFormat(locale?: string, currency?: string): void {
  if (
    locale !== undefined &&
    (typeof locale !== 'string' ||
      !locale ||
      locale.length > 128 ||
      Intl.NumberFormat.supportedLocalesOf([locale]).length !== 1)
  )
    throw new TypeError('Unsupported cell format locale.');
  if (
    currency !== undefined &&
    (typeof currency !== 'string' || !/^[A-Z]{3}$/.test(currency) || !currencies.has(currency))
  )
    throw new TypeError('Unsupported cell format currency code.');
}

/** Exact decimal-string formatting, with an Intl-parts fallback for older ICU runtimes. */
export function formatExactDecimal(
  text: string,
  locale: string,
  options: Intl.NumberFormatOptions = {},
  forceFallback = false,
): string {
  const formatter = new Intl.NumberFormat(locale, options);
  const supportsDecimal =
    new Intl.NumberFormat('en-US', { useGrouping: false, maximumFractionDigits: 0 }).format(
      '9007199254740993' as unknown as number,
    ) === '9007199254740993';
  if (supportsDecimal && !forceFallback) return formatter.format(text as unknown as number);
  const D = Decimal.clone({
    precision: Math.max(text.length + 110, 128),
    rounding: Decimal.ROUND_HALF_UP,
  });
  let amount = new D(text);
  if (options.style === 'percent') amount = amount.mul(100);
  const minimum = options.minimumFractionDigits ?? (options.style === 'currency' ? 2 : 0);
  const maximum =
    options.maximumFractionDigits ??
    Math.max(minimum, options.style === 'currency' ? 2 : options.style === 'percent' ? 0 : 3);
  text = amount.toFixed(maximum);
  if (maximum > minimum) {
    while (
      text.includes('.') &&
      text.endsWith('0') &&
      text.length - text.indexOf('.') - 1 > minimum
    )
      text = text.slice(0, -1);
    if (text.endsWith('.')) text = text.slice(0, -1);
  }
  if (amount.isNegative() && !text.startsWith('-')) text = '-' + text;
  const match = /^(-?)(\d+)(?:\.(\d+))?$/.exec(text);
  if (!match) throw new TypeError('Invalid exact decimal text.');
  const negative = match[1] === '-',
    integer = BigInt(match[2]);
  const digits = Array.from({ length: 10 }, (_, n) =>
    new Intl.NumberFormat(locale, { useGrouping: false }).format(n),
  );
  const local = (value: string) => [...value].map((d) => digits[Number(d)]).join('');
  const integerText = new Intl.NumberFormat(locale, {
    useGrouping: options.useGrouping,
    maximumFractionDigits: 0,
  }).format(integer);
  const fraction = (match[3] ?? '').padEnd(options.minimumFractionDigits ?? 0, '0');
  const decimal =
    new Intl.NumberFormat(locale)
      .formatToParts('1.1' as unknown as number)
      .find((p) => p.type === 'decimal')?.value ?? '.';
  const numeric = integerText + (fraction ? decimal + local(fraction) : '');
  const parts = formatter.formatToParts(negative ? -1 : 1);
  let emitted = false;
  return parts
    .map((part) => {
      if (['integer', 'group', 'decimal', 'fraction'].includes(part.type)) {
        if (emitted) return '';
        emitted = true;
        return numeric;
      }
      return part.value;
    })
    .join('');
}

export function formatWorksheetCell(
  value: CellValue,
  format?: CellFormat,
  options?: DisplayOptions,
): string {
  if (!format || format.kind === 'general' || !isNumeric(value))
    return formatCellValue(value, options);
  const metadata = format as CellFormat & { locale?: string; currency?: string };
  validateNumberFormat(metadata.locale, metadata.currency);
  const locale = metadata.locale ?? 'en-US',
    currency = metadata.currency ?? 'USD';
  if (format.kind === 'date' || format.kind === 'time' || format.kind === 'datetime') {
    try {
      return formatCalendar(value, format.kind, locale);
    } catch {
      return '#VALUE!';
    }
  }
  if (format.kind === 'unit') {
    try {
      const unit =
        ['C', 'F'].includes(format.unit ?? '') &&
        format.expectedDimension?.toLowerCase() === 'temperature'
          ? '°' + format.unit
          : requireDefined(format.unit);
      const displayed = quantityInUnit(value, unit);
      return `${format.decimals === undefined ? formatCellValue(displayed, options) : formatWorksheetCell(displayed, { kind: 'number', decimals: format.decimals, ...(metadata.locale ? { locale: metadata.locale } : {}) }, options)} ${format.unit}`;
    } catch {
      return '#VALUE!';
    }
  }
  if (format.kind === 'fraction')
    return formatCellValue(value, { ...options, notation: 'fraction' });
  const decimals =
    format.decimals ?? (format.kind === 'number' || format.kind === 'currency' ? 2 : 0);
  if (!Number.isInteger(decimals) || decimals < 0 || decimals > 100)
    throw new RangeError('Invalid decimal places.');
  const D = Decimal.clone({
    precision: Math.min(
      100_100,
      Math.max(value.numerator.length, value.denominator.length, decimals) + 110,
    ),
    rounding: Decimal.ROUND_HALF_EVEN,
  });
  const decimal = new D(value.numerator).div(value.denominator);
  // Intl applies percent scaling exactly when it receives the original decimal string.
  const style =
    format.kind === 'currency' ? 'currency' : format.kind === 'percent' ? 'percent' : 'decimal';
  const text =
    format.kind === 'scientific'
      ? decimal.toExponential(decimals)
      : formatExactDecimal(
          decimal.toFixed(format.kind === 'percent' ? decimals + 2 : decimals),
          locale,
          {
            style,
            ...(style === 'currency' ? { currency } : {}),
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals,
            useGrouping: format.kind === 'currency',
          },
        );
  return `${value.approximate && options?.approximationMarker !== false ? '≈' : ''}${text}`;
}
