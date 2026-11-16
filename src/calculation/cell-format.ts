import { Decimal } from 'decimal.js';
import { requireDefined } from '../model/require-defined.js';
import type { CellFormat } from '../model/review.js';
import { formatCalendar } from './calendar.js';
import { type DisplayOptions, formatCellValue } from './format.js';
import { quantityInUnit } from './units.js';
import { type CellValue, isNumeric, type NumericValue } from './values.js';

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

type ExactDecimalOptions = Intl.NumberFormatOptions & {
  /** Use the Intl-parts fallback even where Intl formats decimal strings exactly. */
  forceFallback?: boolean;
};

let exactDecimals: boolean | undefined;
/** Whether this runtime's Intl formats decimal strings without converting them to numbers. */
function formatsExactDecimals(): boolean {
  exactDecimals ??=
    new Intl.NumberFormat('en-US', { useGrouping: false, maximumFractionDigits: 0 }).format(
      '9007199254740993' as unknown as number,
    ) === '9007199254740993';
  return exactDecimals;
}

/** Rounds half up to the format's fraction digits, dropping zeros beyond the minimum. */
function roundedDecimal(text: string, options: Intl.NumberFormatOptions): string {
  const D = Decimal.clone({
    precision: Math.max(text.length + 110, 128),
    rounding: Decimal.ROUND_HALF_UP,
  });
  let amount = new D(text);
  if (options.style === 'percent') amount = amount.mul(100);
  const minimum = options.minimumFractionDigits ?? (options.style === 'currency' ? 2 : 0);
  const fallbackMaximum = options.style === 'currency' ? 2 : options.style === 'percent' ? 0 : 3;
  const maximum = options.maximumFractionDigits ?? Math.max(minimum, fallbackMaximum);
  let rounded = amount.toFixed(maximum);
  if (maximum > minimum) {
    while (
      rounded.includes('.') &&
      rounded.endsWith('0') &&
      rounded.length - rounded.indexOf('.') - 1 > minimum
    )
      rounded = rounded.slice(0, -1);
    if (rounded.endsWith('.')) rounded = rounded.slice(0, -1);
  }
  return amount.isNegative() && !rounded.startsWith('-') ? '-' + rounded : rounded;
}

/** The digits of a plain decimal in the locale's numbering system, grouped as the options say. */
function localDigits(
  match: RegExpExecArray,
  locale: string,
  options: Intl.NumberFormatOptions,
): string {
  const digits = Array.from({ length: 10 }, (_, n) =>
    new Intl.NumberFormat(locale, { useGrouping: false }).format(n),
  );
  const local = (value: string) => [...value].map((d) => digits[Number(d)]).join('');
  const integerText = new Intl.NumberFormat(locale, {
    useGrouping: options.useGrouping,
    maximumFractionDigits: 0,
  }).format(BigInt(match[2]));
  const fraction = (match[3] ?? '').padEnd(options.minimumFractionDigits ?? 0, '0');
  const decimal =
    new Intl.NumberFormat(locale)
      .formatToParts('1.1' as unknown as number)
      .find((p) => p.type === 'decimal')?.value ?? '.';
  return integerText + (fraction ? decimal + local(fraction) : '');
}

/** Exact decimal-string formatting, with an Intl-parts fallback for older ICU runtimes. */
export function formatExactDecimal(
  text: string,
  locale: string,
  { forceFallback = false, ...options }: ExactDecimalOptions = {},
): string {
  const formatter = new Intl.NumberFormat(locale, options);
  if (formatsExactDecimals() && !forceFallback) return formatter.format(text as unknown as number);
  const match = /^(-?)(\d+)(?:\.(\d+))?$/.exec(roundedDecimal(text, options));
  if (!match) throw new TypeError('Invalid exact decimal text.');
  const numeric = localDigits(match, locale, options);
  // The sign, currency and percent come from formatting ±1; the number parts are replaced.
  let emitted = false;
  return formatter
    .formatToParts(match[1] === '-' ? -1 : 1)
    .map((part) => {
      if (!['integer', 'group', 'decimal', 'fraction'].includes(part.type)) return part.value;
      if (emitted) return '';
      emitted = true;
      return numeric;
    })
    .join('');
}

interface FormatContext {
  format: CellFormat;
  locale: string;
  currency: string;
  options?: DisplayOptions;
}

function formatUnitCell(value: NumericValue, context: FormatContext): string {
  const { format, options } = context;
  try {
    const unit =
      ['C', 'F'].includes(format.unit ?? '') &&
      format.expectedDimension?.toLowerCase() === 'temperature'
        ? '°' + format.unit
        : requireDefined(format.unit);
    const displayed = quantityInUnit(value, unit);
    const number =
      format.decimals === undefined
        ? formatCellValue(displayed, options)
        : formatWorksheetCell(
            displayed,
            {
              kind: 'number',
              decimals: format.decimals,
              ...(format.locale ? { locale: format.locale } : {}),
            },
            options,
          );
    return `${number} ${format.unit}`;
  } catch {
    return '#VALUE!';
  }
}

function decimalPlaces(format: CellFormat): number {
  const decimals =
    format.decimals ?? (format.kind === 'number' || format.kind === 'currency' ? 2 : 0);
  if (!Number.isInteger(decimals) || decimals < 0 || decimals > 100)
    throw new RangeError('Invalid decimal places.');
  return decimals;
}

const NUMBER_STYLES: Partial<Record<CellFormat['kind'], 'currency' | 'percent'>> = {
  currency: 'currency',
  percent: 'percent',
};

function formatDecimalCell(value: NumericValue, context: FormatContext): string {
  const { format, locale, currency, options } = context;
  const decimals = decimalPlaces(format);
  const D = Decimal.clone({
    precision: Math.min(
      100_100,
      Math.max(value.numerator.length, value.denominator.length, decimals) + 110,
    ),
    rounding: Decimal.ROUND_HALF_EVEN,
  });
  const decimal = new D(value.numerator).div(value.denominator);
  // Intl applies percent scaling exactly when it receives the original decimal string.
  const style = NUMBER_STYLES[format.kind] ?? 'decimal';
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

const CALENDAR_KINDS = ['date', 'time', 'datetime'] as const;
const isCalendarKind = (kind: CellFormat['kind']): kind is (typeof CALENDAR_KINDS)[number] =>
  (CALENDAR_KINDS as readonly string[]).includes(kind);

export function formatWorksheetCell(
  value: CellValue,
  format?: CellFormat,
  options?: DisplayOptions,
): string {
  if (!format || format.kind === 'general' || !isNumeric(value))
    return formatCellValue(value, options);
  validateNumberFormat(format.locale, format.currency);
  const context = {
    format,
    locale: format.locale ?? 'en-US',
    currency: format.currency ?? 'USD',
    options,
  };
  if (isCalendarKind(format.kind)) {
    try {
      return formatCalendar(value, format.kind, context.locale);
    } catch {
      return '#VALUE!';
    }
  }
  if (format.kind === 'unit') return formatUnitCell(value, context);
  if (format.kind === 'fraction')
    return formatCellValue(value, { ...options, notation: 'fraction' });
  return formatDecimalCell(value, context);
}
