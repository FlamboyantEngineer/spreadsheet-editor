import { type FormulaNode, FormulaSyntaxError } from '../model/formula-parser.js';
import type { CellFormat } from '../model/review.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { CellInput, Workbook } from '../model/workbook.js';
import type { CalculationOptions } from './evaluator.js';
import type { Arithmetic } from './function-call.js';
import { makeQuantity, quantityInUnit, validateUnitDimension } from './units.js';
import {
  assertNumericValue,
  CalculationError,
  type CellValue,
  errorValue,
  isNumeric,
  normalizeNumericValue,
  type NumericValue,
} from './values.js';

/** What reading a cell needs from the calculation it belongs to. */
export interface CellContext {
  readonly workbook: Workbook;
  readonly options: CalculationOptions;
  readonly math: Arithmetic;
}
export interface CellPosition {
  sheet: number;
  x: number;
  y: number;
}

/** The key of a resolved operand in `CalculationContext.operands`: its cell and formula span. */
export const formulaOperandKey = (
  { sheetId, x, y }: { sheetId: string; x: number; y: number },
  { start, end }: { start: number; end: number },
) => JSON.stringify([sheetId, x, y, start, end]);

/** The internal `sheet:x:y` key; the public `[sheetId,x,y]` key is written once per result. */
export const keyFor = ({ sheet, x, y }: CellPosition) => `${sheet}:${x}:${y}`;
/** The memo of one formula cell, made on first use. */
export function memoFor<V>(map: Map<string, WeakMap<FormulaNode, V>>, key: string) {
  let memo = map.get(key);
  if (!memo) {
    memo = new WeakMap();
    map.set(key, memo);
  }
  return memo;
}

const NUMBER_TEXT = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const QUANTITY_TEXT = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s+(.+)$/;

export const isNumericInputText = (text: string) => NUMBER_TEXT.test(text.trim());

/** The error value a failed calculation step yields. */
export function errorValueOf(error: unknown) {
  if (error instanceof CalculationError) return errorValue(error.code);
  return errorValue(error instanceof FormulaSyntaxError ? error.token : '#ERROR!');
}

export const isFormula = (input: CellInput): input is string =>
  typeof input === 'string' && input.trimStart().startsWith('=');

/** The value of a cell that holds no formula. */
export function authoredValue(
  context: CellContext,
  input: CellInput,
  canonical: NumericValue | undefined,
): CellValue {
  if (canonical) {
    assertNumericValue(canonical, context.options.maxIntegerDigits);
    return normalizeNumericValue(canonical);
  }
  if (typeof input === 'number') return context.math.parse(String(input));
  if (typeof input === 'string' && input.startsWith("'")) return input.slice(1);
  if (typeof input === 'string' && isNumericInputText(input)) return context.math.parse(input);
  return input === '' ? null : input;
}

interface UnitFormat {
  format: CellFormat;
  /** The format's unit, as authored. */
  unit: string;
  /** The unit values are shown in. */
  display: string;
  /** The cell holds a canonical input value. */
  canonical: boolean;
}

/** Temperatures are written `C` and `F` in a temperature format. */
function configuredUnit(format: CellFormat, unit: string): string {
  return ['C', 'F'].includes(unit) && format.expectedDimension?.toLowerCase() === 'temperature'
    ? '°' + unit
    : unit;
}

/** Text such as `3 kg` becomes a quantity when it fits the format; other text stays text. */
function quantityFromText(context: CellContext, text: string, unit: UnitFormat): CellValue {
  const match = QUANTITY_TEXT.exec(text.trim());
  if (!match) return text;
  try {
    const candidate = makeQuantity(context.math.parse(match[1]), match[2], {
      ...context.options,
      quantityKind: unit.format.quantityKind,
    });
    quantityInUnit(candidate, unit.display, context.options);
    return candidate;
  } catch {
    return text;
  }
}

function numberInUnit(context: CellContext, value: NumericValue, unit: UnitFormat): NumericValue {
  const { format } = unit;
  let result = value;
  if (!result.dimension && !result.quantityKind && !unit.canonical)
    result = makeQuantity(result, configuredUnit(format, format.inputUnit ?? unit.unit), {
      ...context.options,
      quantityKind: format.quantityKind,
    });
  if (format.quantityKind && result.quantityKind && result.quantityKind !== format.quantityKind)
    throw new CalculationError('#VALUE!');
  if (format.quantityKind && !result.quantityKind)
    result = { ...result, quantityKind: format.quantityKind };
  quantityInUnit(result, unit.display, context.options);
  return result;
}

/** Applies a cell's unit format: plain numbers take its input unit, and the value must fit it. */
export function withUnitFormat(
  context: CellContext,
  cell: CellPosition,
  value: CellValue,
): CellValue {
  const sheet = context.workbook.worksheets[cell.sheet];
  const address = formatSpreadsheetAddress(cell.x, cell.y);
  const format = sheet.formats?.[address];
  if (format?.kind !== 'unit' || !format.unit) return value;
  const unit: UnitFormat = {
    format,
    unit: format.unit,
    display: configuredUnit(format, format.unit),
    canonical: Boolean(sheet.inputValues?.[address]),
  };
  validateUnitDimension(unit.display, format.expectedDimension, context.options);
  if (typeof value === 'string') return quantityFromText(context, value, unit);
  if (isNumeric(value)) return numberInUnit(context, value, unit);
  return value;
}
