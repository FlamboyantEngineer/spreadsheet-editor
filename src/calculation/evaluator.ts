import { parseFormula, FormulaSyntaxError, type FormulaNode } from '../model/formula-parser.js';
import {
  assertWorkbook,
  type CellInput,
  type Workbook,
  type WorkbookLimits,
} from '../model/workbook.js';
import { createArithmetic, type NumericOptions } from './rational.js';
import { formatCellValue } from './format.js';
import {
  CalculationError,
  errorValue,
  isError,
  isNumeric,
  type CellValue,
  type ErrorValue,
  type NumericValue,
} from './values.js';

export interface CalculationOptions extends NumericOptions, WorkbookLimits {
  maxOperations?: number;
  maxDepth?: number;
}
export interface CalculatedWorksheet {
  id: string;
  values: CellValue[][];
}
export interface CalculationResult {
  schemaVersion: 1;
  worksheets: CalculatedWorksheet[];
}
export interface FormulaHost {
  readName?(name: string): CellValue | undefined;
  /** Host functions are synchronous and share this calculation's execution budget. */
  callFunction?(
    name: string,
    args: readonly CellValue[],
    context: { sheetId: string; x: number; y: number },
  ): CellValue | undefined;
}
type Value = CellValue | CellValue[][];

export const SUPPORTED_FUNCTIONS = Object.freeze([
  'IF',
  'IFERROR',
  'AND',
  'OR',
  'NOT',
  'TRUE',
  'FALSE',
  'SUM',
  'PRODUCT',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'ABS',
  'SIGN',
  'SQRT',
  'POWER',
  'MOD',
  'ROUND',
  'ROUNDUP',
  'ROUNDDOWN',
  'INT',
  'TRUNC',
  'CEILING',
  'FLOOR',
  'QUOTIENT',
  'PI',
  'SIN',
  'COS',
  'TAN',
  'ASIN',
  'ACOS',
  'ATAN',
  'EXP',
  'LN',
  'LOG',
  'LOG10',
  'ROW',
  'COLUMN',
  'ROWS',
  'COLUMNS',
  'LEN',
  'LEFT',
  'RIGHT',
  'MID',
  'LOWER',
  'UPPER',
  'TRIM',
  'CONCAT',
  'CONCATENATE',
  'VALUE',
  'ISNUMBER',
  'ISTEXT',
  'ISBLANK',
  'ISERROR',
  'NA',
  'SUMSQ',
  'VAR.S',
  'VAR.P',
  'STDEV.S',
  'STDEV.P',
]);

/** Pure, deterministic workbook evaluation; this module imports no browser or UI code. */
export function calculateWorkbook(
  workbook: Workbook,
  options: CalculationOptions = {},
  host: FormulaHost = {},
): CalculationResult {
  assertWorkbook(workbook, options);
  const arithmetic = createArithmetic(options);
  const maxOperations = options.maxOperations ?? 200_000,
    maxDepth = options.maxDepth ?? 128;
  if (
    !Number.isInteger(maxOperations) ||
    maxOperations < 1 ||
    maxOperations > 10_000_000 ||
    !Number.isInteger(maxDepth) ||
    maxDepth < 1 ||
    maxDepth > 512
  )
    throw new RangeError('Invalid calculation limits.');
  let operations = 0;
  const cache = new Map<string, CellValue>(),
    visiting = new Set<string>();
  const parsed = new Map<string, FormulaNode>();
  const zero = arithmetic.parse('0'),
    one = arithmetic.parse('1');
  const flattened = (values: readonly Value[]) =>
    values.flatMap((value) => (Array.isArray(value) ? value.flat() : [value]));
  function tick(depth: number) {
    if (++operations > maxOperations || depth > maxDepth) throw new CalculationError('#LIMIT!');
  }
  function scalar(value: Value): CellValue {
    return Array.isArray(value) ? (value[0]?.[0] ?? null) : value;
  }
  function checked(value: CellValue): Exclude<CellValue, ErrorValue> {
    if (isError(value)) throw new CalculationError(value.code);
    return value;
  }
  function numeric(value: Value): NumericValue {
    const item = checked(scalar(value));
    if (isNumeric(item)) return item;
    if (item === null) return zero;
    if (typeof item === 'boolean') return item ? one : zero;
    return arithmetic.parse(item);
  }
  function truth(value: Value): boolean {
    const item = checked(scalar(value));
    return isNumeric(item)
      ? item.numerator !== '0'
      : item === null
        ? false
        : typeof item === 'boolean'
          ? item
          : item.toUpperCase() === 'TRUE'
            ? true
            : item.toUpperCase() === 'FALSE'
              ? false
              : numeric(item).numerator !== '0';
  }
  function text(value: Value): string {
    return formatCellValue(checked(scalar(value)), { approximationMarker: false });
  }
  function read(sheetIndex: number, x: number, y: number, depth: number): CellValue {
    tick(depth);
    if (
      !Number.isSafeInteger(x) ||
      !Number.isSafeInteger(y) ||
      x < 0 ||
      y < 0 ||
      x > 16_383 ||
      y > 1_048_575
    )
      return errorValue('#REF!');
    const key = `${sheetIndex}:${x}:${y}`;
    if (cache.has(key)) return cache.get(key)!;
    if (visiting.has(key)) return errorValue('#CYCLE!');
    visiting.add(key);
    let result: CellValue;
    try {
      const input: CellInput = workbook.worksheets[sheetIndex].data[y]?.[x] ?? null;
      if (typeof input === 'string' && input.trimStart().startsWith('=')) {
        const formula = input.trimStart().slice(1);
        let tree = parsed.get(formula);
        if (!tree) {
          tree = parseFormula(formula);
          parsed.set(formula, tree);
        }
        result = scalar(evaluate(tree, sheetIndex, x, y, depth + 1));
      } else if (typeof input === 'number') result = arithmetic.parse(String(input));
      else if (typeof input === 'string' && input.startsWith("'")) result = input.slice(1);
      else if (
        typeof input === 'string' &&
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(input.trim())
      )
        result = arithmetic.parse(input);
      else result = input === '' ? null : input;
    } catch (error) {
      result = errorValue(
        error instanceof CalculationError
          ? error.code
          : error instanceof FormulaSyntaxError
            ? '#ERROR!'
            : '#ERROR!',
      );
    } finally {
      visiting.delete(key);
    }
    cache.set(key, result);
    return result;
  }
  function evaluate(
    node: FormulaNode,
    sheetIndex: number,
    x: number,
    y: number,
    depth: number,
  ): Value {
    tick(depth);
    const run = (child: FormulaNode) => evaluate(child, sheetIndex, x, y, depth + 1);
    switch (node.kind) {
      case 'literal': {
        if (
          typeof node.value === 'string' &&
          /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|NULL!)/.test(node.value)
        ) {
          throw new CalculationError(
            node.value === '#NULL!' ? '#VALUE!' : (node.value as CalculationError['code']),
          );
        }
        return node.value;
      }
      case 'decimal':
        return arithmetic.parse(node.decimal);
      case 'quantity':
        throw new CalculationError('#VALUE!');
      case 'name': {
        const value = host.readName?.(node.name);
        if (value === undefined) throw new CalculationError('#NAME?');
        return value;
      }
      case 'reference':
      case 'lines': {
        const target =
          node.sheetName === null
            ? sheetIndex
            : workbook.worksheets.findIndex(
                (sheet) => sheet.name.toUpperCase() === node.sheetName!.toUpperCase(),
              );
        if (target < 0) throw new CalculationError('#REF!');
        const sheet = workbook.worksheets[target];
        const width = sheet.data.reduce((max, row) => Math.max(max, row.length), 0);
        if (node.kind === 'lines' && (!width || !sheet.data.length)) return [];
        const start =
          node.kind === 'reference'
            ? node.start
            : node.axis === 'column'
              ? { x: Math.min(node.start, node.end), y: 0 }
              : { x: 0, y: Math.min(node.start, node.end) };
        const end =
          node.kind === 'reference'
            ? (node.end ?? node.start)
            : node.axis === 'column'
              ? { x: Math.max(node.start, node.end), y: sheet.data.length - 1 }
              : { x: width - 1, y: Math.max(node.start, node.end) };
        if (node.kind === 'reference' && !node.end)
          return checked(read(target, start.x, start.y, depth + 1));
        const left = Math.min(start.x, end.x),
          top = Math.min(start.y, end.y),
          right = Math.max(start.x, end.x),
          bottom = Math.max(start.y, end.y);
        if ((right - left + 1) * (bottom - top + 1) > maxOperations - operations)
          throw new CalculationError('#LIMIT!');
        const rows: CellValue[][] = [];
        for (let row = top; row <= bottom; row++) {
          const cells: CellValue[] = [];
          for (let col = left; col <= right; col++) cells.push(read(target, col, row, depth + 1));
          rows.push(cells);
        }
        return rows;
      }
      case 'unary': {
        const value = numeric(run(node.operand));
        return node.operator === '-'
          ? arithmetic.negate(value)
          : node.operator === '%'
            ? arithmetic.divide(value, arithmetic.parse('100'))
            : value;
      }
      case 'binary': {
        const left = run(node.left),
          right = run(node.right);
        if (node.operator === '&') return text(left) + text(right);
        if (['=', '<>', '<', '<=', '>', '>='].includes(node.operator)) {
          const a = checked(scalar(left)),
            b = checked(scalar(right));
          const comparison =
            isNumeric(a) || isNumeric(b) || a === null || b === null
              ? arithmetic.compare(numeric(a), numeric(b))
              : String(a).toUpperCase() < String(b).toUpperCase()
                ? -1
                : String(a).toUpperCase() === String(b).toUpperCase()
                  ? 0
                  : 1;
          switch (node.operator) {
            case '=':
              return comparison === 0;
            case '<>':
              return comparison !== 0;
            case '<':
              return comparison < 0;
            case '<=':
              return comparison <= 0;
            case '>':
              return comparison > 0;
            case '>=':
              return comparison >= 0;
          }
        }
        const a = numeric(left),
          b = numeric(right);
        switch (node.operator) {
          case '+':
            return arithmetic.add(a, b);
          case '-':
            return arithmetic.subtract(a, b);
          case '*':
            return arithmetic.multiply(a, b);
          case '/':
            return arithmetic.divide(a, b);
          case '^':
            return arithmetic.power(a, b);
          case '%':
            return arithmetic.subtract(
              a,
              arithmetic.multiply(b, arithmetic.round(arithmetic.divide(a, b), 0, 'down')),
            );
        }
        throw new CalculationError('#ERROR!');
      }
      case 'call': {
        const name = node.name.toUpperCase();
        const count = node.args.length;
        function arity(min: number, max = min) {
          if (count < min || count > max) throw new CalculationError('#VALUE!');
        }
        if (name === 'IF') {
          arity(2, 3);
          return truth(run(node.args[0]))
            ? run(node.args[1])
            : node.args[2]
              ? run(node.args[2])
              : false;
        }
        if (name === 'IFERROR') {
          arity(2);
          try {
            const result = run(node.args[0]);
            checked(scalar(result));
            return result;
          } catch (error) {
            if (error instanceof CalculationError) return run(node.args[1]);
            throw error;
          }
        }
        if (name === 'ISERROR') {
          arity(1);
          try {
            return isError(scalar(run(node.args[0])));
          } catch (error) {
            if (error instanceof CalculationError) return true;
            throw error;
          }
        }
        if (name === 'ROW' || name === 'COLUMN') {
          arity(0, 1);
          if (!count) return arithmetic.parse(String((name === 'ROW' ? y : x) + 1));
          const reference = node.args[0];
          if (reference.kind !== 'reference') throw new CalculationError('#VALUE!');
          return arithmetic.parse(
            String((name === 'ROW' ? reference.start.y : reference.start.x) + 1),
          );
        }
        const args = node.args.map(run),
          flat = flattened(args);
        if (name === 'TRUE' || name === 'FALSE') {
          arity(0);
          return name === 'TRUE';
        }
        if (name === 'NA') {
          arity(0);
          throw new CalculationError('#N/A');
        }
        if (name === 'AND' || name === 'OR') {
          arity(1, 255);
          return name === 'AND' ? flat.every(truth) : flat.some(truth);
        }
        if (name === 'NOT') {
          arity(1);
          return !truth(args[0]);
        }
        if (
          [
            'SUM',
            'PRODUCT',
            'AVERAGE',
            'MIN',
            'MAX',
            'COUNT',
            'COUNTA',
            'COUNTBLANK',
            'SUMSQ',
            'VAR.S',
            'VAR.P',
            'STDEV.S',
            'STDEV.P',
          ].includes(name)
        ) {
          arity(name === 'SUM' ? 0 : 1, 255);
          flat.forEach(checked);
          const values = flat.filter(isNumeric);
          if (name === 'COUNT') return arithmetic.parse(String(values.length));
          if (name === 'COUNTA')
            return arithmetic.parse(String(flat.filter((value) => value !== null).length));
          if (name === 'COUNTBLANK')
            return arithmetic.parse(
              String(flat.filter((value) => value === null || value === '').length),
            );
          const sum = values.reduce(arithmetic.add, zero);
          if (name === 'SUM') return sum;
          if (name === 'PRODUCT')
            return values.length ? values.reduce(arithmetic.multiply, one) : zero;
          if (name === 'SUMSQ')
            return values.reduce(
              (sum, value) => arithmetic.add(sum, arithmetic.multiply(value, value)),
              zero,
            );
          if (name === 'AVERAGE')
            return arithmetic.divide(sum, arithmetic.parse(String(values.length)));
          if (name === 'MIN' || name === 'MAX')
            return values.reduce(
              (a, b) => (arithmetic.compare(a, b) > 0 === (name === 'MIN') ? b : a),
              values[0] ?? zero,
            );
          const denominator = values.length - (name.endsWith('.S') ? 1 : 0);
          if (denominator <= 0) throw new CalculationError('#DIV/0!');
          const average = arithmetic.divide(sum, arithmetic.parse(String(values.length)));
          const squared = values.reduce((sum, value) => {
            const delta = arithmetic.subtract(value, average);
            return arithmetic.add(sum, arithmetic.multiply(delta, delta));
          }, zero);
          const variance = arithmetic.divide(squared, arithmetic.parse(String(denominator)));
          return name.startsWith('STDEV') ? arithmetic.sqrt(variance) : variance;
        }
        if (name === 'ISNUMBER' || name === 'ISTEXT' || name === 'ISBLANK') {
          arity(1);
          const value = checked(scalar(args[0]));
          return name === 'ISNUMBER'
            ? isNumeric(value)
            : name === 'ISTEXT'
              ? typeof value === 'string'
              : value === null;
        }
        if (name === 'ROWS' || name === 'COLUMNS') {
          arity(1);
          const range = args[0];
          return arithmetic.parse(
            String(
              Array.isArray(range) ? (name === 'ROWS' ? range.length : (range[0]?.length ?? 0)) : 1,
            ),
          );
        }
        if (name === 'CONCAT' || name === 'CONCATENATE') {
          arity(1, 255);
          return flat.map(text).join('');
        }
        if (['LEN', 'LOWER', 'UPPER', 'TRIM', 'LEFT', 'RIGHT', 'MID', 'VALUE'].includes(name)) {
          arity(
            name === 'MID' ? 3 : 1,
            ['LEFT', 'RIGHT'].includes(name) ? 2 : name === 'MID' ? 3 : 1,
          );
          const source = text(args[0]);
          if (name === 'VALUE') return arithmetic.parse(source);
          if (name === 'LEN') return arithmetic.parse(String(source.length));
          if (name === 'LOWER') return source.toLowerCase();
          if (name === 'UPPER') return source.toUpperCase();
          if (name === 'TRIM') return source.trim().replace(/ +/g, ' ');
          const count =
            args[1] === undefined ? 1 : arithmetic.integer(numeric(args[1]), 0, 1_000_000);
          if (name === 'LEFT') return source.slice(0, count);
          if (name === 'RIGHT') return count === 0 ? '' : source.slice(-count);
          if (count < 1) throw new CalculationError('#VALUE!');
          return source.slice(
            count - 1,
            count - 1 + arithmetic.integer(numeric(args[2]), 0, 1_000_000),
          );
        }
        if (name === 'PI') {
          arity(0);
          return arithmetic.pi();
        }
        if (['ROUND', 'ROUNDUP', 'ROUNDDOWN', 'TRUNC', 'INT'].includes(name)) {
          arity(['TRUNC', 'INT'].includes(name) ? 1 : 2, name === 'INT' ? 1 : 2);
          const places = args[1] === undefined ? 0 : arithmetic.integer(numeric(args[1]));
          return arithmetic.round(
            numeric(args[0]),
            places,
            name === 'INT'
              ? 'floor'
              : name === 'ROUNDUP'
                ? 'up'
                : name === 'ROUND'
                  ? 'nearest'
                  : 'down',
          );
        }
        if (['POWER', 'MOD', 'QUOTIENT', 'CEILING', 'FLOOR', 'LOG'].includes(name)) {
          arity(name === 'LOG' ? 1 : 2, 2);
          const a = numeric(args[0]),
            b = args[1] === undefined ? arithmetic.parse('10') : numeric(args[1]);
          if (name === 'POWER') return arithmetic.power(a, b);
          if (name === 'LOG')
            return arithmetic.divide(arithmetic.unary('LN', a), arithmetic.unary('LN', b));
          const quotient = arithmetic.round(
            arithmetic.divide(a, b),
            0,
            name === 'MOD' || name === 'FLOOR' ? 'floor' : name === 'CEILING' ? 'ceil' : 'down',
          );
          return name === 'MOD'
            ? arithmetic.subtract(a, arithmetic.multiply(b, quotient))
            : name === 'QUOTIENT'
              ? quotient
              : arithmetic.multiply(b, quotient);
        }
        if (
          [
            'ABS',
            'SIGN',
            'SQRT',
            'SIN',
            'COS',
            'TAN',
            'ASIN',
            'ACOS',
            'ATAN',
            'EXP',
            'LN',
            'LOG10',
          ].includes(name)
        ) {
          arity(1);
          const value = numeric(args[0]);
          if (name === 'ABS')
            return arithmetic.compare(value, zero) < 0 ? arithmetic.negate(value) : value;
          if (name === 'SIGN') return arithmetic.parse(String(arithmetic.compare(value, zero)));
          return name === 'SQRT' ? arithmetic.sqrt(value) : arithmetic.unary(name, value);
        }
        const external = host.callFunction?.(name, flat.map(checked), {
          sheetId: workbook.worksheets[sheetIndex].id,
          x,
          y,
        });
        if (external === undefined) throw new CalculationError('#NAME?');
        return external;
      }
    }
  }
  return {
    schemaVersion: 1,
    worksheets: workbook.worksheets.map((sheet, index) => ({
      id: sheet.id,
      values: sheet.data.map((row, y) =>
        row.map((_cell, x) => {
          try {
            return read(index, x, y, 0);
          } catch (error) {
            return errorValue(error instanceof CalculationError ? error.code : '#ERROR!');
          }
        }),
      ),
    })),
  };
}
