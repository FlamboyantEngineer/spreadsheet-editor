import { worksheetCalculationExtent } from '../model/extent.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { parseFormula, type FormulaNode } from '../model/formula-parser.js';
import {
  assertWorkbook,
  worksheetNameKey,
  type CellInput,
  type Workbook,
  type WorkbookLimits,
} from '../model/workbook.js';
import { type NumericOptions } from './rational.js';
import {
  createQuantityArithmetic,
  makeQuantity,
  quantityInUnit,
  validateUnitDimension,
} from './units.js';
import {
  EXTENDED_FUNCTIONS,
  FUNCTION_ALIASES,
  evaluateExtendedFunction,
} from './extended-functions.js';
import { calendarParts, calendarSerial } from './calendar.js';
import { formatCellValue } from './format.js';
import {
  CalculationError,
  errorValue,
  isError,
  isNumeric,
  type CellValue,
  type ErrorValue,
  type NumericValue,
  type JsonValue,
  assertJsonValue,
  assertNumericValue,
  normalizeNumericValue,
  TEXT_COLLATOR,
} from './values.js';

export interface CalculationOptions extends NumericOptions, WorkbookLimits {
  timeZone?: string;
  maxOperations?: number;
  maxDepth?: number;
}
export interface CalculatedWorksheet {
  id: string;
  values: CellValue[][];
}
export interface ResolvedFormulaOperand {
  formula: string;
  name: string;
  value: CellValue;
  hostContext?: JsonValue;
}
export interface CalculationContext {
  operands: Record<string, ResolvedFormulaOperand>;
  hostContext?: JsonValue;
}
export const formulaOperandKey = (
  sheetId: string,
  x: number,
  y: number,
  start: number,
  end: number,
) => JSON.stringify([sheetId, x, y, start, end]);
export interface CalculationDependencies {
  cells: Record<string, string[]>;
  extents?: Record<string, string[]>;
  /**
   * `'formulas'`: `cells` has an entry for every formula cell and none for other cells, which read
   * nothing. Without it, a cell with no entry has unknown dependencies.
   */
  scope?: 'formulas';
}
export interface FormulaCellContext {
  sheetId: string;
  x: number;
  y: number;
  formula: string;
}
export interface CalculationResult {
  schemaVersion: 1;
  worksheets: CalculatedWorksheet[];
  context?: CalculationContext;
  dependencies?: CalculationDependencies;
}
export interface FormulaHost {
  /** One captured clock value per pass for TODAY and NOW. */
  now?(): Date;
  hostContext?: JsonValue;
  readName?(name: string, context?: FormulaCellContext): CellValue | undefined;
  /** Host functions are synchronous and share this calculation's execution budget. */
  callFunction?(
    name: string,
    args: readonly CellValue[],
    context: FormulaCellContext,
  ): CellValue | undefined;
}
type Value = CellValue | CellValue[][];

export const SUPPORTED_FUNCTIONS = Object.freeze([
  ...EXTENDED_FUNCTIONS,
  ...Object.keys(FUNCTION_ALIASES),
  'SWITCH',
  'ACOSH',
  'ASINH',
  'ATANH',
  'COSH',
  'SINH',
  'TANH',
  'EXPM1',
  'LOG1P',
  'IF',
  'IFERROR',
  'IFNA',
  'ISNA',
  'ISERR',
  'IFS',
  'CHOOSE',
  'DATE',
  'TIME',
  'TODAY',
  'NOW',
  'YEAR',
  'MONTH',
  'DAY',
  'HOUR',
  'MINUTE',
  'SECOND',
  'DAYS',
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
  const arithmetic = createQuantityArithmetic(options);
  const maxOperations = options.maxOperations ?? 2_000_000,
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
  let clock: Date | undefined;
  if (host.hostContext !== undefined) assertJsonValue(host.hostContext);
  if (
    options.wordSizeBits !== undefined &&
    (typeof options.wordSizeBits !== 'string' ||
      !/^[1-9]\d{0,5}$/.test(options.wordSizeBits) ||
      BigInt(options.wordSizeBits) > 100000n)
  )
    throw new RangeError('Invalid wordSizeBits context.');
  const operands: CalculationContext['operands'] = {};
  // Dependencies use the internal `sheet:x:y` key while calculating; the public
  // `[sheetId,x,y]` key is written once per entry in the result.
  const cellDependencies = new Map<string, Set<string>>(),
    extentDependencies = new Map<string, Set<string>>();
  const publicKey = (key: string) => {
    const [sheet, x, y] = key.split(':').map(Number);
    return JSON.stringify([workbook.worksheets[sheet].id, x, y]);
  };
  function rememberRead(
    sheet: number,
    x: number,
    y: number,
    target: number,
    tx: number,
    ty: number,
  ) {
    const key = keyFor(sheet, x, y);
    let reads = cellDependencies.get(key);
    if (!reads) {
      reads = new Set();
      cellDependencies.set(key, reads);
    }
    reads.add(keyFor(target, tx, ty));
  }
  const extents = new Map<number, { columns: number; rows: number }>();
  /** The workbook does not change during a calculation, so each sheet's extent is read once. */
  function calculationExtent(sheetIndex: number) {
    let extent = extents.get(sheetIndex);
    if (!extent) {
      extent = worksheetCalculationExtent(workbook.worksheets[sheetIndex]);
      extents.set(sheetIndex, extent);
    }
    return extent;
  }
  const isFormula = (input: CellInput): boolean =>
    typeof input === 'string' && input.trimStart().startsWith('=');
  const cache = new Map<string, CellValue>(),
    visiting = new Set<string>();
  const parsed = new Map<string, FormulaNode>();
  /**
   * Thrown when a formula reads a formula cell that is not calculated yet; `ensureCell` calculates
   * that cell and resumes. Not an `Error`, so throwing it captures no stack.
   */
  class NeedCell {
    constructor(
      readonly sheet: number,
      readonly x: number,
      readonly y: number,
    ) {}
  }
  const cellMemos = new Map<string, WeakMap<FormulaNode, Value>>();
  const rangeMemos = new Map<
    string,
    WeakMap<FormulaNode, { rows: CellValue[][]; index: number }>
  >();
  const readBudgets = new Map<
    string,
    { used: number; charged: WeakSet<FormulaNode>; exhausted: boolean }
  >();
  const keyFor = (sheet: number, x: number, y: number) => `${sheet}:${x}:${y}`;
  function read(sheetIndex: number, x: number, y: number, depth: number): CellValue {
    tick(depth);
    if (
      !Number.isSafeInteger(x) ||
      !Number.isSafeInteger(y) ||
      x < 0 ||
      y < 0 ||
      x > 16383 ||
      y > 1048575
    )
      return errorValue('#REF!');
    const key = keyFor(sheetIndex, x, y);
    if (cache.has(key)) return cache.get(key)!;
    if (visiting.has(key)) return errorValue('#CYCLE!');
    // A value cell reads nothing, so it is calculated in place instead of suspending the reader.
    if (!isFormula(workbook.worksheets[sheetIndex]?.data[y]?.[x] ?? null))
      return evaluateCell(sheetIndex, x, y, 0);
    throw new NeedCell(sheetIndex, x, y);
  }
  function ensureCell(sheetIndex: number, x: number, y: number): CellValue {
    const initial = keyFor(sheetIndex, x, y);
    const stack = [{ sheet: sheetIndex, x, y }];
    visiting.add(initial);
    while (stack.length) {
      const cell = stack.at(-1)!,
        key = keyFor(cell.sheet, cell.x, cell.y);
      if (cache.has(key)) {
        visiting.delete(key);
        stack.pop();
        continue;
      }
      try {
        evaluateCell(cell.sheet, cell.x, cell.y, 0);
        visiting.delete(key);
        stack.pop();
      } catch (error) {
        if (!(error instanceof NeedCell)) throw error;
        const childKey = keyFor(error.sheet, error.x, error.y);
        visiting.add(childKey);
        stack.push({ sheet: error.sheet, x: error.x, y: error.y });
      }
    }
    return cache.get(initial)!;
  }
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
  function compareValues(a: CellValue, b: CellValue): number {
    checked(a);
    checked(b);
    const rank = (value: CellValue) =>
      typeof value === 'boolean' ? 2 : typeof value === 'string' ? 1 : 0;
    const av =
      a === null ? (typeof b === 'string' ? '' : typeof b === 'boolean' ? false : zero) : a;
    const bv =
      b === null ? (typeof a === 'string' ? '' : typeof a === 'boolean' ? false : zero) : b;
    return rank(av) !== rank(bv)
      ? rank(av) < rank(bv)
        ? -1
        : 1
      : isNumeric(av) && isNumeric(bv)
        ? arithmetic.compare(av, bv)
        : TEXT_COLLATOR.compare(String(av).toUpperCase(), String(bv).toUpperCase());
  }
  function text(value: Value): string {
    return formatCellValue(checked(scalar(value)), { approximationMarker: false });
  }
  function evaluateCell(sheetIndex: number, x: number, y: number, depth: number): CellValue {
    tick(depth);
    const key = keyFor(sheetIndex, x, y);
    let result: CellValue;
    try {
      const input: CellInput = workbook.worksheets[sheetIndex].data[y]?.[x] ?? null;
      const address = formatSpreadsheetAddress(x, y);
      const canonical = workbook.worksheets[sheetIndex].inputValues?.[address];
      if (typeof input === 'string' && isFormula(input)) {
        if (!cellDependencies.has(key)) cellDependencies.set(key, new Set());
        const formula = input.trimStart().slice(1);
        let tree = parsed.get(formula);
        if (!tree) {
          tree = parseFormula(formula);
          parsed.set(formula, tree);
        }
        result = scalar(evaluate(tree, sheetIndex, x, y, depth + 1));
      } else if (canonical) {
        assertNumericValue(canonical, options.maxIntegerDigits);
        result = normalizeNumericValue(canonical);
      } else if (typeof input === 'number') result = arithmetic.parse(String(input));
      else if (typeof input === 'string' && input.startsWith("'")) result = input.slice(1);
      else if (
        typeof input === 'string' &&
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(input.trim())
      )
        result = arithmetic.parse(input);
      else result = input === '' ? null : input;
      const format = workbook.worksheets[sheetIndex].formats?.[formatSpreadsheetAddress(x, y)];
      if (format?.kind === 'unit' && format.unit) {
        const configuredUnit = (unit: string) =>
          ['C', 'F'].includes(unit) && format.expectedDimension?.toLowerCase() === 'temperature'
            ? '°' + unit
            : unit;
        const displayUnit = configuredUnit(format.unit);
        validateUnitDimension(displayUnit, format.expectedDimension, options);
        if (typeof result === 'string') {
          const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s+(.+)$/.exec(
            result.trim(),
          );
          if (match) {
            try {
              const candidate = makeQuantity(
                arithmetic.parse(match[1]),
                match[2],
                options,
                format.quantityKind,
              );
              quantityInUnit(candidate, displayUnit, options);
              result = candidate;
            } catch {
              /* Invalid explicit unit text remains text. */
            }
          }
        } else if (isNumeric(result)) {
          if (!result.dimension && !result.quantityKind && !canonical)
            result = makeQuantity(
              result,
              configuredUnit(format.inputUnit ?? format.unit),
              options,
              format.quantityKind,
            );
          if (
            format.quantityKind &&
            result.quantityKind &&
            result.quantityKind !== format.quantityKind
          )
            throw new CalculationError('#VALUE!');
          if (format.quantityKind && !result.quantityKind)
            result = { ...result, quantityKind: format.quantityKind };
          quantityInUnit(result, displayUnit, options);
        }
      }
    } catch (error) {
      if (error instanceof NeedCell) throw error;
      result = errorValue(error instanceof CalculationError ? error.code : '#ERROR!');
    }
    if (readBudgets.get(key)?.exhausted || operations > maxOperations)
      result = errorValue('#LIMIT!');
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
    const key = keyFor(sheetIndex, x, y);
    let memo = cellMemos.get(key);
    if (!memo) {
      memo = new WeakMap();
      cellMemos.set(key, memo);
    }
    if (memo.has(node)) return memo.get(node)!;
    const value = evaluateFresh(node, sheetIndex, x, y, depth);
    memo.set(node, value);
    return value;
  }
  function rememberOperand(
    node: FormulaNode,
    name: string,
    value: CellValue,
    sheetIndex: number,
    x: number,
    y: number,
  ) {
    if (!node.span) return;
    operands[
      formulaOperandKey(workbook.worksheets[sheetIndex].id, x, y, node.span.start, node.span.end)
    ] = {
      formula: String(workbook.worksheets[sheetIndex].data[y]?.[x] ?? '').trimStart(),
      name,
      value: structuredClone(value),
      ...(host.hostContext === undefined ? {} : { hostContext: structuredClone(host.hostContext) }),
    };
  }
  function evaluateFresh(
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
        return makeQuantity(arithmetic.parse(node.decimal), node.unit, options);
      case 'name': {
        let value = host.readName?.(node.name, {
          sheetId: workbook.worksheets[sheetIndex].id,
          x,
          y,
          formula: String(workbook.worksheets[sheetIndex].data[y]?.[x] ?? ''),
        });
        if (isNumeric(value ?? null)) {
          assertNumericValue(value);
          value = normalizeNumericValue(value);
        }
        if (value === undefined) throw new CalculationError('#NAME?');
        rememberOperand(node, node.name, value, sheetIndex, x, y);
        return value;
      }
      case 'reference':
      case 'lines': {
        const target =
          node.sheetName === null
            ? sheetIndex
            : workbook.worksheets.findIndex(
                (sheet) => worksheetNameKey(sheet.name) === worksheetNameKey(node.sheetName!),
              );
        if (target < 0) throw new CalculationError('#REF!');
        const sheet = workbook.worksheets[target];
        const extent = node.kind === 'lines' ? calculationExtent(target) : null,
          width = extent?.columns ?? 0,
          height = extent?.rows ?? 0;
        if (node.kind === 'lines') {
          const key = keyFor(sheetIndex, x, y),
            extents = extentDependencies.get(key) ?? new Set<string>();
          extents.add(sheet.id);
          extentDependencies.set(key, extents);
        }
        if (node.kind === 'lines' && (!width || !height)) return [];
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
              ? { x: Math.max(node.start, node.end), y: height - 1 }
              : { x: width - 1, y: Math.max(node.start, node.end) };
        const cellKey = keyFor(sheetIndex, x, y);
        let budget = readBudgets.get(cellKey);
        if (!budget) {
          budget = { used: 0, charged: new WeakSet(), exhausted: false };
          readBudgets.set(cellKey, budget);
        }
        if (!budget.charged.has(node)) {
          budget.charged.add(node);
          budget.used += (Math.abs(end.x - start.x) + 1) * (Math.abs(end.y - start.y) + 1);
          if (budget.used > 10000) {
            budget.exhausted = true;
            throw new CalculationError('#LIMIT!');
          }
        }
        if (node.kind === 'reference' && !node.end) {
          rememberRead(sheetIndex, x, y, target, start.x, start.y);
          return checked(read(target, start.x, start.y, depth + 1));
        }
        const left = Math.min(start.x, end.x),
          top = Math.min(start.y, end.y),
          right = Math.max(start.x, end.x),
          bottom = Math.max(start.y, end.y);
        if ((right - left + 1) * (bottom - top + 1) > maxOperations - operations)
          throw new CalculationError('#LIMIT!');
        let states = rangeMemos.get(cellKey);
        if (!states) {
          states = new WeakMap();
          rangeMemos.set(cellKey, states);
        }
        let state = states.get(node);
        if (!state) {
          state = { rows: [], index: 0 };
          states.set(node, state);
        }
        const rangeWidth = right - left + 1,
          total = rangeWidth * (bottom - top + 1);
        while (state.index < total) {
          const row = Math.floor(state.index / rangeWidth),
            col = state.index % rangeWidth;
          rememberRead(sheetIndex, x, y, target, left + col, top + row);
          const value = read(target, left + col, top + row, depth + 1);
          (state.rows[row] ??= []).push(value);
          state.index++;
        }
        const rows = state.rows;
        return rows;
      }
      case 'unary': {
        if (node.operator === '-' && node.operand.kind === 'quantity')
          return makeQuantity(
            arithmetic.parse('-' + node.operand.decimal),
            node.operand.unit,
            options,
          );
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
          const comparison = compareValues(a, b);
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
        const authoredName = node.name.toUpperCase();
        const name = FUNCTION_ALIASES[authoredName] ?? authoredName;
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
        if (name === 'IFERROR' || name === 'IFNA') {
          arity(2);
          try {
            const result = run(node.args[0]);
            checked(scalar(result));
            return result;
          } catch (error) {
            if (error instanceof CalculationError && (name === 'IFERROR' || error.code === '#N/A'))
              return run(node.args[1]);
            throw error;
          }
        }
        if (['ISERROR', 'ISERR', 'ISNA'].includes(name)) {
          arity(1);
          try {
            const result = scalar(run(node.args[0]));
            return (
              isError(result) &&
              (name === 'ISERROR' ||
                (name === 'ISNA' ? result.code === '#N/A' : result.code !== '#N/A'))
            );
          } catch (error) {
            if (error instanceof CalculationError)
              return (
                name === 'ISERROR' ||
                (name === 'ISNA' ? error.code === '#N/A' : error.code !== '#N/A')
              );
            throw error;
          }
        }
        if (name === 'SWITCH') {
          arity(3, 255);
          const target = scalar(run(node.args[0]));
          for (let i = 1; i + 1 < count; i += 2)
            if (compareValues(target, scalar(run(node.args[i]))) === 0)
              return run(node.args[i + 1]);
          if (count % 2 === 0) return run(node.args[count - 1]);
          throw new CalculationError('#N/A');
        }
        if (name === 'IFS') {
          arity(2, 254);
          if (count % 2) throw new CalculationError('#VALUE!');
          for (let i = 0; i < count; i += 2)
            if (truth(run(node.args[i]))) return run(node.args[i + 1]);
          throw new CalculationError('#N/A');
        }
        if (name === 'CHOOSE') {
          arity(2, 255);
          const index = numeric(run(node.args[0]));
          if (index.dimension || index.quantityKind) throw new CalculationError('#VALUE!');
          const choice = Number(BigInt(index.numerator) / BigInt(index.denominator));
          if (choice < 1 || choice >= count) throw new CalculationError('#VALUE!');
          return run(node.args[choice]);
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
        const args = node.args.map((node) => {
            try {
              return run(node);
            } catch (error) {
              if (
                ['COUNT', 'COUNTA', 'COUNTBLANK'].includes(name) &&
                error instanceof CalculationError
              )
                return errorValue(error.code);
              throw error;
            }
          }),
          flat = flattened(args);
        const extended = evaluateExtendedFunction(name, args, arithmetic, compareValues);
        if (extended !== undefined) return extended;
        if (
          [
            'DATE',
            'TIME',
            'TODAY',
            'NOW',
            'YEAR',
            'MONTH',
            'DAY',
            'HOUR',
            'MINUTE',
            'SECOND',
            'DAYS',
          ].includes(name)
        ) {
          const integer = (value: Value) => {
            const n = numeric(value);
            if (n.dimension || n.quantityKind) throw new CalculationError('#VALUE!');
            const result = Number(BigInt(n.numerator) / BigInt(n.denominator));
            if (!Number.isSafeInteger(result)) throw new CalculationError('#NUM!');
            return result;
          };
          if (name === 'DATE') {
            arity(3);
            return calendarSerial(integer(args[0]), integer(args[1]), integer(args[2]));
          }
          if (name === 'TIME') {
            arity(3);
            const seconds = integer(args[0]) * 3600 + integer(args[1]) * 60 + integer(args[2]);
            if (seconds < 0 || !Number.isSafeInteger(seconds)) throw new CalculationError('#NUM!');
            return arithmetic.divide(
              arithmetic.parse(String(seconds % 86400)),
              arithmetic.parse('86400'),
            );
          }
          if (name === 'TODAY' || name === 'NOW') {
            arity(0);
            clock ??= host.now?.() ?? new Date();
            const fields = Object.fromEntries(
              new Intl.DateTimeFormat('en-US', {
                timeZone: options.timeZone ?? 'UTC',
                year: 'numeric',
                month: 'numeric',
                day: 'numeric',
                hour: 'numeric',
                minute: 'numeric',
                second: 'numeric',
                hourCycle: 'h23',
              })
                .formatToParts(clock)
                .map((part) => [part.type, part.value]),
            );
            const date = calendarSerial(
              Number(fields.year),
              Number(fields.month),
              Number(fields.day),
            );
            return name === 'TODAY'
              ? date
              : arithmetic.add(
                  date,
                  arithmetic.divide(
                    arithmetic.parse(
                      String(
                        Number(fields.hour) * 3600 +
                          Number(fields.minute) * 60 +
                          Number(fields.second),
                      ),
                    ),
                    arithmetic.parse('86400'),
                  ),
                );
          }
          if (name === 'DAYS') {
            arity(2);
            return arithmetic.subtract(
              arithmetic.parse(String(integer(args[0]))),
              arithmetic.parse(String(integer(args[1]))),
            );
          }
          arity(1);
          const parts = calendarParts(numeric(args[0]));
          return arithmetic.parse(
            String(
              parts[
                (
                  {
                    YEAR: 'year',
                    MONTH: 'month',
                    DAY: 'day',
                    HOUR: 'hour',
                    MINUTE: 'minute',
                    SECOND: 'second',
                  } as const
                )[name as 'YEAR']
              ],
            ),
          );
        }
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
          if (!['COUNT', 'COUNTA', 'COUNTBLANK'].includes(name)) flat.forEach(checked);
          const values: NumericValue[] = [];
          args.forEach((arg, index) => {
            const reference = ['reference', 'lines'].includes(node.args[index].kind);
            for (const value of flattened([arg])) {
              if (isNumeric(value)) values.push(value);
              else if (
                !['COUNTA', 'COUNTBLANK'].includes(name) &&
                !reference &&
                !Array.isArray(arg) &&
                value !== null &&
                !isError(value)
              ) {
                try {
                  values.push(numeric(value));
                } catch (error) {
                  if (name !== 'COUNT') throw error;
                }
              }
            }
          });
          if (name === 'COUNT') return arithmetic.parse(String(values.length));
          if (name === 'COUNTA')
            return arithmetic.parse(String(flat.filter((value) => value !== null).length));
          if (name === 'COUNTBLANK')
            return arithmetic.parse(
              String(flat.filter((value) => value === null || value === '').length),
            );
          if (name === 'AVERAGE') return arithmetic.average(values);
          const sum = name === 'SUM' ? values.reduce(arithmetic.add, zero) : zero;
          if (name === 'SUM') return sum;
          if (name === 'PRODUCT')
            return values.length ? values.reduce(arithmetic.multiply, one) : zero;
          if (name === 'SUMSQ')
            return values.reduce(
              (sum, value) => arithmetic.add(sum, arithmetic.multiply(value, value)),
              zero,
            );
          if (name === 'MIN' || name === 'MAX')
            return values.reduce(
              (a, b) => (arithmetic.compare(a, b) > 0 === (name === 'MIN') ? b : a),
              values[0] ?? zero,
            );
          const denominator = values.length - (name.endsWith('.S') ? 1 : 0);
          if (denominator <= 0) throw new CalculationError('#DIV/0!');
          const average = arithmetic.average(values);
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
          if (name === 'LOG') {
            const ratio = arithmetic.divide(arithmetic.unary('LN', a), arithmetic.unary('LN', b));
            const candidate = arithmetic.round(ratio, 0);
            // Verify integer powers rationally rather than trusting a rounded ln ratio.
            try {
              arithmetic.integer(candidate, -10000, 10000);
              if (arithmetic.compare(arithmetic.power(b, candidate), a) === 0) {
                const exact = { ...candidate, approximate: a.approximate || b.approximate };
                if (!exact.approximate) delete exact.precision;
                return exact;
              }
            } catch (error) {
              if (!(error instanceof CalculationError)) throw error;
            }
            return ratio;
          }
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
            'ACOSH',
            'ASINH',
            'ATANH',
            'COSH',
            'SINH',
            'TANH',
            'EXPM1',
            'LOG1P',
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
        let external = host.callFunction?.(name, flat.map(checked), {
          sheetId: workbook.worksheets[sheetIndex].id,
          x,
          y,
          formula: String(workbook.worksheets[sheetIndex].data[y]?.[x] ?? ''),
        });
        if (isNumeric(external ?? null)) {
          assertNumericValue(external);
          external = normalizeNumericValue(external);
        }
        if (external === undefined) throw new CalculationError('#NAME?');
        rememberOperand(node, name, external, sheetIndex, x, y);
        return external;
      }
    }
  }
  const result: CalculationResult = {
    schemaVersion: 1,
    worksheets: workbook.worksheets.map((sheet, index) => ({
      id: sheet.id,
      values: sheet.data.map((row, y) =>
        row.map((_cell, x) => {
          try {
            return ensureCell(index, x, y);
          } catch (error) {
            return errorValue(error instanceof CalculationError ? error.code : '#ERROR!');
          }
        }),
      ),
    })),
  };
  result.dependencies = {
    cells: Object.fromEntries(
      [...cellDependencies].map(([key, reads]) => [publicKey(key), [...reads].map(publicKey)]),
    ),
    ...(extentDependencies.size
      ? {
          extents: Object.fromEntries(
            [...extentDependencies].map(([key, sheets]) => [publicKey(key), [...sheets]]),
          ),
        }
      : {}),
    scope: 'formulas',
  };
  if (Object.keys(operands).length || host.hostContext !== undefined)
    result.context = {
      operands,
      ...(host.hostContext === undefined ? {} : { hostContext: structuredClone(host.hostContext) }),
    };
  return result;
}
