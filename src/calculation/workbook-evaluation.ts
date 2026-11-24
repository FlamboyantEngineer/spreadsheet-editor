import { worksheetCalculationExtent } from '../model/extent.js';
import { type FormulaNode, parseFormula } from '../model/formula-parser.js';
import { requireDefined } from '../model/require-defined.js';
import {
  formatSpreadsheetAddress,
  MAX_COLUMN_INDEX,
  MAX_ROW_INDEX,
} from '../model/spreadsheet-addressing.js';
import { type CellInput, type Workbook, worksheetNameKey } from '../model/workbook.js';
import { calculationCellKey } from './calculation-cell-key.js';
import {
  authoredValue,
  type CellContext,
  type CellPosition,
  errorValueOf,
  formulaOperandKey,
  isFormula,
  keyFor,
  memoFor,
  withUnitFormat,
} from './cell-evaluation.js';
import { type Coercion, createCoercion } from './coercion.js';
import type {
  CalculationContext,
  CalculationDependencies,
  CalculationOptions,
  FormulaCellContext,
  FormulaHost,
} from './evaluator.js';
import { evaluateNode } from './formula-evaluation.js';
import type { Arithmetic, Value } from './function-call.js';
import { createQuantityArithmetic } from './units.js';
import { CalculationError, type CellValue, errorValue } from './values.js';

/** A formula being evaluated: the cell that holds it, and how deep the evaluation has gone. */
export interface Frame extends CellPosition {
  depth: number;
}
/** How many cells one formula has addressed; a formula may address 10,000 in all. */
interface ReadBudget {
  used: number;
  charged: WeakSet<FormulaNode>;
  exhausted: boolean;
}
/** How far a formula has read a range, so a suspended read resumes where it stopped. */
interface RangeState {
  rows: CellValue[][];
  index: number;
}

/**
 * Thrown when a formula reads a formula cell that is not calculated yet; `ensureCell` calculates
 * that cell and resumes. Not an `Error`, so throwing it captures no stack.
 */
class NeedCell {
  constructor(readonly cell: CellPosition) {}
}

const outOfGrid = ({ x, y }: CellPosition) =>
  !Number.isSafeInteger(x) ||
  !Number.isSafeInteger(y) ||
  x < 0 ||
  y < 0 ||
  x > MAX_COLUMN_INDEX ||
  y > MAX_ROW_INDEX;

function calculationLimits(options: CalculationOptions) {
  const maxOperations = options.maxOperations ?? 2_000_000,
    maxDepth = options.maxDepth ?? 128;
  const valid =
    Number.isInteger(maxOperations) &&
    maxOperations >= 1 &&
    maxOperations <= 10_000_000 &&
    Number.isInteger(maxDepth) &&
    maxDepth >= 1 &&
    maxDepth <= 512;
  if (!valid) throw new RangeError('Invalid calculation limits.');
  return { maxOperations, maxDepth };
}

/** The state of one `calculateWorkbook` pass: values, dependencies, memos and budgets. */
export class WorkbookEvaluation implements CellContext {
  readonly math: Arithmetic;
  readonly values: Coercion;
  readonly maxOperations: number;
  readonly maxDepth: number;
  readonly operands: CalculationContext['operands'] = {};
  readonly readBudgets = new Map<string, ReadBudget>();
  readonly rangeMemos = new Map<string, WeakMap<FormulaNode, RangeState>>();
  operations = 0;
  /** One captured clock value per pass for TODAY and NOW. */
  clock: Date | undefined;
  private readonly cellDependencies = new Map<string, Set<string>>();
  private readonly extentDependencies = new Map<string, Set<string>>();
  private readonly cache = new Map<string, CellValue>();
  private readonly visiting = new Set<string>();
  private readonly parsed = new Map<string, FormulaNode>();
  private readonly cellMemos = new Map<string, WeakMap<FormulaNode, Value>>();
  private readonly extents = new Map<number, { columns: number; rows: number }>();
  private readonly sheetIndexes: Map<string, number>;

  constructor(
    readonly workbook: Workbook,
    readonly options: CalculationOptions,
    readonly host: FormulaHost,
  ) {
    this.math = createQuantityArithmetic(options);
    this.values = createCoercion(this.math);
    this.sheetIndexes = new Map(
      workbook.worksheets.map((sheet, index) => [worksheetNameKey(sheet.name), index]),
    );
    ({ maxOperations: this.maxOperations, maxDepth: this.maxDepth } = calculationLimits(options));
  }

  tick(depth: number): void {
    if (++this.operations > this.maxOperations || depth > this.maxDepth)
      throw new CalculationError('#LIMIT!');
  }

  cellInput({ sheet, x, y }: CellPosition): CellInput {
    return this.workbook.worksheets[sheet]?.data[y]?.[x] ?? null;
  }

  resolveSheet(name: string | null, current: number): number {
    if (name === null) return current;
    const target = this.sheetIndexes.get(worksheetNameKey(name));
    if (target === undefined) throw new CalculationError('#REF!');
    return target;
  }

  formulaContext(cell: CellPosition): FormulaCellContext {
    return {
      sheetId: this.workbook.worksheets[cell.sheet].id,
      x: cell.x,
      y: cell.y,
      formula: String(this.cellInput(cell) ?? ''),
    };
  }

  rememberRead(reader: CellPosition, target: CellPosition): void {
    const key = keyFor(reader);
    let reads = this.cellDependencies.get(key);
    if (!reads) {
      reads = new Set();
      this.cellDependencies.set(key, reads);
    }
    reads.add(keyFor(target));
  }

  rememberExtent(reader: CellPosition, sheetId: string): void {
    const key = keyFor(reader),
      sheets = this.extentDependencies.get(key) ?? new Set<string>();
    sheets.add(sheetId);
    this.extentDependencies.set(key, sheets);
  }

  /** The workbook does not change during a calculation, so each sheet's extent is read once. */
  calculationExtent(sheet: number) {
    let extent = this.extents.get(sheet);
    if (!extent) {
      extent = worksheetCalculationExtent(this.workbook.worksheets[sheet]);
      this.extents.set(sheet, extent);
    }
    return extent;
  }

  read(target: CellPosition, depth: number): CellValue {
    this.tick(depth);
    if (outOfGrid(target)) return errorValue('#REF!');
    const key = keyFor(target);
    // An empty cell's value is null, so presence is read with has().
    if (this.cache.has(key)) return this.cache.get(key) as CellValue;
    if (this.visiting.has(key)) return errorValue('#CYCLE!');
    // A value cell reads nothing, so it is calculated in place instead of suspending the reader.
    if (!isFormula(this.cellInput(target))) return this.evaluateCell(target, 0);
    // eslint-disable-next-line @typescript-eslint/only-throw-error -- a suspension signal, not a failure: an Error would capture a stack on every suspended read
    throw new NeedCell(target);
  }

  /** Calculates a cell, first calculating every formula cell it reads, without recursion. */
  ensureCell(cell: CellPosition): CellValue {
    const stack = [cell];
    this.visiting.add(keyFor(cell));
    while (stack.length) {
      const current = requireDefined(stack.at(-1)),
        key = keyFor(current);
      try {
        if (!this.cache.has(key)) this.evaluateCell(current, 0);
        this.visiting.delete(key);
        this.cellMemos.delete(key);
        this.rangeMemos.delete(key);
        this.readBudgets.delete(key);
        stack.pop();
      } catch (error) {
        if (!(error instanceof NeedCell)) throw error;
        this.visiting.add(keyFor(error.cell));
        stack.push(error.cell);
      }
    }
    return this.cache.get(keyFor(cell)) as CellValue;
  }

  evaluateCell(cell: CellPosition, depth: number): CellValue {
    this.tick(depth);
    const key = keyFor(cell);
    let result: CellValue;
    try {
      const input = this.cellInput(cell);
      const value = this.inputValue(cell, { input, depth });
      const literalText =
        typeof input === 'string' && input.startsWith("'") && typeof value === 'string';
      result = literalText ? value : withUnitFormat(this, cell, value);
    } catch (error) {
      if (error instanceof NeedCell) throw error;
      result = errorValueOf(error);
    }
    if (this.readBudgets.get(key)?.exhausted || this.operations > this.maxOperations)
      result = errorValue('#LIMIT!');
    this.cache.set(key, result);
    return result;
  }

  evaluate(node: FormulaNode, frame: Frame): Value {
    const memo = memoFor(this.cellMemos, keyFor(frame));
    if (memo.has(node)) return memo.get(node) as Value;
    const value = evaluateNode(this, node, frame);
    memo.set(node, value);
    return value;
  }

  rememberOperand(
    node: FormulaNode,
    cell: CellPosition,
    operand: { name: string; value: CellValue },
  ) {
    if (!node.span) return;
    const sheetId = this.workbook.worksheets[cell.sheet].id;
    this.operands[formulaOperandKey({ sheetId, x: cell.x, y: cell.y }, node.span)] = {
      formula: String(this.cellInput(cell) ?? '').trimStart(),
      name: operand.name,
      value: structuredClone(operand.value),
      ...(this.host.hostContext === undefined
        ? {}
        : { hostContext: structuredClone(this.host.hostContext) }),
    };
  }

  dependencies(): CalculationDependencies {
    const publicKey = (key: string) => {
      const [sheet, x, y] = key.split(':').map(Number);
      return calculationCellKey(this.workbook.worksheets[sheet].id, x, y);
    };
    const extents = [...this.extentDependencies];
    return {
      cells: Object.fromEntries(
        [...this.cellDependencies].map(([key, reads]) => [
          publicKey(key),
          [...reads].map(publicKey),
        ]),
      ),
      ...(extents.length
        ? {
            extents: Object.fromEntries(
              extents.map(([key, sheets]) => [publicKey(key), [...sheets]]),
            ),
          }
        : {}),
      scope: 'formulas',
    };
  }

  private inputValue(
    cell: CellPosition,
    { input, depth }: { input: CellInput; depth: number },
  ): CellValue {
    if (isFormula(input)) return this.formulaValue(cell, { input, depth });
    const canonical =
      this.workbook.worksheets[cell.sheet].inputValues?.[formatSpreadsheetAddress(cell.x, cell.y)];
    return authoredValue(this, input, canonical);
  }

  private formulaValue(cell: CellPosition, { input, depth }: { input: string; depth: number }) {
    const key = keyFor(cell);
    if (!this.cellDependencies.has(key)) this.cellDependencies.set(key, new Set());
    const formula = input.trimStart().slice(1);
    let tree = this.parsed.get(formula);
    if (!tree) {
      tree = parseFormula(formula);
      this.parsed.set(formula, tree);
    }
    return this.values.scalar(this.evaluate(tree, { ...cell, depth: depth + 1 }));
  }
}
