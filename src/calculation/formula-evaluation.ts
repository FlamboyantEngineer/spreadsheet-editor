import type { BinaryOperator, ComparisonOperator, FormulaNode } from '../model/formula-parser.js';
import { deeper, hostValue } from './builtin-call.js';
import { evaluateCall } from './call-evaluation.js';
import type { Arithmetic, Value } from './function-call.js';
import { evaluateReference } from './reference-evaluation.js';
import { makeQuantity } from './units.js';
import { CalculationError, type CellValue, type NumericValue } from './values.js';
import type { Frame, WorkbookEvaluation } from './workbook-evaluation.js';

type NodeOf<K extends FormulaNode['kind']> = Extract<FormulaNode, { kind: K }>;

function name(evaluation: WorkbookEvaluation, node: NodeOf<'name'>, frame: Frame): CellValue {
  const value = hostValue(evaluation.host.readName?.(node.name, evaluation.formulaContext(frame)));
  evaluation.rememberOperand(node, frame, { name: node.name, value });
  return value;
}

function unary(evaluation: WorkbookEvaluation, node: NodeOf<'unary'>, frame: Frame): Value {
  const { math, values } = evaluation;
  if (node.operator === '-' && node.operand.kind === 'quantity')
    return makeQuantity(
      math.parse('-' + node.operand.decimal),
      node.operand.unit,
      evaluation.options,
    );
  const value = values.numeric(evaluation.evaluate(node.operand, deeper(frame)));
  if (node.operator === '-') return math.negate(value);
  return node.operator === '%' ? math.divide(value, math.parse('100')) : value;
}

const COMPARISONS: Readonly<Record<ComparisonOperator, (order: number) => boolean>> = {
  '=': (order) => order === 0,
  '<>': (order) => order !== 0,
  '<': (order) => order < 0,
  '<=': (order) => order <= 0,
  '>': (order) => order > 0,
  '>=': (order) => order >= 0,
};
type ArithmeticOperator = Exclude<BinaryOperator, ComparisonOperator | '&'>;
const ARITHMETIC: Readonly<
  Record<ArithmeticOperator, (math: Arithmetic, a: NumericValue, b: NumericValue) => NumericValue>
> = {
  '+': (math, a, b) => math.add(a, b),
  '-': (math, a, b) => math.subtract(a, b),
  '*': (math, a, b) => math.multiply(a, b),
  '/': (math, a, b) => math.divide(a, b),
  '^': (math, a, b) => math.power(a, b),
  '%': (math, a, b) => math.subtract(a, math.multiply(b, math.round(math.divide(a, b), 0, 'down'))),
};
const isComparison = (operator: BinaryOperator): operator is ComparisonOperator =>
  operator in COMPARISONS;

function binary(evaluation: WorkbookEvaluation, node: NodeOf<'binary'>, frame: Frame): Value {
  const { values } = evaluation;
  const left = evaluation.evaluate(node.left, deeper(frame)),
    right = evaluation.evaluate(node.right, deeper(frame));
  const { operator } = node;
  if (operator === '&') return values.text(left) + values.text(right);
  if (isComparison(operator))
    return COMPARISONS[operator](
      values.compare(values.checked(values.scalar(left)), values.checked(values.scalar(right))),
    );
  return ARITHMETIC[operator](evaluation.math, values.numeric(left), values.numeric(right));
}

/** Evaluates one node of a formula; `WorkbookEvaluation.evaluate` memoizes it per cell. */
export function evaluateNode(
  evaluation: WorkbookEvaluation,
  node: FormulaNode,
  frame: Frame,
): Value {
  evaluation.tick(frame.depth);
  switch (node.kind) {
    case 'literal':
      return node.value;
    case 'error':
      throw new CalculationError(node.code);
    case 'decimal':
      return evaluation.math.parse(node.decimal);
    case 'quantity':
      return makeQuantity(evaluation.math.parse(node.decimal), node.unit, evaluation.options);
    case 'name':
      return name(evaluation, node, frame);
    case 'reference':
    case 'lines':
      return evaluateReference(evaluation, node, frame);
    case 'unary':
      return unary(evaluation, node, frame);
    case 'binary':
      return binary(evaluation, node, frame);
    case 'call':
      return evaluateCall(evaluation, node, frame);
  }
}
