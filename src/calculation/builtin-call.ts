import type { FormulaNode } from '../model/formula-parser.js';
import type { Value } from './function-call.js';
import {
  assertNumericValue,
  CalculationError,
  type CellValue,
  isNumeric,
  normalizeNumericValue,
} from './values.js';
import type { Frame, WorkbookEvaluation } from './workbook-evaluation.js';

export type CallNode = Extract<FormulaNode, { kind: 'call' }>;

/** A call of a built-in function that evaluates its own arguments, such as IF. */
export interface LazyCall {
  readonly evaluation: WorkbookEvaluation;
  readonly node: CallNode;
  readonly frame: Frame;
  /** The function's canonical name. */
  readonly name: string;
  /** Evaluates one argument. */
  run(index: number): Value;
}
/** A call of a built-in function whose arguments are evaluated before it runs. */
export interface EagerCall extends LazyCall {
  readonly args: readonly Value[];
  /** Every argument's cells, in order. */
  readonly flat: readonly CellValue[];
}
export type LazyFunction = (call: LazyCall) => Value;
export type EagerFunction = (call: EagerCall) => Value;

/** Fails with `#VALUE!` unless the call has between `min` and `max` arguments. */
export function requireArity(call: LazyCall, min: number, max = min): void {
  const count = call.node.args.length;
  if (count < min || count > max) throw new CalculationError('#VALUE!');
}

/** The frame of an operand: the same cell, one level deeper. */
export const deeper = (frame: Frame): Frame => ({ ...frame, depth: frame.depth + 1 });

/** A value the host supplies, for a name or a host function, checked like authored input. */
export function hostValue(value: CellValue | undefined): CellValue {
  if (value === undefined) throw new CalculationError('#NAME?');
  if (!isNumeric(value)) return value;
  assertNumericValue(value);
  return normalizeNumericValue(value);
}
