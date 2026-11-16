import { type CallNode, deeper, hostValue, type LazyCall } from './builtin-call.js';
import { BUILTIN_FUNCTIONS } from './builtin-functions.js';
import { CONTROL_FUNCTIONS } from './control-functions.js';
import { EXTENDED_FUNCTION_TABLE, FUNCTION_ALIASES } from './extended-functions.js';
import type { Value } from './function-call.js';
import { CalculationError, errorValue } from './values.js';
import type { Frame, WorkbookEvaluation } from './workbook-evaluation.js';

const COUNTS = ['COUNT', 'COUNTA', 'COUNTBLANK'];

/** Evaluates an argument; the counting functions count an error argument instead of failing. */
function argument(call: LazyCall, index: number): Value {
  try {
    return call.run(index);
  } catch (error) {
    if (COUNTS.includes(call.name) && error instanceof CalculationError)
      return errorValue(error.code);
    throw error;
  }
}

/** A function the host supplies; its arguments may hold no error. */
function hostFunction(call: LazyCall, args: readonly Value[]): Value {
  const { evaluation, frame, name } = call;
  const values = evaluation.values.flattened(args).map(evaluation.values.checked);
  const value = hostValue(
    evaluation.host.callFunction?.(name, values, evaluation.formulaContext(frame)),
  );
  evaluation.rememberOperand(call.node, frame, { name, value });
  return value;
}

/**
 * Evaluates a function call: control functions evaluate their own arguments; the others take
 * evaluated arguments and come from the extended catalog, the built-in one, or the host.
 */
export function evaluateCall(evaluation: WorkbookEvaluation, node: CallNode, frame: Frame): Value {
  const authored = node.name.toUpperCase();
  const name = FUNCTION_ALIASES[authored] ?? authored;
  const call: LazyCall = {
    evaluation,
    node,
    frame,
    name,
    run: (index) => evaluation.evaluate(node.args[index], deeper(frame)),
  };
  const control = CONTROL_FUNCTIONS[name];
  if (control) return control(call);
  const args = node.args.map((_, index) => argument(call, index));
  const extended = EXTENDED_FUNCTION_TABLE[name];
  if (extended)
    return extended({ name, args, math: evaluation.math, compare: evaluation.values.compare });
  const builtin = BUILTIN_FUNCTIONS[name];
  if (builtin) return builtin({ ...call, args, flat: evaluation.values.flattened(args) });
  return hostFunction(call, args);
}
