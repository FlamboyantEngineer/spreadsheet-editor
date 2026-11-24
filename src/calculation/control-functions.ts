import { type LazyCall, type LazyFunction, requireArity } from './builtin-call.js';
import type { Value } from './function-call.js';
import { CalculationError, type FormulaErrorCode, isError } from './values.js';

function conditional(call: LazyCall): Value {
  requireArity(call, 2, 3);
  const { run, node } = call;
  if (call.evaluation.values.truth(run(0))) return run(1);
  return node.args[2] ? run(2) : false;
}

function ifError(call: LazyCall): Value {
  requireArity(call, 2);
  try {
    const result = call.run(0);
    call.evaluation.values.checked(call.evaluation.values.scalar(result));
    return result;
  } catch (error) {
    if (error instanceof CalculationError && (call.name === 'IFERROR' || error.code === '#N/A'))
      return call.run(1);
    throw error;
  }
}

/** ISERROR matches any error, ISNA only `#N/A` and ISERR any other. */
function errorMatches(name: string, code: FormulaErrorCode): boolean {
  if (name === 'ISERROR') return true;
  return name === 'ISNA' ? code === '#N/A' : code !== '#N/A';
}

function isErrorKind(call: LazyCall): Value {
  requireArity(call, 1);
  try {
    const result = call.evaluation.values.scalar(call.run(0));
    return isError(result) && errorMatches(call.name, result.code);
  } catch (error) {
    if (error instanceof CalculationError) return errorMatches(call.name, error.code);
    throw error;
  }
}

function switchCase(call: LazyCall): Value {
  requireArity(call, 3, 255);
  const { run, evaluation } = call;
  const count = call.node.args.length;
  const target = evaluation.values.scalar(run(0));
  for (let i = 1; i + 1 < count; i += 2)
    if (evaluation.values.compare(target, evaluation.values.scalar(run(i))) === 0)
      return run(i + 1);
  if (count % 2 === 0) return run(count - 1);
  throw new CalculationError('#N/A');
}

function ifs(call: LazyCall): Value {
  requireArity(call, 2, 254);
  const count = call.node.args.length;
  if (count % 2) throw new CalculationError('#VALUE!');
  for (let i = 0; i < count; i += 2)
    if (call.evaluation.values.truth(call.run(i))) return call.run(i + 1);
  throw new CalculationError('#N/A');
}

function chooseIndex(call: LazyCall): Value {
  requireArity(call, 2, 255);
  const index = call.evaluation.values.numeric(call.run(0));
  if (index.dimension || index.quantityKind) throw new CalculationError('#VALUE!');
  const choice = Number(BigInt(index.numerator) / BigInt(index.denominator));
  if (choice < 1 || choice >= call.node.args.length) throw new CalculationError('#VALUE!');
  return call.run(choice);
}

/** ROW and COLUMN read the position of a reference, never its value. */
function position(call: LazyCall): Value {
  requireArity(call, 0, 1);
  const row = call.name === 'ROW';
  const { math } = call.evaluation;
  if (!call.node.args.length) return math.parse(String((row ? call.frame.y : call.frame.x) + 1));
  const reference = call.node.args[0];
  if (reference.kind === 'error') throw new CalculationError(reference.code);
  if (reference.kind !== 'reference') throw new CalculationError('#VALUE!');
  call.evaluation.resolveSheet(reference.sheetName, call.frame.sheet);
  return math.parse(String((row ? reference.start.y : reference.start.x) + 1));
}

/** Functions that evaluate only the arguments they need. */
export const CONTROL_FUNCTIONS: Readonly<Record<string, LazyFunction>> = {
  IF: conditional,
  IFERROR: ifError,
  IFNA: ifError,
  ISERROR: isErrorKind,
  ISERR: isErrorKind,
  ISNA: isErrorKind,
  SWITCH: switchCase,
  IFS: ifs,
  CHOOSE: chooseIndex,
  ROW: position,
  COLUMN: position,
};
