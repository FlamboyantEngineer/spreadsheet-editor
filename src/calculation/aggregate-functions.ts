import { type EagerCall, type EagerFunction, requireArity } from './builtin-call.js';
import { CalculationError, isError, isNumeric, type NumericValue } from './values.js';

const COUNTS = ['COUNT', 'COUNTA', 'COUNTBLANK'];

/**
 * The numbers an aggregate reads. Numbers in references always count; other values typed
 * directly as arguments are converted, except by COUNTA and COUNTBLANK, and COUNT skips those
 * that do not convert.
 */
function numbers(call: EagerCall): NumericValue[] {
  const { name, node, evaluation } = call;
  const values: NumericValue[] = [];
  const converts = !['COUNTA', 'COUNTBLANK'].includes(name);
  call.args.forEach((arg, index) => {
    const reference = ['reference', 'lines'].includes(node.args[index].kind);
    for (const value of evaluation.values.flattened([arg])) {
      if (isNumeric(value)) values.push(value);
      else if (converts && !reference && !Array.isArray(arg) && value !== null && !isError(value))
        try {
          values.push(evaluation.values.numeric(value));
        } catch (error) {
          if (name !== 'COUNT') throw error;
        }
    }
  });
  return values;
}

/** Requires one argument or more (SUM takes none), and no error unless the function counts. */
function prepare(call: EagerCall): NumericValue[] {
  requireArity(call, call.name === 'SUM' ? 0 : 1, 255);
  if (!COUNTS.includes(call.name)) call.flat.forEach(call.evaluation.values.checked);
  return numbers(call);
}

const count: EagerFunction = (call) => {
  const values = prepare(call);
  const { math } = call.evaluation;
  if (call.name === 'COUNT') return math.parse(String(values.length));
  if (call.name === 'COUNTA')
    return math.parse(String(call.flat.filter((value) => value !== null).length));
  return math.parse(String(call.flat.filter((value) => value === null || value === '').length));
};

const sum: EagerFunction = (call) => {
  const { math, values } = call.evaluation;
  return prepare(call).reduce(math.add, values.zero);
};
const product: EagerFunction = (call) => {
  const { math, values } = call.evaluation;
  const list = prepare(call);
  return list.length ? list.reduce(math.multiply, values.one) : values.zero;
};
const sumOfSquares: EagerFunction = (call) => {
  const { math, values } = call.evaluation;
  return prepare(call).reduce(
    (total, value) => math.add(total, math.multiply(value, value)),
    values.zero,
  );
};
const average: EagerFunction = (call) => call.evaluation.math.average(prepare(call));

const extreme: EagerFunction = (call) => {
  const { math, values } = call.evaluation;
  const list = prepare(call),
    min = call.name === 'MIN';
  return list.reduce((a, b) => (math.compare(a, b) > 0 === min ? b : a), list[0] ?? values.zero);
};

/** VAR.S, VAR.P, STDEV.S and STDEV.P; the `.S` forms divide by one less than the count. */
const spread: EagerFunction = (call) => {
  const { math, values } = call.evaluation;
  const list = prepare(call);
  const denominator = list.length - (call.name.endsWith('.S') ? 1 : 0);
  if (denominator <= 0) throw new CalculationError('#DIV/0!');
  const mean = math.average(list);
  const squared = list.reduce((total, value) => {
    const delta = math.subtract(value, mean);
    return math.add(total, math.multiply(delta, delta));
  }, values.zero);
  const variance = math.divide(squared, math.parse(String(denominator)));
  return call.name.startsWith('STDEV') ? math.sqrt(variance) : variance;
};

export const AGGREGATE_FUNCTIONS: Readonly<Record<string, EagerFunction>> = {
  SUM: sum,
  PRODUCT: product,
  AVERAGE: average,
  MIN: extreme,
  MAX: extreme,
  COUNT: count,
  COUNTA: count,
  COUNTBLANK: count,
  SUMSQ: sumOfSquares,
  'VAR.S': spread,
  'VAR.P': spread,
  'STDEV.S': spread,
  'STDEV.P': spread,
};
