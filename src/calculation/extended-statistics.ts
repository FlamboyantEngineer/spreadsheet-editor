import {
  arity,
  checked,
  fail,
  flagged,
  flat,
  type FunctionCall,
  type FunctionImpl,
  integer,
  number,
  oneOf,
  scalar,
  zeroOf,
} from './function-call.js';
import { type CellValue, isNumeric, type NumericValue } from './values.js';
import { meetsCriterion } from './wildcard.js';

/** COUNTIF, SUMIF, AVERAGEIF and their several-condition forms, MAXIFS and MINIFS. */
function conditional(call: FunctionCall) {
  const { name, args, math } = call;
  const single = !name.endsWith('IFS'),
    count = name.startsWith('COUNT');
  if (single) arity(call, 2, count ? 2 : 3);
  else arity(call, count ? 2 : 3, 255);
  const outputs = flat(single ? (args[2] ?? args[0]) : args[0]),
    pairs = single ? args.slice(0, 2) : args.slice(count ? 0 : 1);
  if (pairs.length % 2) fail();
  const tests: { values: CellValue[]; condition: CellValue }[] = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const values = flat(pairs[i]);
    if (values.length !== outputs.length) fail();
    tests.push({ values, condition: scalar(pairs[i + 1]) });
  }
  const selected = outputs.filter((_, index) =>
    tests.every(({ values, condition }) => meetsCriterion(call, values[index], condition)),
  );
  if (count) return math.parse(String(selected.length));
  selected.forEach(checked);
  const numbers = selected.filter(isNumeric);
  if (name.startsWith('MAX') || name.startsWith('MIN'))
    return numbers.reduce(
      (a, b) => (call.compare(a, b) > 0 === name.startsWith('MAX') ? a : b),
      numbers[0] ?? zeroOf(call),
    );
  return name.startsWith('AVERAGE')
    ? math.average(numbers)
    : numbers.reduce(math.add, zeroOf(call));
}

function sumProduct(call: FunctionCall) {
  arity(call, 1, 255);
  const { math } = call;
  const ranges = call.args.map(flat),
    size = ranges[0].length;
  if (ranges.some((range) => range.length !== size)) fail();
  let sum = zeroOf(call);
  for (let i = 0; i < size; i++)
    sum = math.add(
      sum,
      ranges.reduce((product, range) => {
        const value = checked(range[i]);
        return math.multiply(product, isNumeric(value) ? value : zeroOf(call));
      }, oneOf(call)),
    );
  return sum;
}

function sortedNumbers(call: FunctionCall, values: ReturnType<typeof flat>): NumericValue[] {
  const numbers = values.map(checked).filter(isNumeric).sort(call.compare);
  if (!numbers.length) fail('#NUM!');
  return numbers;
}

function median(call: FunctionCall) {
  arity(call, 1, 255);
  const values = sortedNumbers(call, call.args.flatMap(flat));
  const mid = Math.floor(values.length / 2);
  return values.length % 2 ? values[mid] : call.math.average([values[mid - 1], values[mid]]);
}

function kth(call: FunctionCall) {
  arity(call, 2);
  const values = sortedNumbers(call, flat(call.args[0]));
  const k = integer(call, call.args[1], [1, values.length]);
  return values[call.name === 'SMALL' ? k - 1 : values.length - k];
}

function rank(call: FunctionCall) {
  arity(call, 2, 3);
  const { args, math } = call;
  const values = sortedNumbers(call, flat(args[1]));
  const target = number(call, args[0]),
    ascending = flagged(call, args[2]);
  const positions = values
    .map((value, index) =>
      call.compare(value, target) === 0 ? (ascending ? index + 1 : values.length - index) : 0,
    )
    .filter(Boolean);
  if (!positions.length) fail('#N/A');
  return call.name === 'RANKAVG'
    ? math.divide(
        math.parse(String(positions.reduce((a, b) => a + b, 0))),
        math.parse(String(positions.length)),
      )
    : math.parse(String(Math.min(...positions)));
}

export const STATISTICS_FUNCTIONS: Readonly<Record<string, FunctionImpl>> = {
  COUNTIF: conditional,
  SUMIF: conditional,
  AVERAGEIF: conditional,
  COUNTIFS: conditional,
  SUMIFS: conditional,
  AVERAGEIFS: conditional,
  MAXIFS: conditional,
  MINIFS: conditional,
  SUMPRODUCT: sumProduct,
  MEDIAN: median,
  LARGE: kth,
  SMALL: kth,
  RANK: rank,
  RANKAVG: rank,
  RANKEQ: rank,
};
