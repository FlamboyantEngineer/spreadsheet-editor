import { formatCellValue } from './format.js';
import { checked, fail, type FunctionCall } from './function-call.js';
import type { CellValue } from './values.js';

interface Token {
  character: string;
  wildcard: boolean;
}

function tokens(pattern: string): Token[] {
  const result: Token[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const character = pattern[i];
    if (character === '~' && /[~*?]/.test(pattern[i + 1] ?? ''))
      result.push({ character: pattern[++i], wildcard: false });
    else result.push({ character, wildcard: character === '*' || character === '?' });
  }
  return result;
}

/** One step of the match table: where `token` can start a match, given where the rest can. */
function step(token: Token, text: string, next: readonly boolean[]): boolean[] {
  const current = Array<boolean>(text.length + 1).fill(false);
  const star = token.wildcard && token.character === '*';
  for (let i = text.length; i >= 0; i--) {
    if (star) current[i] = next[i] || (i < text.length && current[i + 1]);
    else
      current[i] =
        i < text.length &&
        (token.wildcard || token.character.toUpperCase() === text[i].toUpperCase()) &&
        next[i + 1];
  }
  return current;
}

/**
 * Bounded dynamic programming for authored wildcards, never a backtracking regex. A full match
 * returns 0 or -1; otherwise the index of the first match in `text`, or -1.
 */
export function wildcard(pattern: string, text: string, full = true): number {
  if (pattern.length * (text.length + 1) > 1000000) fail('#LIMIT!');
  const list = tokens(pattern);
  let next = Array<boolean>(text.length + 1).fill(!full);
  next[text.length] = true;
  for (let index = list.length - 1; index >= 0; index--) next = step(list[index], text, next);
  if (full) return next[0] ? 0 : -1;
  return next.findIndex(Boolean);
}

const COMPARISONS: Readonly<Record<string, (order: number) => boolean>> = {
  '=': (order) => order === 0,
  '<>': (order) => order !== 0,
  '<': (order) => order < 0,
  '>': (order) => order > 0,
  '<=': (order) => order <= 0,
  '>=': (order) => order >= 0,
};

function criterionTarget(call: FunctionCall, operand: string): CellValue {
  try {
    return call.math.parse(operand);
  } catch {
    const upper = operand.toUpperCase();
    return upper === 'TRUE' || upper === 'FALSE' ? upper === 'TRUE' : operand;
  }
}

/** Whether `value` meets a COUNTIF-style condition such as `">=3"` or `"a*"`. */
export function meetsCriterion(
  call: FunctionCall,
  value: CellValue,
  condition: CellValue,
): boolean {
  checked(value);
  checked(condition);
  if (typeof condition !== 'string') return call.compare(value, condition) === 0;
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(condition),
    operator = match?.[1] ?? '=',
    target = criterionTarget(call, match?.[2] ?? condition);
  if (typeof target === 'string' && /[*?~]/.test(target) && ['=', '<>'].includes(operator))
    return wildcard(target, formatCellValue(value)) >= 0 === (operator === '=');
  return COMPARISONS[operator](call.compare(value, target));
}
