import { checked, fail, type FunctionCall } from './function-call.js';
import type { CellValue } from './values.js';

interface Token {
  character: string;
  wildcard: boolean;
}
interface MatchBuffers {
  next: Uint8Array;
  current: Uint8Array;
}

function tokens(pattern: string): Token[] {
  const result: Token[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const character = pattern[i];
    if (character === '~' && /[~*?]/.test(pattern[i + 1] ?? ''))
      result.push({ character: pattern[++i].toUpperCase(), wildcard: false });
    else
      result.push({
        character: character.toUpperCase(),
        wildcard: character === '*' || character === '?',
      });
  }
  return result;
}

/** One step of the match table: where `token` can start a match, given where the rest can. */
function step(token: Token, text: readonly string[], buffers: MatchBuffers): void {
  const { next, current } = buffers;
  const star = token.wildcard && token.character === '*';
  for (let i = text.length; i >= 0; i--) {
    if (star) current[i] = next[i] || (i < text.length ? current[i + 1] : 0);
    else
      current[i] =
        i < text.length && (token.wildcard || token.character === text[i]) ? next[i + 1] : 0;
  }
  buffers.next = current;
  buffers.current = next;
}

/**
 * Bounded dynamic programming for authored wildcards, never a backtracking regex. A full match
 * returns 0 or -1; otherwise the index of the first match in `text`, or -1.
 */
export function compileWildcard(pattern: string, full = true): (text: string) => number {
  const list = tokens(pattern);
  return (text) => {
    if (pattern.length * (text.length + 1) > 1000000) fail('#LIMIT!');
    const characters = Array.from({ length: text.length }, (_, index) => text[index].toUpperCase());
    const buffers = {
      next: new Uint8Array(text.length + 1),
      current: new Uint8Array(text.length + 1),
    };
    if (!full) buffers.next.fill(1);
    buffers.next[text.length] = 1;
    for (let index = list.length - 1; index >= 0; index--) step(list[index], characters, buffers);
    if (full) return buffers.next[0] ? 0 : -1;
    return buffers.next.findIndex(Boolean);
  };
}

export const hasWildcard = (text: string) => /[*?]|~[~*?]/.test(text);

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

/** Compiles a COUNTIF-style condition once before visiting its range. */
export function compileCriterion(
  call: FunctionCall,
  condition: CellValue,
): (value: CellValue) => boolean {
  checked(condition);
  if (typeof condition !== 'string') return (value) => call.compare(value, condition) === 0;
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(condition),
    operator = match?.[1] ?? '=',
    target = criterionTarget(call, match?.[2] ?? condition);
  if (typeof target === 'string' && hasWildcard(target) && ['=', '<>'].includes(operator)) {
    const match = compileWildcard(target);
    return (value) => {
      const text = checked(value);
      return (typeof text === 'string' && match(text) >= 0) === (operator === '=');
    };
  }
  const compare = COMPARISONS[operator];
  return (value) => compare(call.compare(value, target));
}
