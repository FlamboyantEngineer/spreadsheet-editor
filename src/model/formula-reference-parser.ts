import {
  buildReferenceCandidate,
  type ParsedReferenceCandidate,
} from './formula-reference-candidates.js';
import { FORMULA_WORD_CHARACTER } from './spreadsheet-addressing.js';
export type FormulaReferenceTarget =
  | {
      kind: 'cell';
      startColumn: number;
      startRow: number;
      endColumn: number;
      endRow: number;
    }
  | {
      kind: 'range';
      startColumn: number;
      startRow: number;
      endColumn: number;
      endRow: number;
    }
  | {
      kind: 'column-range';
      startColumn: number;
      endColumn: number;
    }
  | {
      kind: 'row-range';
      startRow: number;
      endRow: number;
    };

export type FormulaReferenceHighlight = {
  id: string;
  originalText: string;
  normalizedKey: string;
  startIndex: number;
  endIndex: number;
  target: FormulaReferenceTarget;
  colorIndex: number;
};

export type FormulaReferenceTokenSegment = {
  color?: string;
  colorIndex?: number;
  referenceKey?: string;
  text: string;
};

const FORMULA_REFERENCE_BOUNDARY_PATTERN = FORMULA_WORD_CHARACTER;
const FORMULA_REFERENCE_TOKEN_STOP_PATTERN = /[\s,;()+\-*/^&=<>{}]/;

export const FORMULA_REFERENCE_HIGHLIGHT_COLORS = [
  '#2563eb',
  '#dc2626',
  '#059669',
  '#7c3aed',
  '#d97706',
  '#0891b2',
  '#be123c',
  '#4d7c0f',
] as const;

function skipQuotedText(source: string, index: number, quote: string): number {
  let cursor = index + 1;
  while (cursor < source.length) {
    if (source[cursor] === '\\' && cursor + 1 < source.length) {
      cursor += 2;
      continue;
    }
    if (source[cursor] === quote && source[cursor + 1] === quote) {
      cursor += 2;
      continue;
    }
    if (source[cursor] === quote) return cursor + 1;
    cursor += 1;
  }
  return source.length;
}

function skipBracketedText(source: string, index: number): number {
  let depth = 1;
  let cursor = index + 1;
  while (cursor < source.length) {
    const character = source[cursor];
    if (character === '"' || character === "'") {
      cursor = skipQuotedText(source, cursor, character);
      continue;
    }
    if (character === '[') depth += 1;
    if (character === ']') {
      depth -= 1;
      if (depth <= 0) return cursor + 1;
    }
    cursor += 1;
  }
  return source.length;
}

function hasReferenceBoundaryBefore(source: string, index: number): boolean {
  const previous = index > 0 ? source[index - 1] : '';
  return !previous || !FORMULA_REFERENCE_BOUNDARY_PATTERN.test(previous);
}

function hasReferenceBoundaryAfter(source: string, index: number): boolean {
  const next = source[index] || '';
  return !next || !FORMULA_REFERENCE_BOUNDARY_PATTERN.test(next);
}

function hasCrossSheetBangBeforeReference(source: string, index: number): boolean {
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    const character = source[cursor] || '';
    if (character === '!') return true;
    if (FORMULA_REFERENCE_TOKEN_STOP_PATTERN.test(character)) return false;
  }
  return false;
}

function shouldIgnoreCandidate(source: string, candidate: ParsedReferenceCandidate): boolean {
  if (!hasReferenceBoundaryBefore(source, candidate.startIndex)) return true;
  if (!hasReferenceBoundaryAfter(source, candidate.endIndex)) return true;
  if (hasCrossSheetBangBeforeReference(source, candidate.startIndex)) return true;
  const nextCharacter = source[candidate.endIndex] || '';
  if (
    nextCharacter === '!' ||
    nextCharacter === '(' ||
    nextCharacter === '[' ||
    nextCharacter === ']'
  )
    return true;
  return false;
}

export function parseFormulaReferenceHighlights(formulaText: string): FormulaReferenceHighlight[] {
  const source = String(formulaText ?? '');
  const colorByNormalizedKey = new Map<string, number>();
  const references: FormulaReferenceHighlight[] = [];
  let index = 0;

  while (index < source.length) {
    const character = source[index];
    if (character === '"' || character === "'") {
      index = skipQuotedText(source, index, character);
      continue;
    }
    if (character === '[') {
      index = skipBracketedText(source, index);
      continue;
    }

    const candidate = buildReferenceCandidate(source, index);
    if (!candidate) {
      index += 1;
      continue;
    }
    if (shouldIgnoreCandidate(source, candidate)) {
      index = Math.max(candidate.endIndex, index + 1);
      continue;
    }
    if (!colorByNormalizedKey.has(candidate.normalizedKey)) {
      colorByNormalizedKey.set(candidate.normalizedKey, colorByNormalizedKey.size);
    }
    const colorIndex = colorByNormalizedKey.get(candidate.normalizedKey) || 0;
    references.push({
      ...candidate,
      id: `${candidate.normalizedKey}:${candidate.startIndex}-${candidate.endIndex}`,
      colorIndex,
    });
    index = candidate.endIndex;
  }

  return references;
}

export function buildFormulaReferenceTokenSegments(
  formulaText: string,
  references: readonly FormulaReferenceHighlight[],
): FormulaReferenceTokenSegment[] {
  const segments: FormulaReferenceTokenSegment[] = [];
  let cursor = 0;
  references
    .slice()
    .sort((left, right) => left.startIndex - right.startIndex || left.endIndex - right.endIndex)
    .forEach((reference) => {
      if (reference.startIndex < cursor || reference.endIndex <= reference.startIndex) return;
      if (reference.startIndex > cursor) {
        segments.push({ text: formulaText.slice(cursor, reference.startIndex) });
      }
      segments.push({
        text: formulaText.slice(reference.startIndex, reference.endIndex),
        referenceKey: reference.normalizedKey,
        colorIndex: reference.colorIndex,
        color:
          FORMULA_REFERENCE_HIGHLIGHT_COLORS[
            reference.colorIndex % FORMULA_REFERENCE_HIGHLIGHT_COLORS.length
          ] || FORMULA_REFERENCE_HIGHLIGHT_COLORS[0],
      });
      cursor = reference.endIndex;
    });
  if (cursor < formulaText.length) segments.push({ text: formulaText.slice(cursor) });
  return segments;
}
