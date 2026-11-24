import { formulaReferenceTokens, type ReferenceCoordinates } from './formula-reference-token.js';
import { formatSpreadsheetAddress, formatSpreadsheetColumnName } from './spreadsheet-addressing.js';

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

function highlightTarget(reference: ReferenceCoordinates): FormulaReferenceTarget {
  if (reference.kind === 'lines') {
    const start = Math.min(reference.start, reference.end);
    const end = Math.max(reference.start, reference.end);
    return reference.axis === 'column'
      ? { kind: 'column-range', startColumn: start, endColumn: end }
      : { kind: 'row-range', startRow: start, endRow: end };
  }
  const { start, end } = reference;
  return {
    kind: end ? 'range' : 'cell',
    startColumn: Math.min(start.x, end?.x ?? start.x),
    startRow: Math.min(start.y, end?.y ?? start.y),
    endColumn: Math.max(start.x, end?.x ?? start.x),
    endRow: Math.max(start.y, end?.y ?? start.y),
  };
}

function referenceKey(target: FormulaReferenceTarget): string {
  if (target.kind === 'column-range')
    return `column:${formatSpreadsheetColumnName(target.startColumn)}:${formatSpreadsheetColumnName(target.endColumn)}`;
  if (target.kind === 'row-range') return `row:${target.startRow + 1}:${target.endRow + 1}`;
  const start = formatSpreadsheetAddress(target.startColumn, target.startRow);
  return target.kind === 'cell'
    ? `cell:${start}`
    : `range:${start}:${formatSpreadsheetAddress(target.endColumn, target.endRow)}`;
}

export function parseFormulaReferenceHighlights(source: string): FormulaReferenceHighlight[] {
  const colorByKey = new Map<string, number>();
  const references: FormulaReferenceHighlight[] = [];
  for (const { startIndex, endIndex, coordinates, sheetName } of formulaReferenceTokens(source)) {
    if (sheetName !== null || !coordinates) continue;
    const target = highlightTarget(coordinates);
    const normalizedKey = referenceKey(target);
    const colorIndex = colorByKey.get(normalizedKey) ?? colorByKey.size;
    colorByKey.set(normalizedKey, colorIndex);
    references.push({
      id: `${normalizedKey}:${startIndex}-${endIndex}`,
      originalText: source.slice(startIndex, endIndex),
      normalizedKey,
      startIndex,
      endIndex,
      target,
      colorIndex,
    });
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
