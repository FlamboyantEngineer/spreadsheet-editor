import type { FormulaNode } from '../model/formula-parser.js';
import type { CellCoords } from '../model/formula-reference-token.js';
import { type CellPosition, keyFor, memoFor } from './cell-evaluation.js';
import type { Value } from './function-call.js';
import { CalculationError } from './values.js';
import type { Frame, WorkbookEvaluation } from './workbook-evaluation.js';

type ReferenceNode = Extract<FormulaNode, { kind: 'reference' | 'lines' }>;
interface Bounds {
  start: CellCoords;
  end: CellCoords;
}
/** A formula may address at most this many cells in all its references. */
const READ_BUDGET = 10000;

/** The cells of whole rows or columns, bounded by the sheet's calculation extent. */
function lineBounds(
  node: Extract<FormulaNode, { kind: 'lines' }>,
  { columns, rows }: { columns: number; rows: number },
): Bounds {
  const first = Math.min(node.start, node.end),
    last = Math.max(node.start, node.end);
  return node.axis === 'column'
    ? { start: { x: first, y: 0 }, end: { x: last, y: rows - 1 } }
    : { start: { x: 0, y: first }, end: { x: columns - 1, y: last } };
}

/** Charges the cells a reference addresses to its formula's budget, once per reference. */
function charge(
  evaluation: WorkbookEvaluation,
  node: ReferenceNode,
  { reader, bounds }: { reader: string; bounds: Bounds },
) {
  let budget = evaluation.readBudgets.get(reader);
  if (!budget) {
    budget = { used: 0, charged: new WeakSet(), exhausted: false };
    evaluation.readBudgets.set(reader, budget);
  }
  if (budget.charged.has(node)) return;
  budget.charged.add(node);
  const { start, end } = bounds;
  budget.used += (Math.abs(end.x - start.x) + 1) * (Math.abs(end.y - start.y) + 1);
  if (budget.used > READ_BUDGET) {
    budget.exhausted = true;
    throw new CalculationError('#LIMIT!');
  }
}

/** Reads a range in row order; a suspended read resumes at the cell it stopped at. */
function readRange(
  evaluation: WorkbookEvaluation,
  node: ReferenceNode,
  { frame, sheet, bounds }: { frame: Frame; sheet: number; bounds: Bounds },
): Value {
  const left = Math.min(bounds.start.x, bounds.end.x),
    top = Math.min(bounds.start.y, bounds.end.y),
    width = Math.max(bounds.start.x, bounds.end.x) - left + 1,
    total = width * (Math.max(bounds.start.y, bounds.end.y) - top + 1);
  const states = memoFor(evaluation.rangeMemos, keyFor(frame));
  let state = states.get(node);
  if (total - (state?.index ?? 0) > evaluation.maxOperations - evaluation.operations)
    throw new CalculationError('#LIMIT!');
  if (!state) {
    state = { rows: [], index: 0 };
    states.set(node, state);
  }
  while (state.index < total) {
    const row = Math.floor(state.index / width);
    const target: CellPosition = { sheet, x: left + (state.index % width), y: top + row };
    evaluation.rememberRead(frame, target);
    (state.rows[row] ??= []).push(evaluation.read(target, frame.depth + 1));
    state.index++;
  }
  return state.rows;
}

/** Evaluates a cell, a range, or whole rows or columns. */
export function evaluateReference(
  evaluation: WorkbookEvaluation,
  node: ReferenceNode,
  frame: Frame,
): Value {
  const sheet = evaluation.resolveSheet(node.sheetName, frame.sheet);
  let bounds: Bounds;
  if (node.kind === 'lines') {
    evaluation.rememberExtent(frame, evaluation.workbook.worksheets[sheet].id);
    const extent = evaluation.calculationExtent(sheet);
    if (!extent.columns || !extent.rows) return [];
    bounds = lineBounds(node, extent);
  } else bounds = { start: node.start, end: node.end ?? node.start };
  charge(evaluation, node, { reader: keyFor(frame), bounds });
  if (node.kind === 'reference' && !node.end) {
    const target = { sheet, ...bounds.start };
    evaluation.rememberRead(frame, target);
    return evaluation.values.checked(evaluation.read(target, frame.depth + 1));
  }
  return readRange(evaluation, node, { frame, sheet, bounds });
}
