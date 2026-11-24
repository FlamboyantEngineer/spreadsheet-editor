import { type CalculationResult, formulaOperandKey } from '../calculation/evaluator.js';
import { type CellValue, isError, isNumeric } from '../calculation/values.js';
import type { FormulaNode } from '../model/formula-parser.js';
import { capturedText } from './captured-text.js';
import type { ExcelHostFunction } from './contracts.js';
import { EXCEL_ERRORS, operandText } from './export-values.js';

/** What inlining one formula needs: where it is, the captured results and how to report a loss. */
interface InlineRequest {
  formula: string;
  sheetId: string;
  x: number;
  y: number;
  result: CalculationResult;
  functions: ReadonlyMap<string, ExcelHostFunction>;
  loss: (code: string, message: string) => void;
}

interface Replacement {
  start: number;
  end: number;
  text: string;
}

/** One pass over a formula: the text replacements found, and whether any part has no equivalent. */
interface Inlining extends InlineRequest {
  replacements: Replacement[];
  unavailable: boolean;
}

const ALIASES: Readonly<Record<string, string>> = {
  AVG: 'AVERAGE',
  CEILINGMATH: 'CEILING.MATH',
  CEILINGPRECISE: 'CEILING.PRECISE',
  FLOORMATH: 'FLOOR.MATH',
  FLOORPRECISE: 'FLOOR.PRECISE',
  RANKAVG: 'RANK.AVG',
  RANKEQ: 'RANK.EQ',
  STDEVS: 'STDEV.S',
  VARS: 'VAR.S',
};

function unavailable(inlining: Inlining, code: string, message: string): undefined {
  inlining.unavailable = true;
  inlining.loss(code, message);
  return undefined;
}

/** The loss for a part Excel has no equivalent for, if the node is one. */
function refusal(node: FormulaNode): [string, string] | undefined {
  if (node.kind === 'quantity' || (node.kind === 'call' && ['EXPM1', 'LOG1P'].includes(node.name)))
    return [
      'formula-stored-value',
      'Excel has no equivalent for this quantity or numerical extension.',
    ];
  if (node.kind === 'literal' && typeof node.value === 'string' && node.value.length > 255)
    return ['formula-string-limit', 'A formula string exceeds Excel’s 255-character limit.'];
  return undefined;
}

/** The operand captured for a host name or call at this formula's position. */
function capturedOperand(inlining: Inlining, node: FormulaNode): CellValue | undefined {
  const stored = node.span
    ? inlining.result.context?.operands[formulaOperandKey(inlining, node.span)]
    : undefined;
  if (
    !stored ||
    stored.formula !== '=' + inlining.formula ||
    !node.span ||
    (isNumeric(stored.value) && (stored.value.dimension || stored.value.quantityKind))
  )
    return unavailable(
      inlining,
      'captured-operand',
      'A matching representable captured operand is unavailable.',
    );
  return stored.value;
}

/** Captured text as the host function declares it, rejected when Excel cannot carry it. */
function representable(inlining: Inlining, node: FormulaNode, captured: CellValue) {
  let value = captured;
  if (typeof value === 'string') {
    const policy =
      node.kind === 'call' ? (inlining.functions.get(node.name)?.capturedText ?? 'plain') : 'plain';
    try {
      value = capturedText(value, policy);
    } catch (error) {
      return unavailable(
        inlining,
        'captured-text',
        error instanceof Error ? error.message : 'Captured text cannot be represented.',
      );
    }
  }
  if (
    (typeof value === 'string' && value.length > 255) ||
    (isError(value) && !EXCEL_ERRORS.includes(value.code))
  )
    return unavailable(
      inlining,
      'captured-operand',
      'Excel cannot represent the captured operand.',
    );
  return { value };
}

function inlineHost(inlining: Inlining, node: FormulaNode) {
  const captured = capturedOperand(inlining, node);
  if (captured === undefined || !node.span) return;
  const operand = representable(inlining, node, captured);
  if (operand) inlining.replacements.push({ ...node.span, text: operandText(operand.value) });
}

function visit(inlining: Inlining, node: FormulaNode) {
  const refused = refusal(node);
  if (refused) return void unavailable(inlining, ...refused);
  if (node.kind === 'name' || (node.kind === 'call' && inlining.functions.has(node.name)))
    return inlineHost(inlining, node);
  if (node.kind === 'call') {
    if (node.span && Object.hasOwn(ALIASES, node.name))
      inlining.replacements.push({
        start: node.span.start,
        end: node.span.start + node.name.length,
        text: ALIASES[node.name],
      });
    node.args.forEach((arg) => visit(inlining, arg));
  } else if (node.kind === 'binary') {
    visit(inlining, node.left);
    visit(inlining, node.right);
  } else if (node.kind === 'unary') visit(inlining, node.operand);
}

/**
 * The formula with captured host operands written in as literals and catalog aliases renamed, or
 * null when Excel cannot carry it, after reporting why.
 */
export function inlineHostOperands(node: FormulaNode, request: InlineRequest): string | null {
  const inlining: Inlining = { ...request, replacements: [], unavailable: false };
  visit(inlining, node);
  if (inlining.unavailable) return null;
  let output = request.formula;
  for (const replacement of inlining.replacements.sort((a, b) => b.start - a.start))
    output = output.slice(0, replacement.start) + replacement.text + output.slice(replacement.end);
  if (output.length > 8192) {
    request.loss('formula-length', 'The expanded formula exceeds Excel’s 8192-character limit.');
    return null;
  }
  return output;
}
