import { parseSpreadsheetColumnName, UNQUOTED_SHEET_NAME } from './spreadsheet-addressing.js';

export interface FormulaSpan {
  start: number;
  end: number;
}
export type CellCoords = { x: number; y: number };
export type ComparisonOperator = '=' | '<>' | '<' | '<=' | '>' | '>=';
export type BinaryOperator = ComparisonOperator | '&' | '+' | '-' | '*' | '/' | '%' | '^';

export type FormulaNode = (
  | { kind: 'literal'; value: string | boolean | null }
  | { kind: 'decimal'; decimal: string }
  | { kind: 'quantity'; decimal: string; unit: string }
  | { kind: 'name'; name: string }
  | { kind: 'reference'; sheetName: string | null; start: CellCoords; end: CellCoords | null }
  | { kind: 'lines'; sheetName: string | null; axis: 'column' | 'row'; start: number; end: number }
  | { kind: 'unary'; operator: '+' | '-' | '%'; operand: FormulaNode }
  | { kind: 'binary'; operator: BinaryOperator; left: FormulaNode; right: FormulaNode }
  | { kind: 'call'; name: string; args: FormulaNode[] }
) & { span?: FormulaSpan };

export class FormulaSyntaxError extends Error {
  constructor(readonly token = '#ERROR') {
    super(token);
  }
}
function fail(token = '#ERROR'): never {
  throw new FormulaSyntaxError(token);
}
const COMPARISON_OPERATORS: readonly ComparisonOperator[] = ['<=', '>=', '<>', '=', '<', '>'];
const CELL_ADDRESS_PATTERN = /^\$?([A-Za-z]+)\$?([1-9]\d*)/;
// Whole columns (`B:D`) and whole rows (`2:4`); a name or a call may not continue them.
const COLUMN_SPAN_PATTERN = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![A-Za-z0-9_.(])/;
const ROW_SPAN_PATTERN = /^\$?([1-9]\d*):\$?([1-9]\d*)(?![A-Za-z0-9_.(])/;
const NUMBER_PATTERN = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;
const ERROR_LITERAL_PATTERN = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)/i;
const IDENTIFIER_PATTERN = /^[A-Za-z_][A-Za-z0-9_.]*/;
const UNQUOTED_SHEET_PATTERN = new RegExp(`^${UNQUOTED_SHEET_NAME}`, 'u');
function readStringLiteral(
  source: string,
  startIndex: number,
): { endIndex: number; value: string } | null {
  const quote = source[startIndex];
  if (quote !== '"' && quote !== "'") return null;
  let value = '';
  let index = startIndex + 1;
  while (index < source.length) {
    const character = source[index] as string;
    if (character === quote) {
      if (source[index + 1] === quote) {
        value += quote;
        index += 2;
        continue;
      }
      return { endIndex: index + 1, value };
    }
    value += character;
    index += 1;
  }
  return null;
}

export function parseFormula(source: string): FormulaNode {
  if (source.length > 200_000) throw new FormulaSyntaxError('#CALC!');
  let index = 0;
  const skipWhitespace = (): void => {
    while (/\s/.test(source[index] || '')) index += 1;
  };
  const peek = (text: string): boolean => {
    skipWhitespace();
    return source.startsWith(text, index);
  };
  const consume = (text: string): boolean => {
    if (!peek(text)) return false;
    index += text.length;
    return true;
  };
  const readAddress = (): CellCoords | null => {
    skipWhitespace();
    const match = CELL_ADDRESS_PATTERN.exec(source.slice(index));
    const x = match ? parseSpreadsheetColumnName(match[1]) : null;
    if (!match || x === null) return null;
    index += match[0].length;
    return { x, y: Number.parseInt(match[2] as string, 10) - 1 };
  };
  const readSpan = (sheetName: string | null): FormulaNode | null => {
    skipWhitespace();
    const rest = source.slice(index);
    const columns = COLUMN_SPAN_PATTERN.exec(rest);
    if (columns) {
      index += columns[0].length;
      const start = parseSpreadsheetColumnName(columns[1]);
      const end = parseSpreadsheetColumnName(columns[2]);
      return start === null || end === null
        ? fail('#REF!')
        : { kind: 'lines', sheetName, axis: 'column', start, end };
    }
    const rows = ROW_SPAN_PATTERN.exec(rest);
    if (!rows) return null;
    index += rows[0].length;
    return {
      kind: 'lines',
      sheetName,
      axis: 'row',
      start: Number(rows[1]) - 1,
      end: Number(rows[2]) - 1,
    };
  };
  const readReference = (sheetName: string | null): FormulaNode => {
    const span = readSpan(sheetName);
    if (span) return span;
    const start = readAddress();
    if (!start) return fail('#REF!');
    if (!consume(':')) return { kind: 'reference', sheetName, start, end: null };
    const end = readAddress();
    return end ? { kind: 'reference', sheetName, start, end } : fail('#REF!');
  };
  const readSheetQualifiedReference = (): FormulaNode | null => {
    skipWhitespace();
    if (source[index] === "'") {
      const quoted = readStringLiteral(source, index);
      if (!quoted || source[quoted.endIndex] !== '!') return null;
      index = quoted.endIndex + 1;
      return readReference(quoted.value);
    }
    const match = UNQUOTED_SHEET_PATTERN.exec(source.slice(index));
    if (!match || source[index + match[0].length] !== '!') return null;
    index += match[0].length + 1;
    return readReference(match[0]);
  };
  const parseArguments = (): FormulaNode[] => {
    const args: FormulaNode[] = [];
    if (consume(')')) return args;
    do {
      // An omitted argument is a blank, as in Excel.
      args.push(peek(',') || peek(')') ? { kind: 'literal', value: null } : parseComparison());
    } while (consume(','));
    return consume(')') ? args : fail();
  };
  const parsePrimary = (): FormulaNode => {
    skipWhitespace();
    const start = index;
    const sheetReference = readSheetQualifiedReference();
    if (sheetReference) return sheetReference;
    const span = readSpan(null);
    if (span) return span;
    const addressMatch = CELL_ADDRESS_PATTERN.exec(source.slice(index));
    // `LOG10(` or `A1B` look like addresses but are identifiers.
    if (addressMatch && !/[A-Za-z0-9_.(]/.test(source[index + addressMatch[0].length] || '')) {
      return readReference(null);
    }
    if (consume('(')) {
      const inner = parseComparison();
      return consume(')') ? inner : fail();
    }
    const literal = readStringLiteral(source, index);
    if (literal) {
      index = literal.endIndex;
      return { kind: 'literal', value: literal.value };
    }
    // An error value written into a formula, such as the `#REF!` a pasted copy leaves for a reference off the grid.
    const errorMatch = ERROR_LITERAL_PATTERN.exec(source.slice(index));
    if (errorMatch) {
      index += errorMatch[0].length;
      return { kind: 'literal', value: errorMatch[0].toUpperCase() };
    }
    const binding = /^\{\{([^{}]+)\}\}/.exec(source.slice(index));
    if (binding) {
      index += binding[0].length;
      return { kind: 'name', name: binding[1], span: { start, end: index } };
    }
    const numberMatch = NUMBER_PATTERN.exec(source.slice(index));
    if (numberMatch) {
      index += numberMatch[0].length;
      const unit =
        /^(?:\s*)((?:fl oz|[\p{L}µΩ°′″'"]+)(?:\^[+-]?\d+|[²³])?(?:(?:\s*[*/]\s*|\s+)[\p{L}µΩ°′″'"]+(?:\^[+-]?\d+|[²³])?)*)/u.exec(
          source.slice(index),
        );
      if (unit && !['TRUE', 'FALSE'].includes(unit[1].toUpperCase())) {
        index += unit[0].length;
        return { kind: 'quantity', decimal: numberMatch[0], unit: unit[1] };
      }
      return { kind: 'decimal', decimal: numberMatch[0] };
    }
    const identifierMatch = IDENTIFIER_PATTERN.exec(source.slice(index));
    if (!identifierMatch) return fail();
    index += identifierMatch[0].length;
    const name = identifierMatch[0].toUpperCase();
    if (consume('(')) {
      const args = parseArguments();
      return { kind: 'call', name, args, span: { start, end: index } };
    }
    if (name === 'TRUE' || name === 'FALSE') return { kind: 'literal', value: name === 'TRUE' };
    return { kind: 'name', name: identifierMatch[0], span: { start, end: index } };
  };
  const parseUnary = (): FormulaNode => {
    if (consume('+')) return { kind: 'unary', operator: '+', operand: parseUnary() };
    if (consume('-')) return { kind: 'unary', operator: '-', operand: parseUnary() };
    return parsePrimary();
  };
  // A `%` before an operand keeps its earlier meaning, the remainder operator; otherwise it is Excel's percent.
  const isPostfixPercent = (): boolean => {
    if (!peek('%')) return false;
    let cursor = index + 1;
    while (/\s/.test(source[cursor] || '')) cursor += 1;
    return !/[\w.$('"]/.test(source[cursor] || '');
  };
  const parsePercent = (): FormulaNode => {
    let operand = parseUnary();
    while (isPostfixPercent()) {
      index += 1;
      operand = { kind: 'unary', operator: '%', operand };
    }
    return operand;
  };
  const parsePower = (): FormulaNode => {
    let left = parsePercent();
    for (;;) {
      // `**` is the earlier right-associative spelling; Excel's `^` associates left.
      if (consume('**')) return { kind: 'binary', operator: '^', left, right: parsePower() };
      if (!consume('^')) return left;
      left = { kind: 'binary', operator: '^', left, right: parsePercent() };
    }
  };
  const parseMultiplicative = (): FormulaNode => {
    let left = parsePower();
    for (;;) {
      if (peek('**')) return left;
      const operator = (['*', '/', '%'] as const).find((candidate) => consume(candidate));
      if (!operator) return left;
      left = { kind: 'binary', operator, left, right: parsePower() };
    }
  };
  const parseAdditive = (): FormulaNode => {
    let left = parseMultiplicative();
    for (;;) {
      const operator = (['+', '-'] as const).find((candidate) => consume(candidate));
      if (!operator) return left;
      left = { kind: 'binary', operator, left, right: parseMultiplicative() };
    }
  };
  const parseConcatenation = (): FormulaNode => {
    let left = parseAdditive();
    while (consume('&')) left = { kind: 'binary', operator: '&', left, right: parseAdditive() };
    return left;
  };
  function parseComparison(): FormulaNode {
    let left = parseConcatenation();
    for (;;) {
      const operator = COMPARISON_OPERATORS.find((candidate) => peek(candidate));
      if (!operator) return left;
      index += operator.length;
      left = { kind: 'binary', operator, left, right: parseConcatenation() };
    }
  }
  const tree = parseComparison();
  skipWhitespace();
  return index === source.length ? tree : fail();
}
