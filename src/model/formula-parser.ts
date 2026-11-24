import type { FormulaErrorCode } from '../calculation/values.js';
import { type CellCoords, formulaReferenceAt, quotedTextEnd } from './formula-reference-token.js';

export interface FormulaSpan {
  start: number;
  end: number;
}
export type ComparisonOperator = '=' | '<>' | '<' | '<=' | '>' | '>=';
export type BinaryOperator = ComparisonOperator | '&' | '+' | '-' | '*' | '/' | '%' | '^';

export type FormulaNode = (
  | { kind: 'literal'; value: string | boolean | null }
  | { kind: 'error'; code: FormulaErrorCode }
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
  constructor(readonly token: '#ERROR!' | '#LIMIT!' = '#ERROR!') {
    super(token);
  }
}
function fail(): never {
  throw new FormulaSyntaxError();
}
const MAX_NESTING = 256;
const COMPARISON_OPERATORS: readonly ComparisonOperator[] = ['<=', '>=', '<>', '=', '<', '>'];
const NUMBER_PATTERN = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;
const ERROR_LITERAL_PATTERN =
  /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|LIMIT!|CYCLE!)/i;
const IDENTIFIER_PATTERN = /^[A-Za-z_][A-Za-z0-9_.]*/;
function readStringLiteral(
  source: string,
  startIndex: number,
): { endIndex: number; value: string } | null {
  const quote = source[startIndex];
  if (quote !== '"' && quote !== "'") return null;
  const endIndex = quotedTextEnd(source, startIndex);
  return endIndex === null
    ? null
    : {
        endIndex,
        value: source.slice(startIndex + 1, endIndex - 1).replaceAll(quote + quote, quote),
      };
}

const UNIT_ATOM = String.raw`(?:fl oz|[\p{L}µΩ°′″'"]+)(?:\^[+-]?\d+|[²³])?(?![\p{L}\p{N}_.$!]|\s*\()`;
const UNIT_SUFFIX_PATTERN = new RegExp(
  `^\\s*(${UNIT_ATOM}(?:(?:\\s*[*/]\\s*|\\s+)${UNIT_ATOM})*)`,
  'u',
);

/** A recursive-descent parser over one formula; `index` is the read position. */
class FormulaParser {
  private index = 0;
  private nesting = 0;
  constructor(private readonly source: string) {}

  parse(): FormulaNode {
    const tree = this.parseComparison();
    this.skipWhitespace();
    return this.index === this.source.length ? tree : fail();
  }

  private get rest() {
    return this.source.slice(this.index);
  }
  private skipWhitespace(): void {
    while (/\s/.test(this.source[this.index] || '')) this.index += 1;
  }
  private peek(text: string): boolean {
    this.skipWhitespace();
    return this.source.startsWith(text, this.index);
  }
  private consume(text: string): boolean {
    if (!this.peek(text)) return false;
    this.index += text.length;
    return true;
  }
  /** Runs a pattern at the read position and moves past its match. */
  private take(pattern: RegExp): RegExpExecArray | null {
    const match = pattern.exec(this.rest);
    if (match) this.index += match[0].length;
    return match;
  }
  private nested(parse: () => FormulaNode): FormulaNode {
    if (this.nesting === MAX_NESTING) throw new FormulaSyntaxError('#LIMIT!');
    this.nesting++;
    const node = parse();
    this.nesting--;
    return node;
  }
  private parseArguments(): FormulaNode[] {
    const args: FormulaNode[] = [];
    if (this.consume(')')) return args;
    do {
      // An omitted argument is a blank, as in Excel.
      args.push(
        this.peek(',') || this.peek(')')
          ? { kind: 'literal', value: null }
          : this.nested(() => this.parseComparison()),
      );
    } while (this.consume(','));
    return this.consume(')') ? args : fail();
  }
  /** A reference, a whole-line span, or a parenthesized expression. */
  private parseGrouping(): FormulaNode | null {
    const reference = formulaReferenceAt(this.source, this.index);
    if (reference) {
      this.index = reference.endIndex;
      return reference.coordinates
        ? { ...reference.coordinates, sheetName: reference.sheetName }
        : { kind: 'error', code: '#REF!' };
    }
    if (!this.consume('(')) return null;
    const inner = this.nested(() => this.parseComparison());
    return this.consume(')') ? inner : fail();
  }
  /** A string, an error value or a `{{binding}}` written into the formula. */
  private parseToken(start: number): FormulaNode | null {
    const literal = readStringLiteral(this.source, this.index);
    if (literal) {
      this.index = literal.endIndex;
      return { kind: 'literal', value: literal.value };
    }
    // An error value written into a formula, such as the `#REF!` a pasted copy leaves for a reference off the grid.
    const errorMatch = this.take(ERROR_LITERAL_PATTERN);
    if (errorMatch) {
      const code = errorMatch[0].toUpperCase();
      return { kind: 'error', code: code === '#NULL!' ? '#VALUE!' : (code as FormulaErrorCode) };
    }
    const binding = this.take(/^\{\{([^{}]+)\}\}/);
    return binding ? { kind: 'name', name: binding[1], span: { start, end: this.index } } : null;
  }
  /** A number, with a unit after it when the text that follows names one. */
  private parseNumber(): FormulaNode | null {
    const numberMatch = this.take(NUMBER_PATTERN);
    if (!numberMatch) return null;
    const unit = UNIT_SUFFIX_PATTERN.exec(this.rest);
    if (!unit || ['TRUE', 'FALSE'].includes(unit[1].toUpperCase()))
      return { kind: 'decimal', decimal: numberMatch[0] };
    this.index += unit[0].length;
    return { kind: 'quantity', decimal: numberMatch[0], unit: unit[1] };
  }
  /** A function call, a logical constant or a name. */
  private parseIdentifier(start: number): FormulaNode {
    const identifierMatch = this.take(IDENTIFIER_PATTERN);
    if (!identifierMatch) return fail();
    const name = identifierMatch[0].toUpperCase();
    if (this.consume('(')) {
      const args = this.parseArguments();
      return { kind: 'call', name, args, span: { start, end: this.index } };
    }
    if (name === 'TRUE' || name === 'FALSE') return { kind: 'literal', value: name === 'TRUE' };
    return { kind: 'name', name: identifierMatch[0], span: { start, end: this.index } };
  }
  private parsePrimary(): FormulaNode {
    this.skipWhitespace();
    const start = this.index;
    return (
      this.parseGrouping() ??
      this.parseToken(start) ??
      this.parseNumber() ??
      this.parseIdentifier(start)
    );
  }
  private parseUnary(): FormulaNode {
    if (this.consume('+'))
      return { kind: 'unary', operator: '+', operand: this.nested(() => this.parseUnary()) };
    if (this.consume('-'))
      return { kind: 'unary', operator: '-', operand: this.nested(() => this.parseUnary()) };
    return this.parsePrimary();
  }
  // A `%` before an operand keeps its earlier meaning, the remainder operator; otherwise it is Excel's percent.
  private isPostfixPercent(): boolean {
    if (!this.peek('%')) return false;
    let cursor = this.index + 1;
    while (/\s/.test(this.source[cursor] || '')) cursor += 1;
    return !/[\w.$('"]/.test(this.source[cursor] || '');
  }
  private parsePercent(): FormulaNode {
    let operand = this.parseUnary();
    while (this.isPostfixPercent()) {
      this.index += 1;
      operand = { kind: 'unary', operator: '%', operand };
    }
    return operand;
  }
  private parsePower(): FormulaNode {
    let left = this.parsePercent();
    for (;;) {
      // `**` is the earlier right-associative spelling; Excel's `^` associates left.
      if (this.consume('**'))
        return { kind: 'binary', operator: '^', left, right: this.nested(() => this.parsePower()) };
      if (!this.consume('^')) return left;
      left = { kind: 'binary', operator: '^', left, right: this.parsePercent() };
    }
  }
  private parseMultiplicative(): FormulaNode {
    let left = this.parsePower();
    for (;;) {
      if (this.peek('**')) return left;
      const operator = (['*', '/', '%'] as const).find((candidate) => this.consume(candidate));
      if (!operator) return left;
      left = { kind: 'binary', operator, left, right: this.parsePower() };
    }
  }
  private parseAdditive(): FormulaNode {
    let left = this.parseMultiplicative();
    for (;;) {
      const operator = (['+', '-'] as const).find((candidate) => this.consume(candidate));
      if (!operator) return left;
      left = { kind: 'binary', operator, left, right: this.parseMultiplicative() };
    }
  }
  private parseConcatenation(): FormulaNode {
    let left = this.parseAdditive();
    while (this.consume('&'))
      left = { kind: 'binary', operator: '&', left, right: this.parseAdditive() };
    return left;
  }
  private parseComparison(): FormulaNode {
    let left = this.parseConcatenation();
    for (;;) {
      const operator = COMPARISON_OPERATORS.find((candidate) => this.peek(candidate));
      if (!operator) return left;
      this.index += operator.length;
      left = { kind: 'binary', operator, left, right: this.parseConcatenation() };
    }
  }
}

export function parseFormula(source: string): FormulaNode {
  if (source.length > 200_000) throw new FormulaSyntaxError('#LIMIT!');
  return new FormulaParser(source).parse();
}
