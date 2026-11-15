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

const UNIT_SUFFIX_PATTERN =
  /^(?:\s*)((?:fl oz|[\p{L}µΩ°′″'"]+)(?:\^[+-]?\d+|[²³])?(?:(?:\s*[*/]\s*|\s+)[\p{L}µΩ°′″'"]+(?:\^[+-]?\d+|[²³])?)*)/u;

/** A recursive-descent parser over one formula; `index` is the read position. */
class FormulaParser {
  private index = 0;
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
  private readAddress(): CellCoords | null {
    this.skipWhitespace();
    const match = CELL_ADDRESS_PATTERN.exec(this.rest);
    const x = match ? parseSpreadsheetColumnName(match[1]) : null;
    if (!match || x === null) return null;
    this.index += match[0].length;
    return { x, y: Number.parseInt(match[2] as string, 10) - 1 };
  }
  private readSpan(sheetName: string | null): FormulaNode | null {
    this.skipWhitespace();
    const columns = this.take(COLUMN_SPAN_PATTERN);
    if (columns) {
      const start = parseSpreadsheetColumnName(columns[1]);
      const end = parseSpreadsheetColumnName(columns[2]);
      return start === null || end === null
        ? fail('#REF!')
        : { kind: 'lines', sheetName, axis: 'column', start, end };
    }
    const rows = this.take(ROW_SPAN_PATTERN);
    if (!rows) return null;
    return {
      kind: 'lines',
      sheetName,
      axis: 'row',
      start: Number(rows[1]) - 1,
      end: Number(rows[2]) - 1,
    };
  }
  private readReference(sheetName: string | null): FormulaNode {
    const span = this.readSpan(sheetName);
    if (span) return span;
    const start = this.readAddress();
    if (!start) return fail('#REF!');
    if (!this.consume(':')) return { kind: 'reference', sheetName, start, end: null };
    const end = this.readAddress();
    return end ? { kind: 'reference', sheetName, start, end } : fail('#REF!');
  }
  private readSheetQualifiedReference(): FormulaNode | null {
    this.skipWhitespace();
    const { source, index } = this;
    if (source[index] === "'") {
      const quoted = readStringLiteral(source, index);
      if (!quoted || source[quoted.endIndex] !== '!') return null;
      this.index = quoted.endIndex + 1;
      return this.readReference(quoted.value);
    }
    const match = UNQUOTED_SHEET_PATTERN.exec(this.rest);
    if (!match || source[index + match[0].length] !== '!') return null;
    this.index += match[0].length + 1;
    return this.readReference(match[0]);
  }
  private parseArguments(): FormulaNode[] {
    const args: FormulaNode[] = [];
    if (this.consume(')')) return args;
    do {
      // An omitted argument is a blank, as in Excel.
      args.push(
        this.peek(',') || this.peek(')')
          ? { kind: 'literal', value: null }
          : this.parseComparison(),
      );
    } while (this.consume(','));
    return this.consume(')') ? args : fail();
  }
  /** A reference, a whole-line span, or a parenthesized expression. */
  private parseGrouping(): FormulaNode | null {
    const sheetReference = this.readSheetQualifiedReference();
    if (sheetReference) return sheetReference;
    const span = this.readSpan(null);
    if (span) return span;
    const addressMatch = CELL_ADDRESS_PATTERN.exec(this.rest);
    // `LOG10(` or `A1B` look like addresses but are identifiers.
    if (
      addressMatch &&
      !/[A-Za-z0-9_.(]/.test(this.source[this.index + addressMatch[0].length] || '')
    )
      return this.readReference(null);
    if (!this.consume('(')) return null;
    const inner = this.parseComparison();
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
    if (errorMatch) return { kind: 'literal', value: errorMatch[0].toUpperCase() };
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
    if (this.consume('+')) return { kind: 'unary', operator: '+', operand: this.parseUnary() };
    if (this.consume('-')) return { kind: 'unary', operator: '-', operand: this.parseUnary() };
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
        return { kind: 'binary', operator: '^', left, right: this.parsePower() };
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
  if (source.length > 200_000) throw new FormulaSyntaxError('#CALC!');
  return new FormulaParser(source).parse();
}
