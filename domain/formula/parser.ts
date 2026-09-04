import type { ASTNode, FormulaErrorType, LiteralNode } from './types.ts';
import { type Token, tokenize } from './tokenizer.ts';
import { makeBlank, makeBoolean, makeError, makeNumber, makeString } from './values.ts';

export class ParseError extends Error {
  errorType: FormulaErrorType;

  constructor(message: string, errorType: FormulaErrorType = '#ERROR!') {
    super(message);
    this.errorType = errorType;
  }
}

export class Parser {
  private pos = 0;
  private tokens: Token[];

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  parse(): ASTNode {
    if (this.tokens.length === 0) {
      return { kind: 'literal', value: makeBlank() };
    }
    const node = this.parseComparison();
    if (this.pos < this.tokens.length) {
      throw new ParseError(`多余的符号 "${this.peek().value}"`, '#ERROR!');
    }
    return node;
  }

  // Comparison: =, <>, <, <=, >, >=
  private parseComparison(): ASTNode {
    let left = this.parseConcat();
    while (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && ['=', '<>', '<', '<=', '>', '>='].includes(token.value)) {
        this.consume();
        const right = this.parseConcat();
        left = {
          kind: 'binary_op',
          op: token.value as any,
          left,
          right,
        };
      } else {
        break;
      }
    }
    return left;
  }

  // Concatenation: &
  private parseConcat(): ASTNode {
    let left = this.parseAddSub();
    while (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && token.value === '&') {
        this.consume();
        const right = this.parseAddSub();
        left = {
          kind: 'binary_op',
          op: '&',
          left,
          right,
        };
      } else {
        break;
      }
    }
    return left;
  }

  // Addition & Subtraction: +, -
  private parseAddSub(): ASTNode {
    let left = this.parseMulDiv();
    while (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && (token.value === '+' || token.value === '-')) {
        this.consume();
        const right = this.parseMulDiv();
        left = {
          kind: 'binary_op',
          op: token.value as any,
          left,
          right,
        };
      } else {
        break;
      }
    }
    return left;
  }

  // Multiplication & Division: *, /
  private parseMulDiv(): ASTNode {
    let left = this.parseExponent();
    while (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && (token.value === '*' || token.value === '/')) {
        this.consume();
        const right = this.parseExponent();
        left = {
          kind: 'binary_op',
          op: token.value as any,
          left,
          right,
        };
      } else {
        break;
      }
    }
    return left;
  }

  // Exponentiation: ^
  private parseExponent(): ASTNode {
    let left = this.parseUnaryPrefix();
    while (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && token.value === '^') {
        this.consume();
        const right = this.parseUnaryPrefix();
        left = {
          kind: 'binary_op',
          op: '^',
          left,
          right,
        };
      } else {
        break;
      }
    }
    return left;
  }

  // Unary Prefix: +, -
  private parseUnaryPrefix(): ASTNode {
    if (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && (token.value === '+' || token.value === '-')) {
        this.consume();
        const argument = this.parseUnaryPrefix();
        return {
          kind: 'unary_op',
          op: token.value as '+' | '-',
          argument,
        };
      }
    }
    return this.parsePostfix();
  }

  // Postfix: % (e.g. 50% -> 0.5)
  private parsePostfix(): ASTNode {
    let node = this.parsePrimary();
    while (this.pos < this.tokens.length) {
      const token = this.peek();
      if (token.type === 'OPERATOR' && token.value === '%') {
        this.consume();
        node = {
          kind: 'unary_op',
          op: '%',
          argument: node,
        };
      } else {
        break;
      }
    }
    return node;
  }

  // Primary: literals, identifiers (function calls), references, ( expr )
  private parsePrimary(): ASTNode {
    if (this.pos >= this.tokens.length) {
      throw new ParseError('公式意外结束，缺少表达式');
    }

    const token = this.peek();

    // Parentheses
    if (token.type === 'LPAREN') {
      this.consume();
      const node = this.parseComparison();
      if (this.pos >= this.tokens.length || this.peek().type !== 'RPAREN') {
        throw new ParseError('缺少闭合括号 ")"');
      }
      this.consume(); // eat ')'
      return node;
    }

    // Number literal
    if (token.type === 'NUMBER') {
      this.consume();
      return {
        kind: 'literal',
        value: makeNumber(Number(token.value)),
        raw: token.value,
      };
    }

    // String literal
    if (token.type === 'STRING') {
      this.consume();
      return {
        kind: 'literal',
        value: makeString(token.value),
        raw: token.value,
      };
    }

    // Boolean literal
    if (token.type === 'BOOLEAN') {
      this.consume();
      return {
        kind: 'literal',
        value: makeBoolean(token.value === 'TRUE'),
        raw: token.value,
      };
    }

    // Error literal
    if (token.type === 'ERROR') {
      this.consume();
      return {
        kind: 'literal',
        value: makeError(token.value as any, `公式包含错误字面量 ${token.value}`),
        raw: token.value,
      };
    }

    // Single Cell Reference
    if (token.type === 'CELL_REF' && token.cellRef) {
      this.consume();
      return {
        kind: 'cell_ref',
        sheetName: token.cellRef.sheetName,
        col: token.cellRef.col,
        row: token.cellRef.row,
        colAbsolute: token.cellRef.colAbsolute,
        rowAbsolute: token.cellRef.rowAbsolute,
        rawText: token.value,
      };
    }

    // Range Reference (including whole-column range)
    if (token.type === 'RANGE_REF' && token.rangeRef) {
      this.consume();
      return {
        kind: 'range_ref',
        sheetName: token.rangeRef.sheetName,
        startCol: token.rangeRef.startCol,
        startRow: token.rangeRef.startRow,
        endCol: token.rangeRef.endCol,
        endRow: token.rangeRef.endRow,
        startColAbs: token.rangeRef.startColAbs,
        startRowAbs: token.rangeRef.startRowAbs,
        endColAbs: token.rangeRef.endColAbs,
        endRowAbs: token.rangeRef.endRowAbs,
        isFullColumn: token.rangeRef.isFullColumn,
        rawText: token.value,
      };
    }

    // Function call or unquoted identifier
    if (token.type === 'IDENTIFIER') {
      const name = token.value.toUpperCase();
      this.consume();
      if (this.pos < this.tokens.length && this.peek().type === 'LPAREN') {
        this.consume(); // eat '('
        const args: ASTNode[] = [];
        if (this.pos < this.tokens.length && this.peek().type !== 'RPAREN') {
          while (true) {
            if (this.pos < this.tokens.length && this.peek().type === 'COMMA') {
              // Empty argument, e.g. func(, 1)
              args.push({ kind: 'literal', value: makeBlank() });
              this.consume();
              continue;
            }
            args.push(this.parseComparison());
            if (this.pos < this.tokens.length && this.peek().type === 'COMMA') {
              this.consume();
              // If comma is followed immediately by ')', it's a trailing empty argument, e.g. func(1,)
              if (this.pos < this.tokens.length && this.peek().type === 'RPAREN') {
                args.push({ kind: 'literal', value: makeBlank() });
                break;
              }
            } else {
              break;
            }
          }
        }
        if (this.pos >= this.tokens.length || this.peek().type !== 'RPAREN') {
          throw new ParseError(`函数 ${name} 缺少闭合括号 ")"`);
        }
        this.consume(); // eat ')'
        return {
          kind: 'function_call',
          name,
          args,
        };
      }
      throw new ParseError(`未知标识符 "${token.value}"`, '#NAME?');
    }

    throw new ParseError(`无法解析的字符 "${token.value}"`);
  }

  private peek(): Token {
    return this.tokens[this.pos];
  }

  private consume(): Token {
    const t = this.tokens[this.pos];
    this.pos++;
    return t;
  }
}

export function parseFormula(formula: string): ASTNode {
  try {
    const tokens = tokenize(formula);
    const parser = new Parser(tokens);
    return parser.parse();
  } catch (err: any) {
    if (err instanceof ParseError) {
      return {
        kind: 'literal',
        value: makeError(err.errorType, err.message),
      };
    }
    return {
      kind: 'literal',
      value: makeError('#ERROR!', err?.message || '公式语法错误'),
    };
  }
}
