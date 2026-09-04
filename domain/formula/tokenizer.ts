import { MAX_COLUMNS, MAX_ROWS } from '../workbook/types.ts';

export type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'CELL_REF'
  | 'RANGE_REF'
  | 'IDENTIFIER'
  | 'OPERATOR'
  | 'LPAREN'
  | 'RPAREN'
  | 'COMMA'
  | 'ERROR';

export interface TokenCellRef {
  sheetName?: string;
  col: number;
  row: number;
  colAbsolute: boolean;
  rowAbsolute: boolean;
}

export interface TokenRangeRef {
  sheetName?: string;
  startCol: number;
  startRow: number | null;
  endCol: number;
  endRow: number | null;
  startColAbs: boolean;
  startRowAbs: boolean;
  endColAbs: boolean;
  endRowAbs: boolean;
  isFullColumn: boolean;
}

export interface Token {
  type: TokenType;
  value: string;
  pos: number;
  cellRef?: TokenCellRef;
  rangeRef?: TokenRangeRef;
}

export function colNameToIndex(letters: string): number {
  return Array.from(letters.toUpperCase()).reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

export function indexToColName(index: number): string {
  let result = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
  }
  return result;
}

const CELL_PATTERN = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/;
const FULL_COL_PATTERN = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/;

export class Tokenizer {
  private input: string;
  private pos = 0;
  private length: number;

  constructor(input: string) {
    this.input = input;
    this.length = input.length;
  }

  tokenize(): Token[] {
    const tokens: Token[] = [];
    while (this.pos < this.length) {
      this.skipWhitespace();
      if (this.pos >= this.length) break;

      const char = this.input[this.pos];
      const startPos = this.pos;

      // String literal
      if (char === '"') {
        tokens.push(this.readString());
        continue;
      }

      // Numbers
      if (this.isDigit(char) || (char === '.' && this.pos + 1 < this.length && this.isDigit(this.input[this.pos + 1]))) {
        tokens.push(this.readNumber());
        continue;
      }

      // Delimiters
      if (char === '(') {
        this.pos++;
        tokens.push({ type: 'LPAREN', value: '(', pos: startPos });
        continue;
      }
      if (char === ')') {
        this.pos++;
        tokens.push({ type: 'RPAREN', value: ')', pos: startPos });
        continue;
      }
      if (char === ',') {
        this.pos++;
        tokens.push({ type: 'COMMA', value: ',', pos: startPos });
        continue;
      }

      // Multi-char operators: <=, >=, <>
      if (char === '<' || char === '>') {
        const next = this.pos + 1 < this.length ? this.input[this.pos + 1] : '';
        if (char === '<' && (next === '=' || next === '>')) {
          this.pos += 2;
          tokens.push({ type: 'OPERATOR', value: char + next, pos: startPos });
          continue;
        }
        if (char === '>' && next === '=') {
          this.pos += 2;
          tokens.push({ type: 'OPERATOR', value: '>=', pos: startPos });
          continue;
        }
        this.pos++;
        tokens.push({ type: 'OPERATOR', value: char, pos: startPos });
        continue;
      }

      // Single-char operators: +, -, *, /, ^, %, &, =
      if (['+', '-', '*', '/', '^', '%', '&', '='].includes(char)) {
        this.pos++;
        tokens.push({ type: 'OPERATOR', value: char, pos: startPos });
        continue;
      }

      // Excel Errors, e.g. #REF!, #DIV/0!, #VALUE!, #N/A, #NAME?, #NUM!, #NULL!
      if (char === '#') {
        const errToken = this.tryReadError();
        if (errToken) {
          tokens.push(errToken);
          continue;
        }
      }

      // References (quoted sheet name or unquoted sheet/cell/range/identifier)
      // Check quoted sheet reference: 'Sheet Name'!A1 or 'Sheet Name'!A1:B10 or 'Sheet Name'!A:A
      if (char === "'") {
        const refToken = this.readQuotedSheetRef();
        tokens.push(refToken);
        continue;
      }

      // Otherwise read identifier / reference
      const token = this.readIdentifierOrRef();
      tokens.push(token);
    }

    return tokens;
  }

  private skipWhitespace() {
    while (this.pos < this.length && /\s/.test(this.input[this.pos])) {
      this.pos++;
    }
  }

  private isDigit(char: string): boolean {
    return char >= '0' && char <= '9';
  }

  private readString(): Token {
    const startPos = this.pos;
    this.pos++; // Skip opening "
    let content = '';
    while (this.pos < this.length) {
      if (this.input[this.pos] === '"') {
        if (this.pos + 1 < this.length && this.input[this.pos + 1] === '"') {
          content += '"';
          this.pos += 2;
        } else {
          this.pos++; // Closing quote
          return { type: 'STRING', value: content, pos: startPos };
        }
      } else {
        content += this.input[this.pos];
        this.pos++;
      }
    }
    // Unterminated string
    return { type: 'STRING', value: content, pos: startPos };
  }

  private readNumber(): Token {
    const startPos = this.pos;
    let text = '';
    while (this.pos < this.length && (this.isDigit(this.input[this.pos]) || this.input[this.pos] === '.')) {
      text += this.input[this.pos];
      this.pos++;
    }
    if (this.pos < this.length && (this.input[this.pos] === 'e' || this.input[this.pos] === 'E')) {
      let expText = this.input[this.pos];
      let p = this.pos + 1;
      if (p < this.length && (this.input[p] === '+' || this.input[p] === '-')) {
        expText += this.input[p];
        p++;
      }
      if (p < this.length && this.isDigit(this.input[p])) {
        while (p < this.length && this.isDigit(this.input[p])) {
          expText += this.input[p];
          p++;
        }
        this.pos = p;
        text += expText;
      }
    }
    return { type: 'NUMBER', value: text, pos: startPos };
  }

  private tryReadError(): Token | null {
    const startPos = this.pos;
    const errors = ['#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#CIRCULAR!', '#ERROR!'];
    for (const err of errors) {
      if (this.input.substring(startPos, startPos + err.length).toUpperCase() === err) {
        this.pos += err.length;
        return { type: 'ERROR', value: err, pos: startPos };
      }
    }
    return null;
  }

  private readQuotedSheetRef(): Token {
    const startPos = this.pos;
    this.pos++; // Skip opening '
    let sheetName = '';
    while (this.pos < this.length) {
      if (this.input[this.pos] === "'") {
        if (this.pos + 1 < this.length && this.input[this.pos + 1] === "'") {
          sheetName += "'";
          this.pos += 2;
        } else {
          this.pos++; // Closing quote
          break;
        }
      } else {
        sheetName += this.input[this.pos];
        this.pos++;
      }
    }

    if (this.pos < this.length && this.input[this.pos] === '!') {
      this.pos++; // Skip '!'
      // Now parse the cell or range after '!'
      return this.parseAddressAfterSheet(sheetName, startPos);
    }

    // If no '!' followed, treat as identifier
    return { type: 'IDENTIFIER', value: sheetName, pos: startPos };
  }

  private readIdentifierOrRef(): Token {
    const startPos = this.pos;
    // Read identifier or possible reference with sheet prefix, e.g. Sheet1!A1, 前日数据!Z:Z, A1, A1:B10, A:A
    let text = '';
    while (this.pos < this.length) {
      const c = this.input[this.pos];
      if (
        c === '(' ||
        c === ')' ||
        c === ',' ||
        c === '+' ||
        c === '-' ||
        c === '*' ||
        c === '/' ||
        c === '^' ||
        c === '%' ||
        c === '&' ||
        c === '=' ||
        c === '<' ||
        c === '>' ||
        c === '"' ||
        c === "'" ||
        /\s/.test(c)
      ) {
        break;
      }
      text += c;
      this.pos++;
    }

    // Check if this text has an unquoted sheet reference: e.g. Sheet1!A1 or 终端销售明细!A:A
    if (text.includes('!')) {
      const bangIndex = text.indexOf('!');
      const sheetName = text.slice(0, bangIndex);
      const refPart = text.slice(bangIndex + 1);
      const parsed = this.parseReferencePart(refPart, sheetName, startPos, text);
      if (parsed) return parsed;
    }

    // Check if directly a cell ref or range ref
    const parsedRef = this.parseReferencePart(text, undefined, startPos, text);
    if (parsedRef) return parsedRef;

    // Check boolean TRUE / FALSE
    const upper = text.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return { type: 'BOOLEAN', value: upper, pos: startPos };
    }

    return { type: 'IDENTIFIER', value: text, pos: startPos };
  }

  private parseAddressAfterSheet(sheetName: string, startPos: number): Token {
    let text = '';
    while (this.pos < this.length) {
      const c = this.input[this.pos];
      if (
        c === '(' ||
        c === ')' ||
        c === ',' ||
        c === '+' ||
        c === '-' ||
        c === '*' ||
        c === '/' ||
        c === '^' ||
        c === '%' ||
        c === '&' ||
        c === '=' ||
        c === '<' ||
        c === '>' ||
        c === '"' ||
        /\s/.test(c)
      ) {
        break;
      }
      text += c;
      this.pos++;
    }

    const fullRaw = `'${sheetName.replace(/'/g, "''")}'!${text}`;
    const parsed = this.parseReferencePart(text, sheetName, startPos, fullRaw);
    if (parsed) return parsed;

    return { type: 'IDENTIFIER', value: fullRaw, pos: startPos };
  }

  private parseReferencePart(text: string, sheetName: string | undefined, pos: number, rawText: string): Token | null {
    // 1. Full Column Range: e.g. A:A, E:H, $A:$A, $A:A, A:$A
    if (text.includes(':')) {
      const parts = text.split(':');
      if (parts.length === 2) {
        const left = parts[0];
        const right = parts[1];

        // Check if both are column-only letters: e.g. A:B or $A:$B
        const colOnlyRegex = /^\$?([A-Za-z]{1,3})$/;
        const matchLeftCol = colOnlyRegex.exec(left);
        const matchRightCol = colOnlyRegex.exec(right);

        if (matchLeftCol && matchRightCol) {
          const startCol = colNameToIndex(matchLeftCol[1]);
          const endCol = colNameToIndex(matchRightCol[1]);
          if (startCol >= 0 && startCol < MAX_COLUMNS && endCol >= 0 && endCol < MAX_COLUMNS) {
            return {
              type: 'RANGE_REF',
              value: rawText,
              pos,
              rangeRef: {
                sheetName,
                startCol: Math.min(startCol, endCol),
                startRow: null,
                endCol: Math.max(startCol, endCol),
                endRow: null,
                startColAbs: left.startsWith('$'),
                startRowAbs: false,
                endColAbs: right.startsWith('$'),
                endRowAbs: false,
                isFullColumn: true,
              },
            };
          }
        }

        // Check standard cell-to-cell range: e.g. A1:B10, $A$1:$B$10
        const cellRegex = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/;
        const matchL = cellRegex.exec(left);
        const matchR = cellRegex.exec(right);
        if (matchL && matchR) {
          const lCol = colNameToIndex(matchL[1]);
          const lRow = Number(matchL[2]) - 1;
          const rCol = colNameToIndex(matchR[1]);
          const rRow = Number(matchR[2]) - 1;
          if (
            lCol >= 0 &&
            lCol < MAX_COLUMNS &&
            lRow >= 0 &&
            lRow < MAX_ROWS &&
            rCol >= 0 &&
            rCol < MAX_COLUMNS &&
            rRow >= 0 &&
            rRow < MAX_ROWS
          ) {
            const startCol = Math.min(lCol, rCol);
            const endCol = Math.max(lCol, rCol);
            const startRow = Math.min(lRow, rRow);
            const endRow = Math.max(lRow, rRow);
            const lRowAbs = left.includes('$', 1) || (left.startsWith('$') && left.lastIndexOf('$') > 0);
            const rRowAbs = right.includes('$', 1) || (right.startsWith('$') && right.lastIndexOf('$') > 0);

            return {
              type: 'RANGE_REF',
              value: rawText,
              pos,
              rangeRef: {
                sheetName,
                startCol,
                startRow,
                endCol,
                endRow,
                startColAbs: left.startsWith('$'),
                startRowAbs: lRowAbs,
                endColAbs: right.startsWith('$'),
                endRowAbs: rRowAbs,
                isFullColumn: false,
              },
            };
          }
        }
      }
    }

    // 2. Single Cell Reference: e.g. A1, $A$1, A$1, $A1
    const cellRegex = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/;
    const m = cellRegex.exec(text);
    if (m) {
      const col = colNameToIndex(m[2]);
      const row = Number(m[4]) - 1;
      if (col >= 0 && col < MAX_COLUMNS && row >= 0 && row < MAX_ROWS) {
        return {
          type: 'CELL_REF',
          value: rawText,
          pos,
          cellRef: {
            sheetName,
            col,
            row,
            colAbsolute: m[1] === '$',
            rowAbsolute: m[3] === '$',
          },
        };
      }
    }

    return null;
  }
}

export function tokenize(formula: string): Token[] {
  const content = formula.startsWith('=') ? formula.slice(1) : formula;
  return new Tokenizer(content).tokenize();
}
