import type { ASTNode, CalculationContext, FormulaValue } from './types.ts';
import { defaultRegistry } from './registry.ts';
import {
  formatFormulaValue,
  isError,
  makeBlank,
  makeBoolean,
  makeError,
  makeNumber,
  makeRange,
  makeString,
  toBoolean,
  toNumber,
  toText,
} from './values.ts';

export function evaluateNode(node: ASTNode, context: CalculationContext): FormulaValue {
  switch (node.kind) {
    case 'literal':
      return node.value;

    case 'cell_ref': {
      const resolved = context.resolveCell(node);
      if ('type' in resolved) {
        return makeError(resolved.type, resolved.message, resolved.details);
      }
      return context.evalCell(resolved.sheetId, resolved.row, resolved.col);
    }

    case 'range_ref': {
      const resolved = context.resolveRange(node);
      if ('type' in resolved) {
        return makeError(resolved.type, resolved.message, resolved.details);
      }
      return makeRange(resolved);
    }

    case 'unary_op': {
      const argVal = context.evalNode(node.argument);
      if (isError(argVal)) return argVal;

      if (node.op === '%') {
        const numRes = toNumber(argVal);
        if (!numRes.ok) return makeError(numRes.error.type, numRes.error.message);
        return makeNumber(numRes.value / 100);
      }

      if (node.op === '+') {
        const numRes = toNumber(argVal);
        if (!numRes.ok) return makeError(numRes.error.type, numRes.error.message);
        return makeNumber(numRes.value);
      }

      if (node.op === '-') {
        const numRes = toNumber(argVal);
        if (!numRes.ok) return makeError(numRes.error.type, numRes.error.message);
        return makeNumber(-numRes.value);
      }

      return makeError('#VALUE!', `未知的一元运算符: ${node.op}`);
    }

    case 'binary_op': {
      const leftVal = context.evalNode(node.left);
      if (isError(leftVal)) return leftVal;
      const rightVal = context.evalNode(node.right);
      if (isError(rightVal)) return rightVal;

      // String concatenation
      if (node.op === '&') {
        const lText = toText(leftVal);
        if (!lText.ok) return makeError(lText.error.type, lText.error.message);
        const rText = toText(rightVal);
        if (!rText.ok) return makeError(rText.error.type, rText.error.message);
        return makeString(lText.text + rText.text);
      }

      // Comparisons
      if (['=', '<>', '<', '<=', '>', '>='].includes(node.op)) {
        return evaluateComparison(node.op, leftVal, rightVal);
      }

      // Arithmetic
      const lNum = toNumber(leftVal);
      if (!lNum.ok) return makeError(lNum.error.type, lNum.error.message);
      const rNum = toNumber(rightVal);
      if (!rNum.ok) return makeError(rNum.error.type, rNum.error.message);

      const a = lNum.value;
      const b = rNum.value;

      switch (node.op) {
        case '+':
          return makeNumber(a + b);
        case '-':
          return makeNumber(a - b);
        case '*':
          return makeNumber(a * b);
        case '/':
          if (b === 0) return makeError('#DIV/0!', '除数为 0');
          return makeNumber(a / b);
        case '^':
          if (a < 0 && !Number.isInteger(b)) {
            return makeError('#NUM!', '负数的分数次方产生虚数');
          }
          const powRes = Math.pow(a, b);
          return makeNumber(powRes);
        default:
          return makeError('#VALUE!', `未知的二元运算符: ${node.op}`);
      }
    }

    case 'function_call': {
      const handler = defaultRegistry.get(node.name);
      if (!handler) {
        return makeError('#NAME?', `函数 ${node.name} 未定义`);
      }
      try {
        return handler(node.args, context);
      } catch (err: any) {
        return makeError('#ERROR!', `执行函数 ${node.name} 时出错: ${err?.message}`);
      }
    }
  }
}

function evaluateComparison(op: string, left: FormulaValue, right: FormulaValue): FormulaValue {
  // If either is range, extract top-left
  let l = left.type === 'range' ? left.range.getCellValue(left.range.startRow, left.range.startCol) : left;
  let r = right.type === 'range' ? right.range.getCellValue(right.range.startRow, right.range.startCol) : right;

  if (isError(l)) return l;
  if (isError(r)) return r;

  // If one is blank, match type with the other
  if (l.type === 'blank' && r.type !== 'blank') {
    if (r.type === 'number') l = makeNumber(0);
    else if (r.type === 'string') l = makeString('');
    else if (r.type === 'boolean') l = makeBoolean(false);
  } else if (r.type === 'blank' && l.type !== 'blank') {
    if (l.type === 'number') r = makeNumber(0);
    else if (l.type === 'string') r = makeString('');
    else if (l.type === 'boolean') r = makeBoolean(false);
  }

  let result = false;

  if (l.type === 'number' && r.type === 'number') {
    result = compareScalars(op, l.value, r.value);
  } else if (l.type === 'string' && r.type === 'string') {
    const cmp = l.value.localeCompare(r.value, undefined, { sensitivity: 'accent' });
    result = compareWithCompareResult(op, cmp);
  } else if (l.type === 'boolean' && r.type === 'boolean') {
    const lVal = l.value ? 1 : 0;
    const rVal = r.value ? 1 : 0;
    result = compareScalars(op, lVal, rVal);
  } else {
    // Differing types: Excel types hierarchy is number < string < boolean
    const typeOrder = { blank: 0, number: 1, string: 2, boolean: 3, error: 4, range: 5 };
    if (op === '=') {
      result = false;
    } else if (op === '<>') {
      result = true;
    } else {
      const lOrder = typeOrder[l.type];
      const rOrder = typeOrder[r.type];
      result = compareScalars(op, lOrder, rOrder);
    }
  }

  return makeBoolean(result);
}

function compareScalars(op: string, a: number, b: number): boolean {
  switch (op) {
    case '=':
      return a === b;
    case '<>':
      return a !== b;
    case '<':
      return a < b;
    case '<=':
      return a <= b;
    case '>':
      return a > b;
    case '>=':
      return a >= b;
    default:
      return false;
  }
}

function compareWithCompareResult(op: string, cmp: number): boolean {
  switch (op) {
    case '=':
      return cmp === 0;
    case '<>':
      return cmp !== 0;
    case '<':
      return cmp < 0;
    case '<=':
      return cmp <= 0;
    case '>':
      return cmp > 0;
    case '>=':
      return cmp >= 0;
    default:
      return false;
  }
}
