import type { ASTNode, CalculationContext, FormulaValue } from './types.ts';
import {
  formatFormulaValue,
  isError,
  makeBlank,
  makeBoolean,
  makeError,
  makeNumber,
  toBoolean,
  toNumber,
} from './values.ts';

export type FormulaFunctionHandler = (args: ASTNode[], context: CalculationContext) => FormulaValue;

export class FunctionRegistry {
  private handlers = new Map<string, FormulaFunctionHandler>();

  register(name: string, handler: FormulaFunctionHandler): void {
    this.handlers.set(name.trim().toUpperCase(), handler);
  }

  get(name: string): FormulaFunctionHandler | undefined {
    return this.handlers.get(name.trim().toUpperCase());
  }

  has(name: string): boolean {
    return this.handlers.has(name.trim().toUpperCase());
  }

  list(): string[] {
    return Array.from(this.handlers.keys()).sort();
  }
}

export const defaultRegistry = new FunctionRegistry();

/**
 * Register core Milestone 3 functions with lazy-evaluation capability.
 */

// IF(condition, value_if_true, [value_if_false])
defaultRegistry.register('IF', (args, context) => {
  if (args.length < 2 || args.length > 3) {
    return makeError('#VALUE!', 'IF 函数需要 2 或 3 个参数');
  }

  // Lazy evaluation: only evaluate condition first
  const condVal = context.evalNode(args[0]);
  if (isError(condVal)) return condVal;

  const boolRes = toBoolean(condVal);
  if (!boolRes.ok) return makeError('#VALUE!', boolRes.error.message);

  if (boolRes.bool) {
    // Only evaluate true branch
    return context.evalNode(args[1]);
  } else {
    // Only evaluate false branch
    if (args.length === 3) {
      return context.evalNode(args[2]);
    }
    return makeBoolean(false);
  }
});

// IFERROR(value, value_if_error)
defaultRegistry.register('IFERROR', (args, context) => {
  if (args.length !== 2) {
    return makeError('#VALUE!', 'IFERROR 函数需要 2 个参数');
  }

  // Lazy evaluation: evaluate first arg; only evaluate fallback if first is error
  const val = context.evalNode(args[0]);
  if (isError(val)) {
    return context.evalNode(args[1]);
  }
  return val;
});

// SUM(val1, [val2], ...)
defaultRegistry.register('SUM', (args, context) => {
  let sum = 0;
  for (const argNode of args) {
    const val = context.evalNode(argNode);
    if (isError(val)) return val;

    if (val.type === 'range') {
      const flat = val.range.flatten();
      for (const cellVal of flat) {
        if (isError(cellVal)) return cellVal;
        if (cellVal.type === 'number') {
          sum += cellVal.value;
        }
      }
    } else if (val.type === 'number') {
      sum += val.value;
    } else if (val.type === 'blank') {
      // Ignored
    } else if (val.type === 'string' || val.type === 'boolean') {
      // Literal string/bool passed directly to SUM is coerced in Excel
      const numRes = toNumber(val);
      if (!numRes.ok) return makeError(numRes.error.type, numRes.error.message);
      sum += numRes.value;
    }
  }
  return makeNumber(sum);
});

// COUNT(val1, [val2], ...)
defaultRegistry.register('COUNT', (args, context) => {
  let count = 0;
  for (const argNode of args) {
    const val = context.evalNode(argNode);
    if (isError(val)) return val;

    if (val.type === 'range') {
      const flat = val.range.flatten();
      for (const cellVal of flat) {
        if (cellVal.type === 'number') count++;
      }
    } else if (val.type === 'number') {
      count++;
    } else if (val.type === 'string' || val.type === 'boolean') {
      const numRes = toNumber(val);
      if (numRes.ok) count++;
    }
  }
  return makeNumber(count);
});

// AVERAGE(val1, [val2], ...)
defaultRegistry.register('AVERAGE', (args, context) => {
  let sum = 0;
  let count = 0;
  for (const argNode of args) {
    const val = context.evalNode(argNode);
    if (isError(val)) return val;

    if (val.type === 'range') {
      const flat = val.range.flatten();
      for (const cellVal of flat) {
        if (isError(cellVal)) return cellVal;
        if (cellVal.type === 'number') {
          sum += cellVal.value;
          count++;
        }
      }
    } else if (val.type === 'number') {
      sum += val.value;
      count++;
    } else if (val.type === 'string' || val.type === 'boolean') {
      const numRes = toNumber(val);
      if (!numRes.ok) return makeError(numRes.error.type, numRes.error.message);
      sum += numRes.value;
      count++;
    }
  }
  if (count === 0) return makeError('#DIV/0!', 'AVERAGE 分母为 0');
  return makeNumber(sum / count);
});

// INT(number)
defaultRegistry.register('INT', (args, context) => {
  if (args.length !== 1) return makeError('#VALUE!', 'INT 函数需要 1 个参数');
  const val = context.evalNode(args[0]);
  if (isError(val)) return val;
  const numRes = toNumber(val);
  if (!numRes.ok) return makeError(numRes.error.type, numRes.error.message);
  return makeNumber(Math.floor(numRes.value));
});
