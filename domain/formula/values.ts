import type { FormulaErrorInfo, FormulaErrorType, FormulaValue, ResolvedRange } from './types.ts';

export function makeNumber(value: number): FormulaValue {
  return Number.isFinite(value) ? { type: 'number', value } : makeError('#NUM!', '数值溢出或无效数值');
}

export function makeString(value: string): FormulaValue {
  return { type: 'string', value };
}

export function makeBoolean(value: boolean): FormulaValue {
  return { type: 'boolean', value };
}

export function makeBlank(): FormulaValue {
  return { type: 'blank' };
}

export function makeError(type: FormulaErrorType, message: string, details?: string): FormulaValue {
  return { type: 'error', error: { type, message, details } };
}

export function makeRange(range: ResolvedRange): FormulaValue {
  return { type: 'range', range };
}

export function isError(val: FormulaValue): val is { type: 'error'; error: FormulaErrorInfo } {
  return val.type === 'error';
}

export function formatFormulaValue(val: FormulaValue): string {
  switch (val.type) {
    case 'number':
      return String(val.value);
    case 'string':
      return val.value;
    case 'boolean':
      return val.value ? 'TRUE' : 'FALSE';
    case 'blank':
      return '';
    case 'error':
      return val.error.type;
    case 'range':
      // Excel returns the top-left cell value when coerced to scalar
      return formatFormulaValue(val.range.getCellValue(val.range.startRow, val.range.startCol));
  }
}

/**
 * Coerce a FormulaValue to number according to Excel rules:
 * - number -> number
 * - blank -> 0
 * - boolean -> 1 / 0
 * - string -> parse float if valid number, else #VALUE!
 * - range -> coerce top-left cell
 * - error -> propagate error
 */
export function toNumber(val: FormulaValue): { ok: true; value: number } | { ok: false; error: FormulaErrorInfo } {
  if (val.type === 'number') return { ok: true, value: val.value };
  if (val.type === 'blank') return { ok: true, value: 0 };
  if (val.type === 'boolean') return { ok: true, value: val.value ? 1 : 0 };
  if (val.type === 'error') return { ok: false, error: val.error };
  if (val.type === 'range') {
    const single = val.range.getCellValue(val.range.startRow, val.range.startCol);
    return toNumber(single);
  }
  if (val.type === 'string') {
    const trimmed = val.value.trim();
    if (trimmed === '') return { ok: true, value: 0 };
    const num = Number(trimmed);
    if (!Number.isNaN(num) && Number.isFinite(num)) return { ok: true, value: num };
    return { ok: false, error: { type: '#VALUE!', message: `无法将文本 "${val.value}" 转换为数值` } };
  }
  return { ok: false, error: { type: '#VALUE!', message: '无法计算的数值类型' } };
}

/**
 * Coerce to string for concatenation & display:
 */
export function toText(val: FormulaValue): { ok: true; text: string } | { ok: false; error: FormulaErrorInfo } {
  if (val.type === 'error') return { ok: false, error: val.error };
  if (val.type === 'range') {
    const single = val.range.getCellValue(val.range.startRow, val.range.startCol);
    return toText(single);
  }
  return { ok: true, text: formatFormulaValue(val) };
}

/**
 * Coerce to boolean:
 */
export function toBoolean(val: FormulaValue): { ok: true; bool: boolean } | { ok: false; error: FormulaErrorInfo } {
  if (val.type === 'boolean') return { ok: true, bool: val.value };
  if (val.type === 'number') return { ok: true, bool: val.value !== 0 };
  if (val.type === 'blank') return { ok: true, bool: false };
  if (val.type === 'error') return { ok: false, error: val.error };
  if (val.type === 'range') {
    const single = val.range.getCellValue(val.range.startRow, val.range.startCol);
    return toBoolean(single);
  }
  if (val.type === 'string') {
    const upper = val.value.trim().toUpperCase();
    if (upper === 'TRUE') return { ok: true, bool: true };
    if (upper === 'FALSE') return { ok: true, bool: false };
    const num = Number(val.value.trim());
    if (!Number.isNaN(num) && Number.isFinite(num)) return { ok: true, bool: num !== 0 };
    return { ok: false, error: { type: '#VALUE!', message: `无法将文本 "${val.value}" 转换为布尔值` } };
  }
  return { ok: false, error: { type: '#VALUE!', message: '无法转换为布尔值' } };
}
