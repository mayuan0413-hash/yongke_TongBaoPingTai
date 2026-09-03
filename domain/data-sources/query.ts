import { WorkbookError } from '../workbook/commands.ts';
import type { DataSourceQuery, SourceColumn } from '../workbook/types.ts';

export const MAX_QUERY_ROWS = 100_000;
export const MAX_QUERY_CELLS = 2_000_000;
export const quoteIdentifier = (name: string) => `"${name.replaceAll('"', '""')}"`;

export function validateQuery(input: unknown): DataSourceQuery {
  if (!input || typeof input !== 'object') throw new WorkbookError('请配置数据查询');
  const q = input as DataSourceQuery;
  if (typeof q.table !== 'string' || !q.table.length || q.table.length > 200) throw new WorkbookError('请选择数据表');
  if (!Array.isArray(q.fields) || !q.fields.length || q.fields.length > 256 || q.fields.some(f => typeof f !== 'string' || !f.length) || new Set(q.fields).size !== q.fields.length) throw new WorkbookError('请至少选择一个字段，且不能重复选择');
  if (!Array.isArray(q.filters) || q.filters.length > 20 || !Array.isArray(q.orderBy) || q.orderBy.length > 8) throw new WorkbookError('查询条件或排序配置无效');
  const operators = new Set(['eq', 'neq', 'gt', 'gte', 'lt', 'lte', 'contains', 'startsWith', 'endsWith', 'isNull', 'notNull']);
  for (const f of q.filters) {
    if (!f || typeof f.field !== 'string' || !operators.has(f.operator)) throw new WorkbookError('查询条件无效');
    if (f.operator !== 'isNull' && f.operator !== 'notNull' && (typeof f.value !== 'string' || f.value.length > 1000)) throw new WorkbookError('请填写条件值（最多 1,000 字符）');
  }
  for (const order of q.orderBy) if (!order || typeof order.field !== 'string' || !['asc', 'desc'].includes(order.direction)) throw new WorkbookError('排序配置无效');
  if (!Number.isInteger(q.rowLimit) || q.rowLimit < 1 || q.rowLimit > MAX_QUERY_ROWS || q.rowLimit * q.fields.length > MAX_QUERY_CELLS) throw new WorkbookError('查询最多 100,000 行、2,000,000 个单元格，请调整行数或字段数量');
  return { table: q.table, fields: [...q.fields], filters: q.filters.map(f => ({ ...f })), orderBy: q.orderBy.map(o => ({ ...o })), rowLimit: q.rowLimit };
}

/** Only identifiers verified against live schema enter SQL; values are always bound. */
export function compileQuery(input: unknown, schema: SourceColumn[], previewRows?: number) {
  const q = validateQuery(input);
  const allowed = new Set(schema.map(c => c.name));
  for (const field of [...q.fields, ...q.filters.map(f => f.field), ...q.orderBy.map(o => o.field)]) if (!allowed.has(field)) throw new WorkbookError(`字段不存在：${field}，请重新配置查询`);
  const params: (string | number)[] = [];
  const operators = { eq: '=', neq: '<>', gt: '>', gte: '>=', lt: '<', lte: '<=' };
  const clauses = q.filters.map(filter => {
    const name = quoteIdentifier(filter.field);
    if (filter.operator === 'isNull') return `${name} IS NULL`;
    if (filter.operator === 'notNull') return `${name} IS NOT NULL`;
    if (filter.operator in operators) { params.push(filter.value!); return `${name} ${operators[filter.operator as keyof typeof operators]} ?`; }
    const escaped = filter.value!.replaceAll('\\', '\\\\').replaceAll('%', '\\%').replaceAll('_', '\\_');
    params.push(filter.operator === 'contains' ? `%${escaped}%` : filter.operator === 'startsWith' ? `${escaped}%` : `%${escaped}`);
    return `${name} LIKE ? ESCAPE '\\'`;
  });
  const limit = previewRows === undefined ? q.rowLimit : Math.min(previewRows, q.rowLimit);
  params.push(limit + 1);
  const sql = `SELECT ${q.fields.map(quoteIdentifier).join(',')} FROM ${quoteIdentifier(q.table)}${clauses.length ? ` WHERE ${clauses.join(' AND ')}` : ''}${q.orderBy.length ? ` ORDER BY ${q.orderBy.map(o => `${quoteIdentifier(o.field)} ${o.direction.toUpperCase()}`).join(',')}` : ''} LIMIT ?`;
  return { sql, params, limit, query: q };
}
