import { WorkbookError } from '../../domain/workbook/commands.ts';
import { compileQuery, quoteIdentifier } from '../../domain/data-sources/query.ts';
import type { DataSourceQuery, QueryPreview, SourceColumn, SourceTable } from '../../domain/workbook/types.ts';

export interface DataSourceAdapter {
  listTables(): Promise<SourceTable[]>;
  columns(table: string): Promise<SourceColumn[]>;
  query(query: DataSourceQuery, previewRows?: number): Promise<QueryPreview>;
}

/** A connection is selected from server-owned bindings, never from client credentials. */
export class D1SourceAdapter implements DataSourceAdapter {
  private db: D1Database;
  constructor(db: D1Database) { this.db = db; }

  async listTables(): Promise<SourceTable[]> {
    const result = await this.db.prepare("SELECT name FROM sqlite_schema WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name NOT LIKE '_cf_%' AND name NOT LIKE 'd1_%' AND name NOT LIKE '_source_%' AND name NOT IN ('projects','sheets','sheet_rows','data_sources','sheet_data_queries','sheet_data_blocks','__drizzle_migrations') ORDER BY name").all<{ name: string }>();
    return result.results.map(row => ({ name: row.name, label: row.name }));
  }
  async columns(table: string): Promise<SourceColumn[]> {
    if (!(await this.listTables()).some(t => t.name === table)) throw new WorkbookError('数据表不存在，请重新选择');
    const result = await this.db.prepare(`PRAGMA table_info(${quoteIdentifier(table)})`).all<{ name: string; type: string; notnull: number }>();
    const hasMeta = await this.db.prepare("SELECT name FROM sqlite_schema WHERE type='table' AND name='_source_column_labels'").first();
    const labels = hasMeta ? (await this.db.prepare('SELECT column_name, label FROM _source_column_labels WHERE table_name=?').bind(table).all<{ column_name: string; label: string }>()).results : [];
    return result.results.map(row => ({ name: row.name, label: labels.find(l => l.column_name === row.name)?.label || row.name, dataType: row.type, nullable: !row.notnull }));
  }
  async query(query: DataSourceQuery, previewRows?: number): Promise<QueryPreview> {
    const schema = await this.columns(query.table);
    const compiled = compileQuery(query, schema, previewRows);
    const result = await this.db.prepare(compiled.sql).bind(...compiled.params).all<Record<string, string | number | null>>();
    for (const row of result.results) for (const value of Object.values(row)) if (value !== null && typeof value !== 'string' && typeof value !== 'number') throw new WorkbookError('当前查询包含二进制字段，请取消选择该字段');
    return { columns: compiled.query.fields, rows: result.results.slice(0, compiled.limit), hasMore: result.results.length > compiled.limit };
  }
}
