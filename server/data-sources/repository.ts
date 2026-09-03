import type { DataSource, SheetDataSourceBinding } from '../../domain/workbook/types.ts';
import { DEFAULT_COLUMNS, DEFAULT_ROWS } from '../../domain/workbook/types.ts';
import { WorkbookError } from '../../domain/workbook/commands.ts';
import { validateQuery } from '../../domain/data-sources/query.ts';
import { blocksToCells, type DataRange, type RefreshResult } from '../../domain/data-sources/change-set.ts';
import { makeDataBlocks, type DataBlock } from '../../domain/data-sources/snapshot.ts';
import { ConflictError, NotFoundError } from '../errors.ts';
import type { DataSourceAdapter } from './adapter.ts';

export interface BindingRecord {
  sheet_id: string; data_source_id: string; data_source_name: string; connection_key: 'business';
  table_name: string; selected_fields: string; filters: string; order_by: string; row_limit: number;
  columns: string; last_refreshed_at: string | null; last_row_count: number; truncated: number;
}
export function mapBinding(row: BindingRecord): SheetDataSourceBinding {
  return { dataSourceId: row.data_source_id, dataSourceName: row.data_source_name, connectionKey: row.connection_key,
    query: { table: row.table_name, fields: JSON.parse(row.selected_fields), filters: JSON.parse(row.filters), orderBy: JSON.parse(row.order_by), rowLimit: row.row_limit },
    columns: JSON.parse(row.columns), lastRefreshedAt: row.last_refreshed_at, lastRowCount: row.last_row_count, truncated: !!row.truncated };
}
type SourceRecord = { id: string; project_id: string; name: string; connection_key: 'business'; created_at: string; updated_at: string };
const mapSource = (row: SourceRecord): DataSource => ({ id: row.id, projectId: row.project_id, name: row.name, connectionKey: row.connection_key, createdAt: row.created_at, updatedAt: row.updated_at });

export class DataSourceRepository {
  private db: D1Database;
  private adapter: DataSourceAdapter;
  constructor(db: D1Database, adapter: DataSourceAdapter) { this.db = db; this.adapter = adapter; }
  async list(projectId: string) {
    return (await this.db.prepare('SELECT * FROM data_sources WHERE project_id=? ORDER BY created_at').bind(projectId).all<SourceRecord>()).results.map(mapSource);
  }
  async create(projectId: string, name: unknown, connection: unknown) {
    if (typeof name !== 'string' || !name.trim() || name.trim().length > 80) throw new WorkbookError('数据源名称需要 1–80 个字符');
    if (connection !== 'business') throw new WorkbookError('数据库连接不存在');
    if (!(await this.db.prepare('SELECT id FROM projects WHERE id=?').bind(projectId).first())) throw new NotFoundError('项目不存在');
    const id = crypto.randomUUID(), now = new Date().toISOString();
    await this.db.prepare('INSERT INTO data_sources(id,project_id,name,connection_key,created_at,updated_at) VALUES(?,?,?,?,?,?)').bind(id, projectId, name.trim(), connection, now, now).run();
    return { id, projectId, name: name.trim(), connectionKey: 'business' as const, createdAt: now, updatedAt: now };
  }
  async remove(projectId: string, sourceId: string) {
    const used = await this.db.prepare('SELECT q.sheet_id FROM sheet_data_queries q JOIN data_sources d ON d.id=q.data_source_id WHERE d.id=? AND d.project_id=? LIMIT 1').bind(sourceId, projectId).first();
    if (used) throw new WorkbookError('数据源正在被 Sheet 使用，请先切换为手工数据');
    const result = await this.db.prepare('DELETE FROM data_sources WHERE id=? AND project_id=? RETURNING id').bind(sourceId, projectId).all();
    if (!result.results.length) throw new NotFoundError('数据源不存在');
  }
  private async check(projectId: string, sheetId: string, revision: number) {
    if (!Number.isInteger(revision) || revision < 0) throw new WorkbookError('保存序号无效');
    const row = await this.db.prepare('SELECT p.revision FROM projects p JOIN sheets s ON s.project_id=p.id WHERE p.id=? AND s.id=?').bind(projectId, sheetId).first<{ revision: number }>();
    if (!row) throw new NotFoundError('项目或 Sheet 不存在');
    if (row.revision !== revision) throw new ConflictError('数据已发生变化，请重新加载项目再操作');
  }
  async binding(projectId: string, sheetId: string) {
    const record = await this.db.prepare('SELECT q.*, d.name AS data_source_name, d.connection_key FROM sheet_data_queries q JOIN data_sources d ON d.id=q.data_source_id JOIN sheets s ON s.id=q.sheet_id WHERE s.project_id=? AND s.id=?').bind(projectId, sheetId).first<BindingRecord>();
    return record ? mapBinding(record) : null;
  }
  async preview(projectId: string, sourceId: string, input: unknown) {
    if (!(await this.db.prepare('SELECT id FROM data_sources WHERE project_id=? AND id=?').bind(projectId, sourceId).first())) throw new NotFoundError('数据源不存在');
    return this.adapter.query(validateQuery(input), 20);
  }
  private async commit(projectId: string, revision: number, statements: D1PreparedStatement[], now: string) {
    statements.push(this.db.prepare('UPDATE projects SET revision=revision+1,updated_at=? WHERE id=? AND revision=? RETURNING revision').bind(now, projectId, revision));
    const results = await this.db.batch(statements);
    if (results.at(-1)!.results.length !== 1) throw new ConflictError('数据已发生变化，本次操作未写入，请重新加载项目');
    return { revision: revision + 1, updatedAt: now };
  }
  async bind(projectId: string, sheetId: string, revision: number, sourceId: string, input: unknown) {
    await this.check(projectId, sheetId, revision);
    if (!(await this.db.prepare('SELECT id FROM data_sources WHERE id=? AND project_id=?').bind(sourceId, projectId).first())) throw new NotFoundError('数据源不存在');
    const query = validateQuery(input);
    // Probe the real schema/query before accepting configuration; existing data survives failures.
    await this.adapter.query(query, 1);
    const now = new Date().toISOString();
    const result = await this.commit(projectId, revision, [this.db.prepare('INSERT INTO sheet_data_queries(sheet_id,data_source_id,table_name,selected_fields,filters,order_by,row_limit) SELECT ?,?,?,?,?,?,? WHERE EXISTS(SELECT 1 FROM projects WHERE id=? AND revision=?) ON CONFLICT(sheet_id) DO UPDATE SET data_source_id=excluded.data_source_id,table_name=excluded.table_name,selected_fields=excluded.selected_fields,filters=excluded.filters,order_by=excluded.order_by,row_limit=excluded.row_limit').bind(sheetId, sourceId, query.table, JSON.stringify(query.fields), JSON.stringify(query.filters), JSON.stringify(query.orderBy), query.rowLimit, projectId, revision)], now);
    return result;
  }
  async refresh(projectId: string, sheetId: string, revision: number): Promise<RefreshResult> {
    await this.check(projectId, sheetId, revision);
    const binding = await this.binding(projectId, sheetId);
    if (!binding) throw new WorkbookError('该 Sheet 尚未绑定数据库');
    const result = await this.adapter.query(binding.query);
    const schema = await this.adapter.columns(binding.query.table);
    const labels = result.columns.map(name => schema.find(c => c.name === name)?.label || name);
    const blocks = makeDataBlocks(result.rows, result.columns);
    const previous = (await this.db.prepare('SELECT block_index,rows FROM sheet_data_blocks WHERE sheet_id=? ORDER BY block_index').bind(sheetId).all<DataBlock>()).results;
    const guard = 'EXISTS(SELECT 1 FROM projects WHERE id=? AND revision=?)';
    const statements: D1PreparedStatement[] = [];
    const dirtyRanges: DataRange[] = [];
    const maxCol = Math.max(labels.length, binding.columns.length) - 1;
    if (!binding.lastRefreshedAt || JSON.stringify(labels) !== JSON.stringify(binding.columns)) dirtyRanges.push({ startRow: 0, endRow: Math.max(binding.lastRowCount, result.rows.length), startColumn: 0, endColumn: maxCol });
    for (let i = 0; i < Math.max(blocks.length, previous.length); i++) {
      if (blocks[i]?.rows === previous[i]?.rows) continue;
      dirtyRanges.push({ startRow: i * 128 + 1, endRow: Math.min((i + 1) * 128, Math.max(binding.lastRowCount, result.rows.length)), startColumn: 0, endColumn: maxCol });
      if (blocks[i]) statements.push(this.db.prepare(`INSERT INTO sheet_data_blocks(sheet_id,block_index,rows) SELECT ?,?,? WHERE ${guard} ON CONFLICT(sheet_id,block_index) DO UPDATE SET rows=excluded.rows`).bind(sheetId, i, blocks[i].rows, projectId, revision));
      else statements.push(this.db.prepare(`DELETE FROM sheet_data_blocks WHERE sheet_id=? AND block_index=? AND ${guard}`).bind(sheetId, i, projectId, revision));
    }
    const now = new Date().toISOString();
    // One atomic publication of the new snapshot; failures never erase the old snapshot.
    statements.push(this.db.prepare(`DELETE FROM sheet_rows WHERE sheet_id=? AND ${guard}`).bind(sheetId, projectId, revision));
    statements.push(this.db.prepare(`UPDATE sheets SET row_count=?,column_count=? WHERE id=? AND ${guard}`).bind(Math.max(DEFAULT_ROWS, result.rows.length + 1), Math.max(DEFAULT_COLUMNS, labels.length), sheetId, projectId, revision));
    statements.push(this.db.prepare(`UPDATE sheet_data_queries SET columns=?,last_refreshed_at=?,last_row_count=?,truncated=? WHERE sheet_id=? AND ${guard}`).bind(JSON.stringify(labels), now, result.rows.length, result.hasMore ? 1 : 0, sheetId, projectId, revision));
    const committed = await this.commit(projectId, revision, statements, now);
    return { ...committed, rowCount: result.rows.length, columnCount: labels.length, truncated: result.hasMore,
      changeSet: { projectId, sheetId, revision: committed.revision, reason: 'data-refresh', dirtyRanges, previousDataRows: binding.lastRowCount, dataRows: result.rows.length, rowDelta: result.rows.length - binding.lastRowCount } };
  }
  async unbind(projectId: string, sheetId: string, revision: number) {
    await this.check(projectId, sheetId, revision);
    const binding = await this.binding(projectId, sheetId);
    if (!binding) throw new WorkbookError('该 Sheet 已是手工数据');
    const guard = 'EXISTS(SELECT 1 FROM projects WHERE id=? AND revision=?)';
    const statements: D1PreparedStatement[] = [];
    if (binding.lastRefreshedAt) {
      const blocks = (await this.db.prepare('SELECT block_index,rows FROM sheet_data_blocks WHERE sheet_id=? ORDER BY block_index').bind(sheetId).all<DataBlock>()).results;
      const cells = blocksToCells(binding.columns, blocks), rows = new Map<number, Record<string, { input: string }>>();
      for (const [key, cell] of Object.entries(cells)) { const [r, c] = key.split(':').map(Number); if (!rows.has(r)) rows.set(r, {}); rows.get(r)![c] = { input: cell.input }; }
      statements.push(this.db.prepare(`DELETE FROM sheet_rows WHERE sheet_id=? AND ${guard}`).bind(sheetId, projectId, revision));
      const entries = [...rows];
      for (let i = 0; i < entries.length; i += 25) {
        const batch = entries.slice(i, i + 25);
        statements.push(this.db.prepare(`WITH incoming(sheet_id,row_index,cells) AS (VALUES ${batch.map(() => '(?,?,?)').join(',')}) INSERT INTO sheet_rows SELECT sheet_id,row_index,cells FROM incoming WHERE ${guard}`).bind(...batch.flatMap(([r, values]) => [sheetId, r, JSON.stringify(values)]), projectId, revision));
      }
    }
    statements.push(this.db.prepare(`DELETE FROM sheet_data_blocks WHERE sheet_id=? AND ${guard}`).bind(sheetId, projectId, revision));
    statements.push(this.db.prepare(`DELETE FROM sheet_data_queries WHERE sheet_id=? AND ${guard}`).bind(sheetId, projectId, revision));
    return this.commit(projectId, revision, statements, new Date().toISOString());
  }
}
