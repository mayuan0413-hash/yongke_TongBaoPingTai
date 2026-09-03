import { applyCommand, createProject } from '../domain/workbook/commands.ts';
import { cellKey, fromKey } from '../domain/workbook/address.ts';
import type { Project, ProjectSummary, Sheet, WorkbookCommand } from '../domain/workbook/types.ts';
import { blocksToCells } from '../domain/data-sources/change-set.ts';
import { mapBinding, type BindingRecord } from './data-sources/repository.ts';
import { ConflictError, NotFoundError } from './errors.ts';

export { ConflictError, NotFoundError } from './errors.ts';
type ProjectRecord = { id: string; name: string; revision: number; created_at: string; updated_at: string };
type SheetRecord = { id: string; name: string; row_count: number; column_count: number; row_heights: string; column_widths: string };

function rowsOf(sheet: Sheet): Map<number, string> {
  const rows = new Map<number, Record<string, { input: string }>>();
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const { row, col } = fromKey(key);
    if (!rows.has(row)) rows.set(row, {});
    rows.get(row)![col] = cell;
  }
  return new Map([...rows].map(([row, values]) => [row, JSON.stringify(values)]));
}

/** D1 also runs as durable local SQLite in development. All writes are atomic. */
export class ProjectRepository {
  private db: D1Database;
  constructor(db: D1Database) { this.db = db; }

  async list(): Promise<ProjectSummary[]> {
    const result = await this.db.prepare('SELECT p.id, p.name, p.updated_at, COUNT(s.id) AS sheet_count FROM projects p LEFT JOIN sheets s ON s.project_id = p.id GROUP BY p.id ORDER BY p.updated_at DESC').all<{ id: string; name: string; updated_at: string; sheet_count: number }>();
    return result.results.map(r => ({ id: r.id, name: r.name, updatedAt: r.updated_at, sheetCount: r.sheet_count }));
  }

  async get(id: string): Promise<Project> {
    const [projectResult, sheetsResult, rowsResult, bindingsResult, blocksResult] = await this.db.batch([
      this.db.prepare('SELECT * FROM projects WHERE id = ?').bind(id),
      this.db.prepare('SELECT * FROM sheets WHERE project_id = ? ORDER BY position').bind(id),
      this.db.prepare('SELECT r.sheet_id, r.row_index, r.cells FROM sheet_rows r JOIN sheets s ON s.id = r.sheet_id WHERE s.project_id = ? ORDER BY r.row_index').bind(id),
      this.db.prepare('SELECT q.*, d.name AS data_source_name, d.connection_key FROM sheet_data_queries q JOIN sheets s ON s.id=q.sheet_id JOIN data_sources d ON d.id=q.data_source_id WHERE s.project_id=?').bind(id),
      this.db.prepare('SELECT b.* FROM sheet_data_blocks b JOIN sheets s ON s.id=b.sheet_id WHERE s.project_id=? ORDER BY b.block_index').bind(id),
    ]);
    const record = projectResult.results[0] as ProjectRecord | undefined;
    if (!record) throw new NotFoundError('项目不存在');
    const sheets = (sheetsResult.results as SheetRecord[]).map(s => ({ id: s.id, name: s.name, rowCount: s.row_count, columnCount: s.column_count, cells: {}, rowHeights: JSON.parse(s.row_heights), columnWidths: JSON.parse(s.column_widths), dataSource: null } as Sheet));
    const byId = new Map(sheets.map(s => [s.id, s]));
    for (const row of rowsResult.results as { sheet_id: string; row_index: number; cells: string }[]) {
      const sheet = byId.get(row.sheet_id)!;
      for (const [col, cell] of Object.entries(JSON.parse(row.cells) as Record<string, { input: string }>)) sheet.cells[cellKey(row.row_index, Number(col))] = cell;
    }
    for (const record of bindingsResult.results as BindingRecord[]) {
      const sheet = byId.get(record.sheet_id)!;
      sheet.dataSource = mapBinding(record);
      if (sheet.dataSource.lastRefreshedAt) sheet.cells = blocksToCells(sheet.dataSource.columns, (blocksResult.results as { sheet_id: string; block_index: number; rows: string }[]).filter(b => b.sheet_id === sheet.id));
    }
    return { id: record.id, name: record.name, revision: record.revision, createdAt: record.created_at, updatedAt: record.updated_at, sheets };
  }

  async create(name: string): Promise<Project> {
    const project = createProject(name, crypto.randomUUID(), crypto.randomUUID());
    const sheet = project.sheets[0];
    await this.db.batch([
      this.db.prepare('INSERT INTO projects(id,name,revision,created_at,updated_at) VALUES(?,?,?,?,?)').bind(project.id, project.name, 0, project.createdAt, project.updatedAt),
      this.db.prepare('INSERT INTO sheets(id,project_id,name,position,row_count,column_count) VALUES(?,?,?,?,?,?)').bind(sheet.id, project.id, sheet.name, 0, sheet.rowCount, sheet.columnCount),
    ]);
    return project;
  }

  async execute(id: string, expectedRevision: number, command: WorkbookCommand): Promise<{ revision: number; updatedAt: string }> {
    const before = await this.get(id);
    if (before.revision !== expectedRevision) throw new ConflictError('项目已被其他页面修改，请重新加载后再编辑');
    const after = applyCommand(before, command);
    const guard = 'EXISTS (SELECT 1 FROM projects WHERE id = ? AND revision = ?)';
    const statements: D1PreparedStatement[] = [];
    for (const oldSheet of before.sheets) if (!after.sheets.some(s => s.id === oldSheet.id)) statements.push(this.db.prepare(`DELETE FROM sheets WHERE id = ? AND ${guard}`).bind(oldSheet.id, id, expectedRevision));
    for (const [position, sheet] of after.sheets.entries()) {
      const oldSheet = before.sheets.find(s => s.id === sheet.id);
      const previousPosition = before.sheets.findIndex(s => s.id === sheet.id);
      if (oldSheet !== sheet || previousPosition !== position) {
        statements.push(this.db.prepare(`INSERT INTO sheets(id,project_id,name,position,row_count,column_count,row_heights,column_widths) SELECT ?,?,?,?,?,?,?,? WHERE ${guard} ON CONFLICT(id) DO UPDATE SET name=excluded.name,position=excluded.position,row_count=excluded.row_count,column_count=excluded.column_count,row_heights=excluded.row_heights,column_widths=excluded.column_widths`).bind(sheet.id, id, sheet.name, position, sheet.rowCount, sheet.columnCount, JSON.stringify(sheet.rowHeights), JSON.stringify(sheet.columnWidths), id, expectedRevision));
      }
      if (sheet.dataSource && !oldSheet && command.type === 'duplicateSheet') {
        statements.push(this.db.prepare(`INSERT INTO sheet_data_queries SELECT ?,data_source_id,table_name,selected_fields,filters,order_by,row_limit,columns,last_refreshed_at,last_row_count,truncated FROM sheet_data_queries WHERE sheet_id=? AND ${guard}`).bind(sheet.id, command.sheetId, id, expectedRevision));
        statements.push(this.db.prepare(`INSERT INTO sheet_data_blocks SELECT ?,block_index,rows FROM sheet_data_blocks WHERE sheet_id=? AND ${guard}`).bind(sheet.id, command.sheetId, id, expectedRevision));
      }
      if (oldSheet?.cells === sheet.cells || sheet.dataSource?.lastRefreshedAt) continue;
      const previousRows = oldSheet ? rowsOf(oldSheet) : new Map<number, string>();
      const nextRows = rowsOf(sheet);
      for (const row of previousRows.keys()) if (!nextRows.has(row)) statements.push(this.db.prepare(`DELETE FROM sheet_rows WHERE sheet_id = ? AND row_index = ? AND ${guard}`).bind(sheet.id, row, id, expectedRevision));
      const changed = [...nextRows].filter(([row, cells]) => cells !== previousRows.get(row));
      // Batch values under D1's 100 bind-parameter limit; no blank cells stored.
      for (let start = 0; start < changed.length; start += 25) {
        const chunk = changed.slice(start, start + 25);
        const values = chunk.map(() => '(?,?,?)').join(',');
        statements.push(this.db.prepare(`WITH incoming(sheet_id,row_index,cells) AS (VALUES ${values}) INSERT INTO sheet_rows(sheet_id,row_index,cells) SELECT sheet_id,row_index,cells FROM incoming WHERE ${guard} ON CONFLICT(sheet_id,row_index) DO UPDATE SET cells=excluded.cells`).bind(...chunk.flatMap(([row, cells]) => [sheet.id, row, cells]), id, expectedRevision));
      }
    }
    statements.push(this.db.prepare('UPDATE projects SET name=?,revision=?,updated_at=? WHERE id=? AND revision=?').bind(after.name, after.revision, after.updatedAt, id, expectedRevision));
    const result = await this.db.batch(statements);
    if (result[result.length - 1].meta.changes !== 1) throw new ConflictError('项目已被其他页面修改，请重新加载后再编辑');
    return { revision: after.revision, updatedAt: after.updatedAt };
  }

  async delete(id: string, revision: number) {
    const results = await this.db.batch([
      this.db.prepare('DELETE FROM sheet_data_queries WHERE sheet_id IN(SELECT id FROM sheets WHERE project_id=?) AND EXISTS(SELECT 1 FROM projects WHERE id=? AND revision=?)').bind(id, id, revision),
      this.db.prepare('DELETE FROM projects WHERE id=? AND revision=? RETURNING id').bind(id, revision),
    ]);
    const result = results[1];
    if (result.results.length !== 1) throw new ConflictError('项目已变化或不存在，请刷新项目列表');
  }
}
