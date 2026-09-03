import { index, integer, primaryKey, sqliteTable, text } from 'drizzle-orm/sqlite-core';

export const projects = sqliteTable('projects', {
  id: text('id').primaryKey(),
  name: text('name').notNull(),
  revision: integer('revision').notNull().default(0),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull(),
});

export const sheets = sqliteTable('sheets', {
  id: text('id').primaryKey(),
  projectId: text('project_id').notNull().references(() => projects.id, { onDelete: 'cascade' }),
  name: text('name').notNull(),
  position: integer('position').notNull(),
  rowCount: integer('row_count').notNull(),
  columnCount: integer('column_count').notNull(),
  rowHeights: text('row_heights').notNull().default('{}'),
  columnWidths: text('column_widths').notNull().default('{}'),
}, table => [index('idx_sheets_project_position').on(table.projectId, table.position)]);

/** One record per populated row; cells inside the row are sparse by column. */
export const sheetRows = sqliteTable('sheet_rows', {
  sheetId: text('sheet_id').notNull().references(() => sheets.id, { onDelete: 'cascade' }),
  rowIndex: integer('row_index').notNull(),
  cells: text('cells').notNull(),
}, table => [primaryKey({ columns: [table.sheetId, table.rowIndex] })]);

export const dataSources = sqliteTable('data_sources', {
  id: text('id').primaryKey(),
  projectId: text('project_id').notNull().references(() => projects.id, { onDelete: 'cascade' }),
  name: text('name').notNull(),
  connectionKey: text('connection_key').notNull(),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull(),
}, table => [index('idx_data_sources_project').on(table.projectId)]);

export const sheetDataQueries = sqliteTable('sheet_data_queries', {
  sheetId: text('sheet_id').primaryKey().references(() => sheets.id, { onDelete: 'cascade' }),
  dataSourceId: text('data_source_id').notNull().references(() => dataSources.id, { onDelete: 'restrict' }),
  tableName: text('table_name').notNull(),
  selectedFields: text('selected_fields').notNull(),
  filters: text('filters').notNull().default('[]'),
  orderBy: text('order_by').notNull().default('[]'),
  rowLimit: integer('row_limit').notNull().default(10000),
  columns: text('columns').notNull().default('[]'),
  lastRefreshedAt: text('last_refreshed_at'),
  lastRowCount: integer('last_row_count').notNull().default(0),
  truncated: integer('truncated', { mode: 'boolean' }).notNull().default(false),
}, table => [index('idx_sheet_queries_data_source').on(table.dataSourceId)]);

/** Refreshed database rows are immutable blocks; manual cells remain in sheet_rows. */
export const sheetDataBlocks = sqliteTable('sheet_data_blocks', {
  sheetId: text('sheet_id').notNull().references(() => sheets.id, { onDelete: 'cascade' }),
  blockIndex: integer('block_index').notNull(),
  rows: text('rows').notNull(),
}, table => [primaryKey({ columns: [table.sheetId, table.blockIndex] })]);
