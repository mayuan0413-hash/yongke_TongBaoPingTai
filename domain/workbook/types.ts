/** Persist only populated cells. Grid dimensions do not allocate cell records. */
export interface Cell { input: string; sourceValue?: string | number | null }

export interface Sheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, Cell>;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  dataSource: SheetDataSourceBinding | null;
}

export type DataSourceConnectionKey = 'business';
export type QueryOperator = 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains' | 'startsWith' | 'endsWith' | 'isNull' | 'notNull';
export interface QueryFilter { field: string; operator: QueryOperator; value?: string }
export interface QueryOrder { field: string; direction: 'asc' | 'desc' }
export interface DataSourceQuery {
  table: string;
  fields: string[];
  filters: QueryFilter[];
  orderBy: QueryOrder[];
  rowLimit: number;
}
export interface SheetDataSourceBinding {
  dataSourceId: string;
  dataSourceName: string;
  connectionKey: DataSourceConnectionKey;
  query: DataSourceQuery;
  columns: string[];
  lastRefreshedAt: string | null;
  lastRowCount: number;
  truncated: boolean;
}
export interface DataSource {
  id: string;
  projectId: string;
  name: string;
  connectionKey: DataSourceConnectionKey;
  createdAt: string;
  updatedAt: string;
}
export interface DataSourceConnection { key: DataSourceConnectionKey; name: string; kind: 'd1-sqlite' }
export interface SourceTable { name: string; label: string }
export interface SourceColumn { name: string; label: string; dataType: string; nullable: boolean }
export interface QueryPreview { columns: string[]; rows: Record<string, string | number | null>[]; hasMore: boolean }

export interface Project {
  id: string;
  name: string;
  sheets: Sheet[];
  revision: number;
  createdAt: string;
  updatedAt: string;
}

export interface ProjectSummary {
  id: string;
  name: string;
  sheetCount: number;
  updatedAt: string;
}

export interface Position { row: number; col: number }
export interface Selection { anchor: Position; focus: Position }
export type Axis = 'row' | 'column';

export type WorkbookCommand =
  | { type: 'renameProject'; name: string }
  | { type: 'addSheet'; id: string; name: string }
  | { type: 'renameSheet'; sheetId: string; name: string }
  | { type: 'deleteSheet'; sheetId: string }
  | { type: 'duplicateSheet'; sheetId: string; id: string; name: string }
  | { type: 'moveSheet'; sheetId: string; toIndex: number }
  | { type: 'setCells'; sheetId: string; changes: { row: number; col: number; input: string }[] }
  | { type: 'clearCells'; sheetId: string; selection: Selection }
  | { type: 'insertAxis' | 'deleteAxis'; sheetId: string; axis: Axis; index: number; count: number }
  | { type: 'resizeAxis'; sheetId: string; axis: Axis; index: number; size: number }
  | { type: 'expandSheet'; sheetId: string; rowCount: number; columnCount: number };

export const DEFAULT_ROWS = 1000;
export const DEFAULT_COLUMNS = 52;
export const MAX_ROWS = 1_000_000;
export const MAX_COLUMNS = 16_384;
export const MAX_PASTE_CELLS = 50_000;
export const DEFAULT_ROW_HEIGHT = 29;
export const DEFAULT_COLUMN_WIDTH = 112;
