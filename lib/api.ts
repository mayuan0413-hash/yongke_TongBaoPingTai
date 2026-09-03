import type { DataSource, DataSourceConnection, DataSourceQuery, Project, ProjectSummary, QueryPreview, SourceColumn, SourceTable, WorkbookCommand } from '@/domain/workbook/types';
import type { RefreshResult } from '@/domain/data-sources/change-set';

async function request<T>(url: string, init?: RequestInit): Promise<T> {
 const response = await fetch(url, init);
 const data = await response.json().catch(() => ({})) as { error?: string };
 if (!response.ok) throw new Error(data.error || '请求失败');
 return data as T;
}
const json = (body: unknown): RequestInit => ({ method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
export const api = {
 list: () => request<ProjectSummary[]>('/api/projects'),
 get: (id: string) => request<Project>(`/api/projects/${encodeURIComponent(id)}`),
 create: (name: string) => request<Project>('/api/projects', json({ name })),
 command: (id: string, revision: number, command: WorkbookCommand) => request<{revision:number;updatedAt:string}>(`/api/projects/${encodeURIComponent(id)}`, { ...json({ revision, command }), method: 'PATCH' }),
 delete: (id: string, revision: number) => request<{deleted:true}>(`/api/projects/${encodeURIComponent(id)}`, { ...json({ revision }), method: 'DELETE' }),
 connections: () => request<DataSourceConnection[]>('/api/data-sources/connections'),
 tables: (connection: string) => request<SourceTable[]>(`/api/data-sources/schema?connection=${encodeURIComponent(connection)}`),
 columns: (connection: string, table: string) => request<SourceColumn[]>(`/api/data-sources/schema?connection=${encodeURIComponent(connection)}&table=${encodeURIComponent(table)}`),
 dataSources: (projectId: string) => request<DataSource[]>(`/api/projects/${encodeURIComponent(projectId)}/data-sources`),
 createDataSource: (projectId: string, name: string, connectionKey: string) => request<DataSource>(`/api/projects/${encodeURIComponent(projectId)}/data-sources`, json({ name, connectionKey })),
 deleteDataSource: (projectId: string, id: string) => request<{deleted:true}>(`/api/projects/${encodeURIComponent(projectId)}/data-sources`, { ...json({ id }), method: 'DELETE' }),
 preview: (projectId: string, dataSourceId: string, query: DataSourceQuery) => request<QueryPreview>(`/api/projects/${encodeURIComponent(projectId)}/data-sources/preview`, json({ dataSourceId, query })),
 bindSource: (projectId: string, sheetId: string, revision: number, dataSourceId: string, query: DataSourceQuery) => request<{revision:number;updatedAt:string}>(`/api/projects/${encodeURIComponent(projectId)}/sheets/${encodeURIComponent(sheetId)}/source`, { ...json({ revision, dataSourceId, query }), method: 'PUT' }),
 unbindSource: (projectId: string, sheetId: string, revision: number) => request<{revision:number;updatedAt:string}>(`/api/projects/${encodeURIComponent(projectId)}/sheets/${encodeURIComponent(sheetId)}/source`, { ...json({ revision }), method: 'DELETE' }),
 refreshSource: (projectId: string, sheetId: string, revision: number) => request<RefreshResult>(`/api/projects/${encodeURIComponent(projectId)}/sheets/${encodeURIComponent(sheetId)}/source/refresh`, json({ revision })),
};
