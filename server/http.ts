import { env } from 'cloudflare:workers';
import { WorkbookError } from '../domain/workbook/commands.ts';
import { ConflictError, NotFoundError, ProjectRepository } from './project-repository.ts';
import { D1SourceAdapter } from './data-sources/adapter.ts';
import { DataSourceRepository } from './data-sources/repository.ts';

export function repository() {
  const db = (env as unknown as { DB?: D1Database }).DB;
  if (!db) throw new Error('数据库未配置，请先执行 npm run db:migrate');
  return new ProjectRepository(db);
}
export function dataSourceRepository() {
  const bindings = env as unknown as { DB?: D1Database; BUSINESS_DB?: D1Database };
  if (!bindings.DB) throw new Error('应用数据库未配置');
  const source = bindings.BUSINESS_DB ?? bindings.DB;
  return new DataSourceRepository(bindings.DB, new D1SourceAdapter(source));
}
export function sourceAdapter() {
  const bindings = env as unknown as { DB?: D1Database; BUSINESS_DB?: D1Database };
  const source = bindings.BUSINESS_DB ?? bindings.DB;
  if (!source) throw new Error('业务数据库未配置');
  return new D1SourceAdapter(source);
}

export async function jsonBody(request: Request): Promise<Record<string, unknown>> {
  if (!request.headers.get('content-type')?.includes('application/json')) throw new WorkbookError('请求需要 JSON 格式');
  const origin = request.headers.get('origin');
  if (origin && origin !== new URL(request.url).origin) throw new WorkbookError('不接受跨站写入请求');
  const reader = request.body?.getReader();
  if (!reader) throw new WorkbookError('请求为空');
  const decoder = new TextDecoder(); let total = 0; let text = '';
  while (true) {
    const { done, value } = await reader.read(); if (done) break;
    total += value.byteLength;
    if (total > 8_000_000) { await reader.cancel(); throw new WorkbookError('请求过大，请分批提交'); }
    text += decoder.decode(value, { stream: true });
  }
  text += decoder.decode();
  let body: unknown;
  try { body = JSON.parse(text); } catch { throw new WorkbookError('JSON 格式无效'); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) throw new WorkbookError('请求内容无效');
  return body as Record<string, unknown>;
}

export function response(data: unknown, status = 200) { return Response.json(data, { status, headers: { 'Cache-Control': 'no-store' } }); }
export function failure(error: unknown) {
  if (error instanceof ConflictError) return response({ error: error.message }, 409);
  if (error instanceof NotFoundError) return response({ error: error.message }, 404);
  if (error instanceof WorkbookError) return response({ error: error.message }, 400);
  console.error('Project API failed', error);
  return response({ error: '服务暂时无法保存数据，请检查数据库状态后重试。未保存的编辑会保留在当前页面。' }, 500);
}
