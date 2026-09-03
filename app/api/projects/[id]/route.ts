import { failure, jsonBody, repository, response } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';
import type { WorkbookCommand } from '@/domain/workbook/types';

type Context = { params: Promise<{ id: string }> };
export async function GET(_request: Request, context: Context) {
  try { return response(await repository().get((await context.params).id)); } catch (e) { return failure(e); }
}
export async function PATCH(request: Request, context: Context) {
  try {
    const body = await jsonBody(request);
    if (!Number.isInteger(body.revision) || Number(body.revision) < 0) throw new WorkbookError('保存序号无效');
    if (!body.command || typeof body.command !== 'object') throw new WorkbookError('缺少操作内容');
    return response(await repository().execute((await context.params).id, body.revision as number, body.command as WorkbookCommand));
  } catch (e) { return failure(e); }
}
export async function DELETE(request: Request, context: Context) {
  try {
    const body = await jsonBody(request);
    if (!Number.isInteger(body.revision) || Number(body.revision) < 0) throw new WorkbookError('保存序号无效');
    await repository().delete((await context.params).id, body.revision as number);
    return response({ deleted: true });
  } catch (e) { return failure(e); }
}
