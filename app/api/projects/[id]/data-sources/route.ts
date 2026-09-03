import { dataSourceRepository, failure, jsonBody, response } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';
type Context = { params: Promise<{ id: string }> };
export async function GET(_request: Request, context: Context) {
  try { return response(await dataSourceRepository().list((await context.params).id)); } catch (e) { return failure(e); }
}
export async function POST(request: Request, context: Context) {
  try { const body = await jsonBody(request); return response(await dataSourceRepository().create((await context.params).id, body.name, body.connectionKey), 201); } catch (e) { return failure(e); }
}
export async function DELETE(request: Request, context: Context) {
  try {
    const body = await jsonBody(request); if (typeof body.id !== 'string') throw new WorkbookError('缺少数据源标识');
    await dataSourceRepository().remove((await context.params).id, body.id); return response({ deleted: true });
  } catch (e) { return failure(e); }
}
