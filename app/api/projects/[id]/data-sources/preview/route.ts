import { dataSourceRepository, failure, jsonBody, response } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';
type Context = { params: Promise<{ id: string }> };
export async function POST(request: Request, context: Context) {
  try {
    const body = await jsonBody(request); if (typeof body.dataSourceId !== 'string') throw new WorkbookError('请选择数据源');
    return response(await dataSourceRepository().preview((await context.params).id, body.dataSourceId, body.query));
  } catch (e) { return failure(e); }
}
