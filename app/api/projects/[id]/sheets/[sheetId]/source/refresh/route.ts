import { dataSourceRepository, failure, jsonBody, response } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';
type Context = { params: Promise<{ id: string; sheetId: string }> };
export async function POST(request: Request, context: Context) {
  try {
    const body=await jsonBody(request); if(!Number.isInteger(body.revision)) throw new WorkbookError('保存序号无效');
    const p=await context.params; return response(await dataSourceRepository().refresh(p.id,p.sheetId,body.revision as number));
  } catch(e){return failure(e);}
}
