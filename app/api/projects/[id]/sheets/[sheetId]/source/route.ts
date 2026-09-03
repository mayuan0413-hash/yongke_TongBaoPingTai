import { dataSourceRepository, failure, jsonBody, response } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';
type Context = { params: Promise<{ id: string; sheetId: string }> };
export async function GET(_request: Request, context: Context) {
  try { const p=await context.params; return response(await dataSourceRepository().binding(p.id,p.sheetId)); } catch(e){return failure(e);}
}
export async function PUT(request: Request, context: Context) {
  try {
    const body=await jsonBody(request); if(typeof body.dataSourceId!=='string'||!Number.isInteger(body.revision)) throw new WorkbookError('数据源配置无效');
    const p=await context.params; return response(await dataSourceRepository().bind(p.id,p.sheetId,body.revision as number,body.dataSourceId,body.query));
  } catch(e){return failure(e);}
}
export async function DELETE(request: Request, context: Context) {
  try {
    const body=await jsonBody(request); if(!Number.isInteger(body.revision)) throw new WorkbookError('保存序号无效');
    const p=await context.params; return response(await dataSourceRepository().unbind(p.id,p.sheetId,body.revision as number));
  } catch(e){return failure(e);}
}
