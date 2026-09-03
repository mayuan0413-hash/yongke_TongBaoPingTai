import { failure, response, sourceAdapter } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';
export async function GET(request: Request) {
  try {
    const url = new URL(request.url);
    if (url.searchParams.get('connection') !== 'business') throw new WorkbookError('数据库连接不存在');
    const table = url.searchParams.get('table');
    return response(table ? await sourceAdapter().columns(table) : await sourceAdapter().listTables());
  } catch (e) { return failure(e); }
}
