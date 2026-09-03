import { failure, jsonBody, repository, response } from '@/server/http';
import { WorkbookError } from '@/domain/workbook/commands';

export async function GET() {
  try { return response(await repository().list()); } catch (e) { return failure(e); }
}
export async function POST(request: Request) {
  try {
    const body = await jsonBody(request);
    if (typeof body.name !== 'string') throw new WorkbookError('请输入项目名称');
    return response(await repository().create(body.name), 201);
  } catch (e) { return failure(e); }
}
