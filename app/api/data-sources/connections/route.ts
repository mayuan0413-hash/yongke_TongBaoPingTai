import { response } from '@/server/http';
export async function GET() { return response([{ key: 'business', name: '业务数据库', kind: 'd1-sqlite' }]); }
