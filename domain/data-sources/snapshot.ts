import { WorkbookError } from '../workbook/commands.ts';

export interface DataBlock { block_index: number; rows: string }
export function makeDataBlocks(rows: Record<string, string | number | null>[], fields: string[]): DataBlock[] {
  const blocks: DataBlock[] = [];
  for (let start = 0; start < rows.length; start += 128) {
    const values = rows.slice(start, start + 128).map(row => fields.map(field => {
      const value = row[field] ?? null;
      if (typeof value === 'string' && value.length > 32767) throw new WorkbookError('数据中存在超过 32,767 字符的单元格，请缩小字段内容');
      return value;
    }));
    const payload = JSON.stringify({ startRow: start + 1, values });
    if (new TextEncoder().encode(payload).length > 1_500_000) throw new WorkbookError('查询内容过大，请减少长文本字段');
    blocks.push({ block_index: blocks.length, rows: payload });
  }
  return blocks;
}
