import { bounds, cellKey } from './address.ts';
import { MAX_COLUMNS, MAX_PASTE_CELLS, MAX_ROWS, type Position, type Selection, type Sheet } from './types.ts';

/** Excel clipboard TSV supports quoted multiline fields and escaped quotes. */
export function parseTsv(text: string): string[][] {
  if (text.length > 8_000_000) throw new Error('粘贴内容过大，请分批粘贴');
  const rows: string[][] = []; let row: string[] = []; let value = ''; let quoted = false; let count = 0;
  const pushCell = () => { if (++count > MAX_PASTE_CELLS) throw new Error(`一次最多粘贴 ${MAX_PASTE_CELLS.toLocaleString()} 个单元格`); row.push(value); value = ''; };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i + 1] === '"') { value += '"'; i++; }
      else if (quoted || value.length === 0) quoted = !quoted;
      else value += c;
    } else if (!quoted && c === '\t') pushCell();
    else if (!quoted && (c === '\r' || c === '\n')) {
      pushCell(); rows.push(row); row = [];
      if (c === '\r' && text[i + 1] === '\n') i++;
    } else value += c;
  }
  if (quoted) throw new Error('粘贴内容的双引号未闭合');
  if (value !== '' || row.length || !rows.length || !/[\r\n]$/.test(text)) { pushCell(); rows.push(row); }
  return rows;
}

export function pasteChanges(text: string, start: Position) {
  const rows = parseTsv(text);
  const width = Math.max(...rows.map(r => r.length));
  if (rows.length * width > MAX_PASTE_CELLS) throw new Error('粘贴范围超过单次上限');
  if (start.row + rows.length > MAX_ROWS || start.col + width > MAX_COLUMNS) throw new Error('粘贴范围超出表格边界');
  return { changes: rows.flatMap((row, r) => Array.from({ length: width }, (_, c) => ({ row: start.row + r, col: start.col + c, input: row[c] ?? '' }))), end: { row: start.row + rows.length - 1, col: start.col + width - 1 } };
}

export function copyTsv(sheet: Sheet, selection: Selection): string {
  const b = bounds(selection);
  if ((b.bottom - b.top + 1) * (b.right - b.left + 1) > MAX_PASTE_CELLS) throw new Error('复制范围超过 50,000 个单元格，请缩小选区');
  const quote = (s: string) => /[\t\r\n"]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
  return Array.from({ length: b.bottom - b.top + 1 }, (_, r) => Array.from({ length: b.right - b.left + 1 }, (_, c) => quote(sheet.cells[cellKey(b.top + r, b.left + c)]?.input ?? '')).join('\t')).join('\r\n');
}

export function numericValue(input: string): number | null {
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(input)) return null;
  if (/^[+-]?0\d/.test(input)) return null; // preserve identifiers with leading zeroes
  const value = Number(input); return Number.isFinite(value) ? value : null;
}
