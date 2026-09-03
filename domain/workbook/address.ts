import { MAX_COLUMNS, MAX_ROWS, type Position, type Selection } from './types.ts';

export function columnLabel(index: number): string {
  if (!Number.isInteger(index) || index < 0 || index >= MAX_COLUMNS) throw new Error('列超出有效范围');
  let result = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
  return result;
}

export function address(position: Position): string { return `${columnLabel(position.col)}${position.row + 1}`; }

export function parseAddress(input: string): Position {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(input.trim());
  if (!match) throw new Error('请输入有效单元格地址，例如 A1 或 AA100');
  const col = Array.from(match[1].toUpperCase()).reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  if (row >= MAX_ROWS || col >= MAX_COLUMNS) throw new Error('单元格地址超出支持范围');
  return { row, col };
}

export const cellKey = (row: number, col: number): string => `${row}:${col}`;
export function fromKey(key: string): Position { const [row, col] = key.split(':').map(Number); return { row, col }; }

export function bounds(selection: Selection) {
  return { top: Math.min(selection.anchor.row, selection.focus.row), bottom: Math.max(selection.anchor.row, selection.focus.row), left: Math.min(selection.anchor.col, selection.focus.col), right: Math.max(selection.anchor.col, selection.focus.col) };
}

export function selectionLabel(selection: Selection): string {
  const { top, bottom, left, right } = bounds(selection);
  return top === bottom && left === right ? address({ row: top, col: left }) : `${address({ row: top, col: left })}:${address({ row: bottom, col: right })}`;
}

export function inSelection(position: Position, selection: Selection): boolean {
  const b = bounds(selection);
  return position.row >= b.top && position.row <= b.bottom && position.col >= b.left && position.col <= b.right;
}
