import { fromKey } from '../workbook/address.ts';
import type { Project, Sheet } from '../workbook/types.ts';

export interface DataRange { startRow: number; endRow: number; startColumn: number; endColumn: number }
export interface WorkbookDataChange {
  projectId: string;
  sheetId: string;
  revision: number;
  reason: 'data-refresh' | 'binding-change' | 'manual-edit';
  dirtyRanges: DataRange[];
  previousDataRows: number;
  dataRows: number;
  rowDelta: number;
}
export interface RefreshResult {
  revision: number;
  updatedAt: string;
  rowCount: number;
  columnCount: number;
  truncated: boolean;
  changeSet: WorkbookDataChange;
}

/** M3 can consume this read-only adapter without depending on React or D1. */
export function createCalculationInput(project: Project) {
  const sheets = new Map(project.sheets.map(s => [s.id, s]));
  const extents = new Map<string, DataRange | null>();
  return {
    revision: project.revision,
    sheetIds: [...sheets.keys()],
    findSheetId: (name: string) => project.sheets.find(s => s.name === name)?.id,
    getCell: (sheetId: string, row: number, column: number) => sheets.get(sheetId)?.cells[`${row}:${column}`] ?? null,
    getUsedRange(sheetId: string): DataRange | null {
      if (extents.has(sheetId)) return extents.get(sheetId)!;
      const sheet = sheets.get(sheetId); if (!sheet) return null;
      let maxRow = -1, maxCol = -1;
      for (const key of Object.keys(sheet.cells)) { const p = fromKey(key); maxRow = Math.max(maxRow, p.row); maxCol = Math.max(maxCol, p.col); }
      const extent = maxRow < 0 ? null : { startRow: 0, endRow: maxRow, startColumn: 0, endColumn: maxCol };
      extents.set(sheetId, extent); return extent;
    },
  };
}

export function blocksToCells(columns: string[], blocks: { rows: string; block_index: number }[]): Sheet['cells'] {
  const cells: Sheet['cells'] = {};
  columns.forEach((name, col) => { cells[`0:${col}`] = { input: name }; });
  for (const block of blocks) {
    const parsed = JSON.parse(block.rows) as { startRow: number; values: (string | number | null)[][] };
    parsed.values.forEach((row, i) => row.forEach((value, col) => {
      if (value !== null && value !== '') cells[`${parsed.startRow + i}:${col}`] = { input: String(value), sourceValue: value };
    }));
  }
  return cells;
}
