import { fromKey } from '../workbook/address.ts';
import type { Sheet } from '../workbook/types.ts';
import type {
  CalculationContext,
  CellRefNode,
  FormulaErrorInfo,
  FormulaValue,
  RangeRefNode,
  ResolvedRange,
} from './types.ts';
import { makeBlank, makeError } from './values.ts';

export function findSheetByName(sheets: Sheet[], name: string): Sheet | undefined {
  const norm = name.trim().toLowerCase();
  return sheets.find(s => s.name.toLowerCase() === norm);
}

export function getSheetUsedRange(sheet: Sheet): { startRow: number; endRow: number; startCol: number; endCol: number } {
  let maxRow = -1;
  let maxCol = -1;
  for (const key of Object.keys(sheet.cells)) {
    const pos = fromKey(key);
    maxRow = Math.max(maxRow, pos.row);
    maxCol = Math.max(maxCol, pos.col);
  }
  return {
    startRow: 0,
    endRow: Math.max(0, maxRow),
    startCol: 0,
    endCol: Math.max(0, maxCol),
  };
}

export function resolveCellReference(
  ref: CellRefNode,
  context: CalculationContext,
): { sheetId: string; row: number; col: number } | FormulaErrorInfo {
  let sheet: Sheet | undefined;
  if (ref.sheetName) {
    sheet = findSheetByName(context.project.sheets, ref.sheetName);
    if (!sheet) {
      return {
        type: '#REF!',
        message: `工作表 "${ref.sheetName}" 不存在`,
      };
    }
  } else {
    sheet = context.project.sheets.find(s => s.id === context.currentSheetId);
    if (!sheet) {
      return {
        type: '#REF!',
        message: '当前工作表不存在',
      };
    }
  }

  return {
    sheetId: sheet.id,
    row: ref.row,
    col: ref.col,
  };
}

export function resolveRangeReference(
  ref: RangeRefNode,
  context: CalculationContext,
): ResolvedRange | FormulaErrorInfo {
  let sheet: Sheet | undefined;
  if (ref.sheetName) {
    sheet = findSheetByName(context.project.sheets, ref.sheetName);
    if (!sheet) {
      return {
        type: '#REF!',
        message: `工作表 "${ref.sheetName}" 不存在`,
      };
    }
  } else {
    sheet = context.project.sheets.find(s => s.id === context.currentSheetId);
    if (!sheet) {
      return {
        type: '#REF!',
        message: '当前工作表不存在',
      };
    }
  }

  const sheetId = sheet.id;
  const sheetName = sheet.name;
  const startCol = ref.startCol;
  const endCol = ref.endCol;

  let startRow: number;
  let endRow: number;

  if (ref.isFullColumn) {
    // Whole-column reference A:A or E:H
    // Must be bounded by usedRange to avoid scanning theoretical max rows
    const used = getSheetUsedRange(sheet);
    startRow = 0;
    endRow = used.endRow;
  } else {
    startRow = ref.startRow ?? 0;
    endRow = ref.endRow ?? 0;
  }

  const rowCount = Math.max(0, endRow - startRow + 1);
  const columnCount = Math.max(0, endCol - startCol + 1);

  const getCellValue = (r: number, c: number): FormulaValue => {
    if (r < startRow || r > endRow || c < startCol || c > endCol) {
      return makeError('#REF!', '单元格超出区域范围');
    }
    return context.evalCell(sheetId, r, c);
  };

  const getValues = (): FormulaValue[][] => {
    const rows: FormulaValue[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const rowVals: FormulaValue[] = [];
      for (let c = startCol; c <= endCol; c++) {
        rowVals.push(getCellValue(r, c));
      }
      rows.push(rowVals);
    }
    return rows;
  };

  const flatten = (): FormulaValue[] => {
    const list: FormulaValue[] = [];
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        list.push(getCellValue(r, c));
      }
    }
    return list;
  };

  return {
    sheetId,
    sheetName,
    startRow,
    endRow,
    startCol,
    endCol,
    isFullColumn: ref.isFullColumn,
    rowCount,
    columnCount,
    getCellValue,
    getValues,
    flatten,
  };
}
