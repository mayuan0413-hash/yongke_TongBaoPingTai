import { indexToColName, type Token, tokenize } from './tokenizer.ts';

export interface AxisMutation {
  sheetName: string;
  axis: 'row' | 'column';
  index: number;
  count: number;
  isInsert: boolean;
}

/**
 * Rewrites a formula string when an axis (row or column) is inserted or deleted in target sheet.
 *
 * Rules:
 * 1. Structural change (insert/delete) shifts references regardless of '$' (absolute) flags.
 * 2. Only references pointing to the mutated sheet (explicit or implicit) are shifted.
 * 3. References in deleted areas become #REF!.
 * 4. Token-based replacement preserves all original whitespace, formatting, and casing.
 */
export function rewriteFormulaOnAxisMutation(
  formula: string,
  currentSheetName: string,
  mutation: AxisMutation,
): string {
  if (!formula.startsWith('=')) return formula;

  const content = formula.slice(1);
  const tokens = tokenize(content);
  if (tokens.length === 0) return formula;

  // We process tokens in reverse order of pos so replacements don't invalidate earlier positions
  let result = content;
  const isRow = mutation.axis === 'row';
  const mutSheet = mutation.sheetName.toLowerCase();

  for (let i = tokens.length - 1; i >= 0; i--) {
    const token = tokens[i];
    if (token.type !== 'CELL_REF' && token.type !== 'RANGE_REF') {
      continue;
    }

    // Determine target sheet of this reference
    const refSheet = (token.cellRef?.sheetName ?? token.rangeRef?.sheetName ?? currentSheetName).toLowerCase();
    if (refSheet !== mutSheet) {
      // This reference points to a different sheet, unaffected by this mutation
      continue;
    }

    let newRefText: string | null = null;

    if (token.type === 'CELL_REF' && token.cellRef) {
      const { sheetName, col, row, colAbsolute, rowAbsolute } = token.cellRef;
      const targetCoord = isRow ? row : col;
      const shifted = shiftCoordinate(targetCoord, mutation.index, mutation.count, mutation.isInsert);

      if (shifted === null) {
        // Deleted -> #REF!
        newRefText = '#REF!';
      } else if (shifted !== targetCoord) {
        const newCol = isRow ? col : shifted;
        const newRow = isRow ? shifted : row;
        newRefText = formatCellRef(sheetName, newCol, newRow, colAbsolute, rowAbsolute);
      }
    } else if (token.type === 'RANGE_REF' && token.rangeRef) {
      const {
        sheetName,
        startCol,
        startRow,
        endCol,
        endRow,
        startColAbs,
        startRowAbs,
        endColAbs,
        endRowAbs,
        isFullColumn,
      } = token.rangeRef;

      if (isFullColumn) {
        if (isRow) {
          // Whole-column A:A is unaffected by row insertions/deletions
          continue;
        } else {
          // Column mutation on A:A or E:H
          const shiftedStart = shiftCoordinate(startCol, mutation.index, mutation.count, mutation.isInsert);
          const shiftedEnd = shiftCoordinate(endCol, mutation.index, mutation.count, mutation.isInsert);

          if (shiftedStart === null && shiftedEnd === null) {
            newRefText = '#REF!';
          } else if (shiftedStart === null || shiftedEnd === null) {
            // One boundary was deleted
            newRefText = '#REF!';
          } else if (shiftedStart !== startCol || shiftedEnd !== endCol) {
            newRefText = formatRangeRef(
              sheetName,
              shiftedStart,
              null,
              shiftedEnd,
              null,
              startColAbs,
              false,
              endColAbs,
              false,
              true,
            );
          }
        }
      } else {
        // Rectangular range, e.g. A1:B10
        const startCoord = isRow ? (startRow ?? 0) : startCol;
        const endCoord = isRow ? (endRow ?? 0) : endCol;

        if (mutation.isInsert) {
          const shiftedStart = startCoord >= mutation.index ? startCoord + mutation.count : startCoord;
          const shiftedEnd = endCoord >= mutation.index ? endCoord + mutation.count : endCoord;

          if (shiftedStart !== startCoord || shiftedEnd !== endCoord) {
            newRefText = formatRangeRef(
              sheetName,
              isRow ? startCol : shiftedStart,
              isRow ? shiftedStart : startRow,
              isRow ? endCol : shiftedEnd,
              isRow ? shiftedEnd : endRow,
              startColAbs,
              startRowAbs,
              endColAbs,
              endRowAbs,
              false,
            );
          }
        } else {
          // Deletion
          const delStart = mutation.index;
          const delEnd = mutation.index + mutation.count - 1;

          if (startCoord >= delStart && endCoord <= delEnd) {
            // Entire range was deleted
            newRefText = '#REF!';
          } else if (startCoord >= delStart && startCoord <= delEnd) {
            // Start endpoint deleted
            newRefText = '#REF!';
          } else if (endCoord >= delStart && endCoord <= delEnd) {
            // End endpoint deleted
            newRefText = '#REF!';
          } else {
            // Shifting
            let s = startCoord;
            let e = endCoord;
            if (delStart < startCoord && delEnd < startCoord) {
              s -= mutation.count;
              e -= mutation.count;
            } else if (delStart > startCoord && delEnd < endCoord) {
              // Deleted slice inside the range -> range shrinks
              e -= mutation.count;
            }
            if (s !== startCoord || e !== endCoord) {
              newRefText = formatRangeRef(
                sheetName,
                isRow ? startCol : s,
                isRow ? s : startRow,
                isRow ? endCol : e,
                isRow ? e : endRow,
                startColAbs,
                startRowAbs,
                endColAbs,
                endRowAbs,
                false,
              );
            }
          }
        }
      }
    }

    if (newRefText !== null && newRefText !== token.value) {
      result = result.slice(0, token.pos) + newRefText + result.slice(token.pos + token.value.length);
    }
  }

  return '=' + result;
}

function shiftCoordinate(coord: number, index: number, count: number, isInsert: boolean): number | null {
  if (isInsert) {
    return coord >= index ? coord + count : coord;
  } else {
    // Delete
    if (coord >= index && coord < index + count) {
      return null; // Deleted
    }
    if (coord >= index + count) {
      return coord - count;
    }
    return coord;
  }
}

function formatSheetPrefix(sheetName?: string): string {
  if (!sheetName) return '';
  // If sheet name contains special characters or spaces, wrap in single quotes
  if (/^[A-Za-z0-9_]+$/.test(sheetName)) {
    return `${sheetName}!`;
  }
  return `'${sheetName.replace(/'/g, "''")}'!`;
}

function formatCellRef(
  sheetName: string | undefined,
  col: number,
  row: number,
  colAbs: boolean,
  rowAbs: boolean,
): string {
  const colPart = (colAbs ? '$' : '') + indexToColName(col);
  const rowPart = (rowAbs ? '$' : '') + (row + 1);
  return `${formatSheetPrefix(sheetName)}${colPart}${rowPart}`;
}

function formatRangeRef(
  sheetName: string | undefined,
  startCol: number,
  startRow: number | null,
  endCol: number,
  endRow: number | null,
  startColAbs: boolean,
  startRowAbs: boolean,
  endColAbs: boolean,
  endRowAbs: boolean,
  isFullColumn: boolean,
): string {
  const prefix = formatSheetPrefix(sheetName);
  if (isFullColumn) {
    const left = (startColAbs ? '$' : '') + indexToColName(startCol);
    const right = (endColAbs ? '$' : '') + indexToColName(endCol);
    return `${prefix}${left}:${right}`;
  }
  const left = (startColAbs ? '$' : '') + indexToColName(startCol) + (startRowAbs ? '$' : '') + ((startRow ?? 0) + 1);
  const right = (endColAbs ? '$' : '') + indexToColName(endCol) + (endRowAbs ? '$' : '') + ((endRow ?? 0) + 1);
  return `${prefix}${left}:${right}`;
}
