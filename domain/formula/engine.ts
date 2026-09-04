import { cellKey, fromKey } from '../workbook/address.ts';
import { numericValue } from '../workbook/clipboard.ts';
import type { Project, Sheet } from '../workbook/types.ts';
import { evaluateNode } from './evaluator.ts';
import { parseFormula } from './parser.ts';
import { resolveCellReference, resolveRangeReference } from './resolver.ts';
import { type AxisMutation, rewriteFormulaOnAxisMutation } from './rewriter.ts';
import type {
  ASTNode,
  CalculationContext,
  CellEvaluationResult,
  FormulaErrorInfo,
  FormulaValue,
  SheetCalculationResult,
  WorkbookCalculationResult,
} from './types.ts';
import {
  formatFormulaValue,
  isError,
  makeBlank,
  makeBoolean,
  makeError,
  makeNumber,
  makeString,
} from './values.ts';

export class CalculationCache {
  private parsedAsts = new Map<string, ASTNode>();
  private evaluatedResults = new Map<string, CellEvaluationResult>();
  private currentProjectId = '';
  private currentRevision = -1;

  getAst(formula: string): ASTNode {
    let ast = this.parsedAsts.get(formula);
    if (!ast) {
      ast = parseFormula(formula);
      this.parsedAsts.set(formula, ast);
    }
    return ast;
  }

  getEvaluated(projectId: string, revision: number, sheetId: string, row: number, col: number): CellEvaluationResult | undefined {
    if (this.currentProjectId !== projectId || this.currentRevision !== revision) {
      this.evaluatedResults.clear();
      this.currentProjectId = projectId;
      this.currentRevision = revision;
      return undefined;
    }
    return this.evaluatedResults.get(`${sheetId}:${row}:${col}`);
  }

  setEvaluated(projectId: string, revision: number, sheetId: string, row: number, col: number, result: CellEvaluationResult): void {
    if (this.currentProjectId !== projectId || this.currentRevision !== revision) {
      this.evaluatedResults.clear();
      this.currentProjectId = projectId;
      this.currentRevision = revision;
    }
    this.evaluatedResults.set(`${sheetId}:${row}:${col}`, result);
  }

  invalidate(sheetId?: string): void {
    if (!sheetId) {
      this.evaluatedResults.clear();
    } else {
      for (const key of this.evaluatedResults.keys()) {
        if (key.startsWith(`${sheetId}:`)) {
          this.evaluatedResults.delete(key);
        }
      }
    }
  }
}

export const defaultCalculationCache = new CalculationCache();

export function evaluateCell(
  project: Project,
  sheetId: string,
  row: number,
  col: number,
  cache = defaultCalculationCache,
): CellEvaluationResult {
  const cached = cache.getEvaluated(project.id, project.revision, sheetId, row, col);
  if (cached) return cached;

  const sheet = project.sheets.find(s => s.id === sheetId);
  if (!sheet) {
    const res: CellEvaluationResult = {
      value: makeError('#REF!', '工作表不存在'),
      display: '#REF!',
      error: { type: '#REF!', message: '工作表不存在' },
    };
    cache.setEvaluated(project.id, project.revision, sheetId, row, col, res);
    return res;
  }

  const cell = sheet.cells[cellKey(row, col)];
  if (!cell || cell.input === '') {
    const res: CellEvaluationResult = {
      value: makeBlank(),
      display: '',
    };
    cache.setEvaluated(project.id, project.revision, sheetId, row, col, res);
    return res;
  }

  const callStack = new Set<string>();

  const context: CalculationContext = {
    project,
    currentSheetId: sheet.id,
    currentSheetName: sheet.name,
    callStack,
    evalNode: (node: ASTNode) => evaluateNode(node, context),
    evalCell: (targetSheetId: string, r: number, c: number) => {
      const targetSheet = project.sheets.find(s => s.id === targetSheetId);
      if (!targetSheet) {
        return makeError('#REF!', '引用的工作表不存在');
      }

      const key = `${targetSheetId}:${r}:${c}`;
      if (callStack.has(key)) {
        return makeError('#CIRCULAR!', `检测到循环引用: ${targetSheet.name}!${r + 1}:${c + 1}`);
      }

      const cCached = cache.getEvaluated(project.id, project.revision, targetSheetId, r, c);
      if (cCached) return cCached.value;

      const targetCell = targetSheet.cells[cellKey(r, c)];
      if (!targetCell || targetCell.input === '') {
        return makeBlank();
      }

      // If formula
      if (targetCell.input.startsWith('=')) {
        callStack.add(key);
        const subContext: CalculationContext = {
          ...context,
          currentSheetId: targetSheet.id,
          currentSheetName: targetSheet.name,
        };
        const ast = cache.getAst(targetCell.input);
        const val = evaluateNode(ast, subContext);
        callStack.delete(key);

        const subRes: CellEvaluationResult = {
          value: val,
          display: formatFormulaValue(val),
          error: isError(val) ? val.error : undefined,
        };
        cache.setEvaluated(project.id, project.revision, targetSheetId, r, c, subRes);
        return val;
      }

      // Literal cell
      const litVal = evaluateLiteralCell(targetCell.input, targetCell.sourceValue);
      const litRes: CellEvaluationResult = {
        value: litVal,
        display: formatFormulaValue(litVal),
        error: isError(litVal) ? litVal.error : undefined,
      };
      cache.setEvaluated(project.id, project.revision, targetSheetId, r, c, litRes);
      return litVal;
    },
    resolveCell: (ref) => resolveCellReference(ref, context),
    resolveRange: (ref) => resolveRangeReference(ref, context),
  };

  const key = `${sheetId}:${row}:${col}`;
  callStack.add(key);

  let val: FormulaValue;
  if (cell.input.startsWith('=')) {
    const ast = cache.getAst(cell.input);
    val = evaluateNode(ast, context);
  } else {
    val = evaluateLiteralCell(cell.input, cell.sourceValue);
  }

  callStack.delete(key);

  const res: CellEvaluationResult = {
    value: val,
    display: formatFormulaValue(val),
    error: isError(val) ? val.error : undefined,
  };
  cache.setEvaluated(project.id, project.revision, sheetId, row, col, res);
  return res;
}

export function evaluateWorkbook(
  project: Project,
  cache = defaultCalculationCache,
): WorkbookCalculationResult {
  const sheetResults = new Map<string, SheetCalculationResult>();

  for (const sheet of project.sheets) {
    const cellMap = new Map<string, CellEvaluationResult>();
    for (const key of Object.keys(sheet.cells)) {
      const pos = fromKey(key);
      const evaluated = evaluateCell(project, sheet.id, pos.row, pos.col, cache);
      cellMap.set(key, evaluated);
    }
    sheetResults.set(sheet.id, {
      sheetId: sheet.id,
      cells: cellMap,
    });
  }

  return {
    revision: project.revision,
    sheets: sheetResults,
    getCell: (sheetId: string, row: number, col: number): CellEvaluationResult => {
      const s = sheetResults.get(sheetId);
      const key = cellKey(row, col);
      if (s?.cells.has(key)) {
        return s.cells.get(key)!;
      }
      return evaluateCell(project, sheetId, row, col, cache);
    },
  };
}

function evaluateLiteralCell(input: string, sourceValue?: string | number | null): FormulaValue {
  if (sourceValue !== undefined && sourceValue !== null) {
    if (typeof sourceValue === 'number') {
      return makeNumber(sourceValue);
    }
    if (typeof sourceValue === 'string') {
      const num = Number(sourceValue);
      if (!Number.isNaN(num) && Number.isFinite(num) && !sourceValue.startsWith('0') && sourceValue.trim() !== '') {
        return makeNumber(num);
      }
      return makeString(sourceValue);
    }
  }

  if (input.startsWith("'")) {
    return makeString(input.slice(1));
  }

  const num = numericValue(input);
  if (num !== null) {
    return makeNumber(num);
  }

  const upper = input.trim().toUpperCase();
  if (upper === 'TRUE') return makeBoolean(true);
  if (upper === 'FALSE') return makeBoolean(false);

  return makeString(input);
}

export { parseFormula } from './parser.ts';
export { rewriteFormulaOnAxisMutation } from './rewriter.ts';
