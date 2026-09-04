import type { Project } from '../workbook/types.ts';

export type FormulaErrorType =
  | '#REF!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#N/A'
  | '#NAME?'
  | '#NUM!'
  | '#NULL!'
  | '#CIRCULAR!'
  | '#ERROR!';

export interface FormulaErrorInfo {
  type: FormulaErrorType;
  message: string;
  details?: string;
}

export interface ResolvedRange {
  sheetId: string;
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  isFullColumn: boolean;
  rowCount: number;
  columnCount: number;
  getCellValue: (row: number, col: number) => FormulaValue;
  getValues: () => FormulaValue[][];
  flatten: () => FormulaValue[];
}

export type FormulaValue =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'blank' }
  | { type: 'error'; error: FormulaErrorInfo }
  | { type: 'range'; range: ResolvedRange };

export type ASTNode =
  | LiteralNode
  | CellRefNode
  | RangeRefNode
  | UnaryOpNode
  | BinaryOpNode
  | FunctionCallNode;

export interface LiteralNode {
  kind: 'literal';
  value: FormulaValue;
  raw?: string;
}

export interface CellRefNode {
  kind: 'cell_ref';
  sheetName?: string;
  col: number;
  row: number;
  colAbsolute: boolean;
  rowAbsolute: boolean;
  rawText: string;
}

export interface RangeRefNode {
  kind: 'range_ref';
  sheetName?: string;
  startCol: number;
  startRow: number | null;
  endCol: number;
  endRow: number | null;
  startColAbs: boolean;
  startRowAbs: boolean;
  endColAbs: boolean;
  endRowAbs: boolean;
  isFullColumn: boolean;
  rawText: string;
}

export interface UnaryOpNode {
  kind: 'unary_op';
  op: '+' | '-' | '%';
  argument: ASTNode;
}

export interface BinaryOpNode {
  kind: 'binary_op';
  op: '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '<=' | '>' | '>=';
  left: ASTNode;
  right: ASTNode;
}

export interface FunctionCallNode {
  kind: 'function_call';
  name: string;
  args: ASTNode[];
}

export interface CellEvaluationResult {
  value: FormulaValue;
  display: string;
  error?: FormulaErrorInfo;
}

export interface CalculationContext {
  project: Project;
  currentSheetId: string;
  currentSheetName: string;
  callStack: Set<string>;
  evalNode: (node: ASTNode) => FormulaValue;
  evalCell: (sheetId: string, row: number, col: number) => FormulaValue;
  resolveCell: (ref: CellRefNode) => { sheetId: string; row: number; col: number } | FormulaErrorInfo;
  resolveRange: (ref: RangeRefNode) => ResolvedRange | FormulaErrorInfo;
}

export interface SheetCalculationResult {
  sheetId: string;
  cells: Map<string, CellEvaluationResult>;
}

export interface WorkbookCalculationResult {
  revision: number;
  sheets: Map<string, SheetCalculationResult>;
  getCell: (sheetId: string, row: number, col: number) => CellEvaluationResult;
}
