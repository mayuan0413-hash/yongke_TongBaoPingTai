import assert from 'node:assert/strict';
import test from 'node:test';
import {
  evaluateCell,
  evaluateWorkbook,
  formatFormulaValue,
  parseFormula,
  rewriteFormulaOnAxisMutation,
  tokenize,
} from '../domain/formula/index.ts';
import { applyCommand, createProject } from '../domain/workbook/commands.ts';
import type { Project } from '../domain/workbook/types.ts';

let testId = 0;
function createTestWorkbook(): Project {
  let project = createProject('公式测试项目', `proj-${++testId}`, 'sheet-1');
  // Add a second sheet
  project = applyCommand(project, { type: 'addSheet', id: 'sheet-2', name: 'Sheet2' });
  // Add a sheet with special characters in name
  project = applyCommand(project, { type: 'addSheet', id: 'sheet-3', name: '小时级通报 (2)' });
  return project;
}

void test('词法与语法解析正确识别运算符、百分比与字面量', () => {
  const t1 = tokenize('=50%');
  assert.equal(t1.length, 2);
  assert.equal(t1[0].value, '50');
  assert.equal(t1[1].value, '%');

  const ast1 = parseFormula('=50%');
  assert.equal(ast1.kind, 'unary_op');

  const ast2 = parseFormula('=(1+2)*3');
  assert.equal(ast2.kind, 'binary_op');
  if (ast2.kind === 'binary_op') {
    assert.equal(ast2.op, '*');
    assert.equal(ast2.left.kind, 'binary_op');
  }

  const ast3 = parseFormula('="A" & "B"');
  assert.equal(ast3.kind, 'binary_op');
  if (ast3.kind === 'binary_op') {
    assert.equal(ast3.op, '&');
  }
});

void test('基础四则运算、乘方、百分比与字符串拼接求值', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 0, col: 0, input: '=50%' }, // A1
      { row: 0, col: 1, input: '=2^3' }, // B1
      { row: 0, col: 2, input: '="A"&"B"' }, // C1
      { row: 0, col: 3, input: '=10/0' }, // D1
      { row: 0, col: 4, input: '=(2+3)*4' }, // E1
    ],
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 0, 0).display, '0.5');
  assert.equal(wb.getCell('sheet-1', 0, 1).display, '8');
  assert.equal(wb.getCell('sheet-1', 0, 2).display, 'AB');
  assert.equal(wb.getCell('sheet-1', 0, 3).display, '#DIV/0!');
  assert.equal(wb.getCell('sheet-1', 0, 4).display, '20');
});

void test('单元格引用求值：A1=10, B1=20, C1==A1+B1 结果为 30', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 0, col: 0, input: '10' }, // A1
      { row: 0, col: 1, input: '20' }, // B1
      { row: 0, col: 2, input: '=A1+B1' }, // C1
      { row: 0, col: 3, input: '=$A$1+$B$1' }, // D1
    ],
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 0, 2).display, '30');
  assert.equal(wb.getCell('sheet-1', 0, 3).display, '30');
  // 验证用户输入的原公式不被重写覆盖
  assert.equal(project.sheets[0].cells['0:2'].input, '=A1+B1');
});

void test('惰性求值：IF 与 IFERROR 未选择分支不应触发计算错误', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 0, col: 0, input: '=IF(FALSE, 1/0, 100)' }, // A1
      { row: 0, col: 1, input: '=IF(TRUE, 42, 1/0)' }, // B1
      { row: 0, col: 2, input: '=IFERROR(10/0, "fallback")' }, // C1
      { row: 0, col: 3, input: '=IFERROR(88, "fallback")' }, // D1
    ],
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 0, 0).display, '100');
  assert.equal(wb.getCell('sheet-1', 0, 1).display, '42');
  assert.equal(wb.getCell('sheet-1', 0, 2).display, 'fallback');
  assert.equal(wb.getCell('sheet-1', 0, 3).display, '88');
});

void test('跨 Sheet 引用与特殊名称 Sheet 解析求值', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [{ row: 0, col: 0, input: '100' }], // Sheet1!A1 = 100
  });
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-2',
    changes: [{ row: 0, col: 0, input: '=Sheet1!A1*2' }], // Sheet2!A1 = Sheet1!A1*2
  });
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-3', // '小时级通报 (2)'
    changes: [{ row: 0, col: 0, input: '50' }],
  });
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [{ row: 1, col: 0, input: "='小时级通报 (2)'!A1 + 10" }], // Sheet1!A2
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-2', 0, 0).display, '200');
  assert.equal(wb.getCell('sheet-1', 1, 0).display, '60');
});

void test('区域与整列引用求值：SUM(A1:A3) 与 SUM(A:A)', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 0, col: 0, input: '1' }, // A1
      { row: 1, col: 0, input: '2' }, // A2
      { row: 2, col: 0, input: '3' }, // A3
      { row: 0, col: 1, input: '=SUM(A1:A3)' }, // B1
      { row: 1, col: 1, input: '=SUM(A:A)' }, // B2
    ],
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 0, 1).display, '6');
  assert.equal(wb.getCell('sheet-1', 1, 1).display, '6');
});

void test('循环引用检测：A1=B1, B1=A1 以及自引用 A1=A1+1 返回 #CIRCULAR! 且不发生死循环', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 0, col: 0, input: '=B1' }, // A1
      { row: 0, col: 1, input: '=A1' }, // B1
      { row: 1, col: 0, input: '=A2+1' }, // A2
    ],
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 0, 0).display, '#CIRCULAR!');
  assert.equal(wb.getCell('sheet-1', 0, 1).display, '#CIRCULAR!');
  assert.equal(wb.getCell('sheet-1', 1, 0).display, '#CIRCULAR!');
});

void test('错误处理：不存在的工作表返回 #REF!，未定义函数返回 #NAME?', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 0, col: 0, input: '=NoSuchSheet!A1' },
      { row: 0, col: 1, input: '=UNKNOWN_FUNCTION(1, 2)' },
      { row: 0, col: 2, input: '="abc" + 10' },
    ],
  });

  const wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 0, 0).display, '#REF!');
  assert.equal(wb.getCell('sheet-1', 0, 1).display, '#NAME?');
  assert.equal(wb.getCell('sheet-1', 0, 2).display, '#VALUE!');
});

void test('公式重写器：插入/删除行时绝对引用与相对引用均做结构调整，被删除区域转为 #REF!', () => {
  // 1. 插入行测试：A3 和 $A$3 在第 1 行之前插入一行后，均移动到第 4 行 (A4, $A$4)
  const f1 = rewriteFormulaOnAxisMutation(
    '=SUM(  A3 ,  $A$3 )',
    'Sheet1',
    { sheetName: 'Sheet1', axis: 'row', index: 0, count: 1, isInsert: true },
  );
  // 保持空格
  assert.equal(f1, '=SUM(  A4 ,  $A$4 )');

  // 2. 删除行测试：若第 3 行被删除 (index: 2, count: 1)，A3 变成 #REF!
  const f2 = rewriteFormulaOnAxisMutation(
    '=A3 + 10',
    'Sheet1',
    { sheetName: 'Sheet1', axis: 'row', index: 2, count: 1, isInsert: false },
  );
  assert.equal(f2, '=#REF! + 10');

  // 3. 删除行测试：若第 1 行被删除 (index: 0, count: 1)，A3 前移为 A2
  const f3 = rewriteFormulaOnAxisMutation(
    '=A3 + 10',
    'Sheet1',
    { sheetName: 'Sheet1', axis: 'row', index: 0, count: 1, isInsert: false },
  );
  assert.equal(f3, '=A2 + 10');

  // 4. 整列 A:A 引用不受行插入/删除影响
  const f4 = rewriteFormulaOnAxisMutation(
    '=SUM(A:A)',
    'Sheet1',
    { sheetName: 'Sheet1', axis: 'row', index: 0, count: 1, isInsert: true },
  );
  assert.equal(f4, '=SUM(A:A)');

  // 5. 跨表引用只有目标表被修改时才移动
  const f5 = rewriteFormulaOnAxisMutation(
    "=Sheet2!A3 + '小时级通报 (2)'!A3",
    'Sheet1',
    { sheetName: 'Sheet2', axis: 'row', index: 0, count: 1, isInsert: true },
  );
  assert.equal(f5, "=Sheet2!A4 + '小时级通报 (2)'!A3");
});

void test('工作簿命令联动：执行 insertAxis / deleteAxis 后公式自动平移或变 #REF!', () => {
  let project = createTestWorkbook();
  project = applyCommand(project, {
    type: 'setCells',
    sheetId: 'sheet-1',
    changes: [
      { row: 2, col: 0, input: '100' }, // A3
      { row: 0, col: 1, input: '=A3*2' }, // B1 = A3*2
    ],
  });

  // 在第 1 行前插入 1 行，原 B1(0:1) 移到 B2(1:1)，原 A3 移到 A4，公式应变为 =A4*2
  project = applyCommand(project, {
    type: 'insertAxis',
    sheetId: 'sheet-1',
    axis: 'row',
    index: 0,
    count: 1,
  });
  assert.equal(project.sheets[0].cells['1:1']?.input, '=A4*2');
  let wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 1, 1).display, '200');

  // 删除当前 A4 所在的第 4 行 (index 3)
  project = applyCommand(project, {
    type: 'deleteAxis',
    sheetId: 'sheet-1',
    axis: 'row',
    index: 3,
    count: 1,
  });
  assert.equal(project.sheets[0].cells['1:1']?.input, '=#REF!*2');
  wb = evaluateWorkbook(project);
  assert.equal(wb.getCell('sheet-1', 1, 1).display, '#REF!');
});

void test('四种引用模式 ($A$1, $A1, A$1, A1) 在行列结构插入时全部正确平移', () => {
  // 插入行：A2, $A$2, A$2, $A2 全部平移至第 3 行 (A3, $A$3, A$3, $A3)
  const rowRewritten = rewriteFormulaOnAxisMutation(
    '=A2 + $A$2 + A$2 + $A2',
    'Sheet1',
    { sheetName: 'Sheet1', axis: 'row', index: 0, count: 1, isInsert: true },
  );
  assert.equal(rowRewritten, '=A3 + $A$3 + A$3 + $A3');

  // 插入列：B1, $B$1, B$1, $B1 全部平移至 C 列 (C1, $C$1, C$1, $C1)
  const colRewritten = rewriteFormulaOnAxisMutation(
    '=B1 + $B$1 + B$1 + $B1',
    'Sheet1',
    { sheetName: 'Sheet1', axis: 'column', index: 0, count: 1, isInsert: true },
  );
  assert.equal(colRewritten, '=C1 + $C$1 + C$1 + $C1');
});

