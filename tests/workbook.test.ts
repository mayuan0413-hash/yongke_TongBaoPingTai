import assert from 'node:assert/strict';
import test from 'node:test';
import { address, cellKey, columnLabel, parseAddress } from '../domain/workbook/address.ts';
import { copyTsv, numericValue, parseTsv, pasteChanges } from '../domain/workbook/clipboard.ts';
import { applyCommand, copySheetName, createProject } from '../domain/workbook/commands.ts';

const fresh = () => createProject('每日销量通报', 'project-1', 'sheet-1', '2026-09-02T00:00:00.000Z');

void test('A1 地址支持多字母列和 Excel 最大边界', () => {
  assert.equal(columnLabel(26), 'AA');
  assert.deepEqual(parseAddress('$XFD$1000000'), { row: 999_999, col: 16_383 });
  assert.equal(address({ row: 99, col: 701 }), 'ZZ100');
});

void test('Excel TSV 正确解析引号、制表符、换行和尾随空值', () => {
  assert.deepEqual(parseTsv('"甲\t类"\t12\r\n"两\n行"\t'), [['甲\t类', '12'], ['两\n行', '']]);
  const pasted = pasteChanges('A\tB\n1\t2', { row: 4, col: 2 });
  assert.equal(pasted.changes.length, 4);
  assert.deepEqual(pasted.end, { row: 5, col: 3 });
});

void test('项目命令只保存非空单元格，并递增修改序号', () => {
  const project = applyCommand(fresh(), { type: 'setCells', sheetId: 'sheet-1', changes: [{ row: 1, col: 2, input: '42' }, { row: 2, col: 2, input: '' }] });
  assert.equal(project.sheets[0].cells[cellKey(1, 2)].input, '42');
  assert.equal(Object.keys(project.sheets[0].cells).length, 1);
  assert.equal(project.revision, 1);
});

void test('插入和删除行保持已有单元格与自定义行高对齐', () => {
  let project = applyCommand(fresh(), { type: 'setCells', sheetId: 'sheet-1', changes: [{ row: 3, col: 1, input: '渠道' }] });
  project = applyCommand(project, { type: 'resizeAxis', sheetId: 'sheet-1', axis: 'row', index: 3, size: 48 });
  project = applyCommand(project, { type: 'insertAxis', sheetId: 'sheet-1', axis: 'row', index: 2, count: 2 });
  assert.equal(project.sheets[0].cells[cellKey(5, 1)].input, '渠道');
  assert.equal(project.sheets[0].rowHeights[5], 48);
  project = applyCommand(project, { type: 'deleteAxis', sheetId: 'sheet-1', axis: 'row', index: 1, count: 3 });
  assert.equal(project.sheets[0].cells[cellKey(2, 1)].input, '渠道');
});

void test('插入列时智能平移已有公式引用', () => {
  let project = applyCommand(fresh(), { type: 'setCells', sheetId: 'sheet-1', changes: [{ row: 0, col: 0, input: '=B1' }] });
  project = applyCommand(project, { type: 'insertAxis', sheetId: 'sheet-1', axis: 'column', index: 0, count: 1 });
  // 原 A1(0:0) 移动到 B1(0:1)，其引用的 B1 平移为 C1
  assert.equal(project.sheets[0].cells['0:1']?.input, '=C1');
});

void test('复制 Sheet 使用独立数据并生成不冲突的名称', () => {
  let project = applyCommand(fresh(), { type: 'setCells', sheetId: 'sheet-1', changes: [{ row: 0, col: 0, input: '原值' }] });
  const name = copySheetName(project, project.sheets[0].name);
  project = applyCommand(project, { type: 'duplicateSheet', sheetId: 'sheet-1', id: 'sheet-2', name });
  project = applyCommand(project, { type: 'setCells', sheetId: 'sheet-2', changes: [{ row: 0, col: 0, input: '副本值' }] });
  assert.equal(project.sheets[0].cells['0:0'].input, '原值');
  assert.equal(project.sheets[1].cells['0:0'].input, '副本值');
  assert.equal(name, 'Sheet1 副本');
});

void test('复制选区保留标识文本和数字含义', () => {
  const project = applyCommand(fresh(), { type: 'setCells', sheetId: 'sheet-1', changes: [{ row: 0, col: 0, input: '0012' }, { row: 0, col: 1, input: '12.5' }] });
  assert.equal(copyTsv(project.sheets[0], { anchor: { row: 0, col: 0 }, focus: { row: 0, col: 1 } }), '0012\t12.5');
  assert.equal(numericValue('0012'), null);
  assert.equal(numericValue('12.5'), 12.5);
});
