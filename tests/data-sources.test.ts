import assert from 'node:assert/strict';
import test from 'node:test';
import { blocksToCells, createCalculationInput } from '../domain/data-sources/change-set.ts';
import { compileQuery, validateQuery } from '../domain/data-sources/query.ts';
import { makeDataBlocks } from '../domain/data-sources/snapshot.ts';
import { applyCommand, createProject } from '../domain/workbook/commands.ts';
import type { DataSourceQuery, SourceColumn } from '../domain/workbook/types.ts';

const schema: SourceColumn[] = [
  { name: '编码', label: '门店编码', dataType: 'TEXT', nullable: false },
  { name: '销量', label: '销量', dataType: 'REAL', nullable: true },
  { name: '渠道', label: '渠道', dataType: 'TEXT', nullable: true },
];

void test('查询编译仅允许实时表结构中的字段，并绑定所有条件值', () => {
  const query: DataSourceQuery = {
    table: '终端销售明细', fields: ['编码', '销量'],
    filters: [{ field: '编码', operator: 'contains', value: "01%_' OR 1=1 --" }, { field: '销量', operator: 'gte', value: '10' }],
    orderBy: [{ field: '销量', direction: 'desc' }], rowLimit: 500,
  };
  const compiled = compileQuery(query, schema, 20);
  assert.match(compiled.sql, /^SELECT "编码","销量" FROM "终端销售明细" WHERE/);
  assert.ok(!compiled.sql.includes('OR 1=1'));
  assert.deepEqual(compiled.params, ["%01\\%\\_' OR 1=1 --%", '10', 21]);
  assert.throws(() => compileQuery({ ...query, fields: ['不存在'] }, schema), /字段不存在/);
});

void test('查询限制同时约束行数、单元格总量和条件数量', () => {
  assert.throws(() => validateQuery({ table: 't', fields: ['a'], filters: [], orderBy: [], rowLimit: 100001 }), /最多/);
  assert.throws(() => validateQuery({ table: 't', fields: Array.from({ length: 21 }, (_, i) => `f${i}`), filters: [], orderBy: [], rowLimit: 100000 }), /最多/);
});

void test('数据库快照保留数值类型，并转换为公式引擎可读单元格', () => {
  const blocks = makeDataBlocks([{ code: '0012', amount: 12.5 }, { code: '0013', amount: null }], ['code', 'amount']);
  const cells = blocksToCells(['门店编码', '销量'], blocks);
  assert.deepEqual(cells['1:0'], { input: '0012', sourceValue: '0012' });
  assert.deepEqual(cells['1:1'], { input: '12.5', sourceValue: 12.5 });
  assert.equal(cells['2:1'], undefined);

  const project = createProject('日报', 'project-1', 'sheet-1');
  project.sheets[0].cells = cells;
  const input = createCalculationInput(project);
  assert.equal(input.getCell('sheet-1', 1, 1)?.sourceValue, 12.5);
  assert.deepEqual(input.getUsedRange('sheet-1'), { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
  assert.equal(input.findSheetId('Sheet1'), 'sheet-1');
});

void test('数据库绑定 Sheet 拒绝手工写入和结构修改', () => {
  const project = createProject('日报', 'project-1', 'sheet-1');
  project.sheets[0].dataSource = {
    dataSourceId: 'source-1', dataSourceName: '业务库', connectionKey: 'business',
    query: { table: '销售', fields: ['销量'], filters: [], orderBy: [], rowLimit: 100 },
    columns: ['销量'], lastRefreshedAt: null, lastRowCount: 0, truncated: false,
  };
  assert.throws(() => applyCommand(project, { type: 'setCells', sheetId: 'sheet-1', changes: [{ row: 0, col: 0, input: '1' }] }), /数据库区域/);
  assert.throws(() => applyCommand(project, { type: 'deleteAxis', sheetId: 'sheet-1', axis: 'row', index: 0, count: 1 }), /数据库区域/);
});
