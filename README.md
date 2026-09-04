# 通报平台

这是从零构建的通报平台。当前完成 **Milestone 3：公式 MVP**，已有独立的公式解析与求值引擎、强类型 FormulaValue 体系、惰性求值函数注册机制、跨 Sheet 与整列引用、循环引用检测，以及行列结构变更时的公式智能平移。

## 当前能力

- 创建、切换、重命名和删除通报项目
- 新建、删除、重命名、复制和拖拽排序 Sheet
- Excel 风格行号、列号和 A1 地址栏
- 单元格键盘编辑、双击编辑、复制、粘贴和拖选多单元格
- 插入与删除行列、拖拽或精确设置行高列宽
- 行列插入/删除时公式引用智能平移（支持 $A$1, $A1, A$1, A1 四种模式），被删区域置为 #REF!
- 横纵双向虚拟渲染，大范围空白网格不会生成单元格记录
- D1 / SQLite 持久化，单元格按非空行稀疏保存
- 顺序自动保存和修改序号检查，防止不同页面静默覆盖数据
- 服务端数据源配置；数据库凭据和连接信息不会下发浏览器
- 浏览业务表结构，按字段、条件、排序和行数上限构建参数化查询
- 查询预览、绑定 Sheet、手工刷新以及刷新行数和截断状态提示
- 数据快照分块保存，刷新失败时保留上一版完整数据
- 数据库 Sheet 只读；可把当前快照转换成普通手工数据继续编辑
- 独立的公式计算层（`domain/formula/`），支持以 `=` 开头的公式输入与解析
- 强类型公式值体系：`number`、`string`、`boolean`、`blank`、`error`、`range`
- 基础运算符：`+`、`-`、`*`、`/`、`%`、`^`、`&` 及比较运算符，严格遵循 Excel 优先级
- 引用系统：普通单元格（A1, AA100）、矩形区域（A1:B10）、整列引用（A:A, E:H，按有效 usedRange 截断）、跨 Sheet 引用及单引号特殊 Sheet 名称（如 `'小时级通报 (2)'!A1`）
- 惰性求值与函数注册体系：内置支持 `IF`、`IFERROR`（未选择分支不求值，避免意外除零报错）、`SUM`、`COUNT`、`AVERAGE`、`INT`
- 完善的公式错误体系：`#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#CIRCULAR!`（死循环环路检测）
- 界面公式与计算值分离：双击/公式栏编辑原公式，单元格与状态栏渲染计算值及错误高亮，且渲染过程不重复重算工作簿

## 本地运行

需要 Node.js 22.13 或更高版本。

```bash
npm install
npm run db:migrate
npm run source:import -- /path/to/xlsx-or-directory
npm run dev
```

开发地址通常为 <http://localhost:3000>。

## 验证

```bash
npm run typecheck
npm test
npm run build
```

应用数据库模型在 `db/schema.ts`，迁移文件在 `drizzle/`。本地业务数据库由 `wrangler.source.json` 中的 `BUSINESS_DB` 独立管理；部署时也应把该绑定指向只读或最小权限的业务数据副本。`scripts/xlsx_to_sql.py` 可把样表首行转成字段标签，其余非空行导入业务库，供本地联调。

工作簿纯数据模型与命令位于 `domain/workbook/`，数据查询、快照和刷新变更协议位于 `domain/data-sources/`。所有查询字段都会先与实时表结构比对，表名和字段名只通过标识符转义进入 SQL，条件值全部使用绑定参数。

## 公式系统与 Milestone 4 接口

- `domain/formula/` 包含独立的公式解析、引用解析、函数注册与求值计算层。
- `FormulaValue` 强类型体系（`number`、`string`、`boolean`、`blank`、`error`、`range`），其中 `range` 保留了完整的 `ResolvedRange`（工作表 ID、起始与结束坐标、`getCellValue`、`getValues`、`flatten`），为 Milestone 4 的 `SUMIFS`、`COUNTIFS`、`VLOOKUP`、`XLOOKUP` 提供零损耗多维区域访问支持。
- `defaultRegistry.register(name, (args, context) => FormulaValue)` 提供惰性求值函数签名，函数直接接收 AST 节点，自主控制参数求值时机（例如 `IF` 分支选择、`IFERROR` 异常捕获）。
- `rewriteFormulaOnAxisMutation(...)` 提供基于 Token 替换的高保真公式平移能力，在保留用户原始格式与空白的同时，对四种引用模式（$A$1, $A1, A$1, A1）实现行列插入移动与越界 `#REF!` 标记。
- `CalculationCache` 结合 `project.id` 与 `project.revision` 实现计算结果缓存，防止界面渲染时重复执行整表重算。

## 数据边界

- 行：最多 1,000,000
- 列：最多 16,384（XFD）
- 单个 Sheet：默认 1,000 行、52 列，粘贴可按需扩展
- 单次粘贴 / 复制：最多 50,000 个单元格
- 单元格输入：最多 32,767 个字符
- 单次数据查询：最多 100,000 行、2,000,000 个单元格
- 查询条件：最多 20 个；排序字段：最多 8 个

这些限制与 Excel 使用习惯接近，并防止一次浏览器请求占用过多内存。数据库不保存空白单元格。
