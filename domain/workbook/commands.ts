import { cellKey, fromKey, inSelection } from './address.ts';
import { DEFAULT_COLUMNS, DEFAULT_ROWS, MAX_COLUMNS, MAX_PASTE_CELLS, MAX_ROWS, type Project, type Sheet, type WorkbookCommand } from './types.ts';
import { rewriteFormulaOnAxisMutation } from '../formula/rewriter.ts';

export class WorkbookError extends Error {}
function ensure(valid: unknown, message: string): asserts valid { if (!valid) throw new WorkbookError(message); }

export function projectName(input: string): string {
  ensure(typeof input === 'string', '项目名称必须为文本');
  const name = input.trim();
  ensure(name.length > 0 && name.length <= 80, '项目名称需要 1–80 个字符');
  return name;
}

function sheetName(input: string, project: Project, exceptId?: string): string {
  ensure(typeof input === 'string', 'Sheet 名称必须为文本');
  const name = input.trim().normalize('NFC');
  ensure(name.length > 0 && name.length <= 31, 'Sheet 名称需要 1–31 个字符');
  ensure(!['\\', '/', '?', '*', '[', ']', ':'].some(char => name.includes(char)) && !Array.from(name).some(char => char.charCodeAt(0) < 32) && !name.startsWith("'") && !name.endsWith("'"), 'Sheet 名称不能包含 \\ / ? * [ ] :，也不能以单引号开头或结尾');
  ensure(!project.sheets.some(s => s.id !== exceptId && s.name.toLowerCase() === name.toLowerCase()), '已有同名 Sheet，请使用其他名称');
  return name;
}

function validId(id: string) { ensure(typeof id === 'string' && /^[a-zA-Z0-9-]{1,64}$/.test(id), '无效标识'); }
function integer(value: number, min: number, max: number, label: string) { ensure(Number.isInteger(value) && value >= min && value <= max, `${label}超出有效范围`); }

export function createSheet(id: string, name: string): Sheet {
  validId(id);
  return { id, name, rowCount: DEFAULT_ROWS, columnCount: DEFAULT_COLUMNS, cells: {}, rowHeights: {}, columnWidths: {}, dataSource: null };
}

export function createProject(name: string, id: string, sheetId: string, now = new Date().toISOString()): Project {
  validId(id);
  return { id, name: projectName(name), sheets: [createSheet(sheetId, 'Sheet1')], revision: 0, createdAt: now, updatedAt: now };
}

export function nextSheetName(project: Project, prefix = 'Sheet'): string {
  let suffix = 1;
  while (project.sheets.some(s => s.name.toLowerCase() === `${prefix}${suffix}`.toLowerCase())) suffix++;
  return `${prefix}${suffix}`;
}

export function copySheetName(project: Project, name: string): string {
  for (let i = 1; ; i++) {
    const suffix = i === 1 ? ' 副本' : ` 副本${i}`;
    const candidate = name.slice(0, 31 - suffix.length) + suffix;
    if (!project.sheets.some(s => s.name.toLowerCase() === candidate.toLowerCase())) return candidate;
  }
}

/** Pure commands are shared by the UI and API. No formulas or I/O live here. */
export function applyCommand(project: Project, command: WorkbookCommand, now = new Date().toISOString()): Project {
  ensure(command && typeof command === 'object' && typeof command.type === 'string', '操作格式无效');
  const next = { ...project, sheets: [...project.sheets], revision: project.revision + 1, updatedAt: now };
  if (command.type === 'renameProject') { next.name = projectName(command.name); return next; }
  if (command.type === 'addSheet') {
    validId(command.id);
    ensure(project.sheets.length < 100, '一个项目最多支持 100 个 Sheet');
    ensure(!project.sheets.some(s => s.id === command.id), 'Sheet 标识重复');
    next.sheets.push(createSheet(command.id, sheetName(command.name, project)));
    return next;
  }
  ensure('sheetId' in command, '未知操作');
  const index = project.sheets.findIndex(s => s.id === command.sheetId);
  ensure(index >= 0, 'Sheet 不存在');
  const source = project.sheets[index];
  if (source.dataSource && ['setCells', 'clearCells', 'insertAxis', 'deleteAxis'].includes(command.type)) throw new WorkbookError('数据库区域由查询维护。请刷新数据，或切换为手工数据后编辑');
  const sheet = { ...source };
  next.sheets[index] = sheet;
  switch (command.type) {
    case 'renameSheet': sheet.name = sheetName(command.name, project, source.id); break;
    case 'deleteSheet':
      ensure(project.sheets.length > 1, '项目至少需要保留一个 Sheet');
      next.sheets.splice(index, 1); break;
    case 'duplicateSheet': {
      validId(command.id);
      ensure(project.sheets.length < 100, '一个项目最多支持 100 个 Sheet');
      ensure(!project.sheets.some(s => s.id === command.id), 'Sheet 标识重复');
      const copy = structuredClone(source);
      copy.id = command.id; copy.name = sheetName(command.name, project);
      next.sheets.splice(index + 1, 0, copy); break;
    }
    case 'moveSheet':
      integer(command.toIndex, 0, project.sheets.length - 1, '排序位置');
      next.sheets.splice(index, 1); next.sheets.splice(command.toIndex, 0, source); break;
    case 'setCells': {
      ensure(Array.isArray(command.changes) && command.changes.length <= MAX_PASTE_CELLS, `一次最多编辑 ${MAX_PASTE_CELLS.toLocaleString()} 个单元格`);
      sheet.cells = { ...source.cells };
      for (const change of command.changes) {
        ensure(change && typeof change === 'object', '单元格数据无效');
        integer(change.row, 0, MAX_ROWS - 1, '行'); integer(change.col, 0, MAX_COLUMNS - 1, '列');
        ensure(typeof change.input === 'string' && change.input.length <= 32767, '单元格内容最多 32,767 个字符');
        const key = cellKey(change.row, change.col);
        if (change.input === '') delete sheet.cells[key]; else sheet.cells[key] = { input: change.input };
        sheet.rowCount = Math.max(sheet.rowCount, change.row + 1); sheet.columnCount = Math.max(sheet.columnCount, change.col + 1);
      }
      break;
    }
    case 'clearCells': {
      for (const position of [command.selection?.anchor, command.selection?.focus]) {
        ensure(position, '选区无效'); integer(position.row, 0, sheet.rowCount - 1, '行'); integer(position.col, 0, sheet.columnCount - 1, '列');
      }
      sheet.cells = Object.fromEntries(Object.entries(source.cells).filter(([key]) => !inSelection(fromKey(key), command.selection)));
      break;
    }
    case 'insertAxis':
    case 'deleteAxis': {
      ensure(command.axis === 'row' || command.axis === 'column', '无效的行列类型');
      const isRow = command.axis === 'row';
      const length = isRow ? source.rowCount : source.columnCount;
      const maximum = isRow ? MAX_ROWS : MAX_COLUMNS;
      const insert = command.type === 'insertAxis';
      integer(command.index, 0, insert ? length : length - 1, '位置'); integer(command.count, 1, maximum, '数量');
      ensure(insert ? length + command.count <= maximum : command.index + command.count <= length && length - command.count >= 1, '不能删除全部行列或超出边界');
      const shift = (value: number): number | null => insert ? (value >= command.index ? value + command.count : value) : value < command.index ? value : value < command.index + command.count ? null : value - command.count;
      const cells: Sheet['cells'] = {};
      for (const [key, cell] of Object.entries(source.cells)) {
        const pos = fromKey(key); const moved = shift(isRow ? pos.row : pos.col);
        if (moved !== null) cells[cellKey(isRow ? moved : pos.row, isRow ? pos.col : moved)] = cell;
      }
      sheet.cells = cells;
      const sizes = isRow ? source.rowHeights : source.columnWidths;
      const shifted: Record<number, number> = {};
      for (const [key, value] of Object.entries(sizes)) { const target = shift(Number(key)); if (target !== null) shifted[target] = value; }
      if (isRow) { sheet.rowCount = length + (insert ? command.count : -command.count); sheet.rowHeights = shifted; }
      else { sheet.columnCount = length + (insert ? command.count : -command.count); sheet.columnWidths = shifted; }

      // Rewrite formulas across all sheets targeting the mutated sheet
      const mutation = { sheetName: source.name, axis: command.axis, index: command.index, count: command.count, isInsert: insert };
      for (let sIdx = 0; sIdx < next.sheets.length; sIdx++) {
        const targetSheet = next.sheets[sIdx];
        let hasChanges = false;
        const rewrittenCells: Sheet['cells'] = { ...targetSheet.cells };
        for (const [k, c] of Object.entries(targetSheet.cells)) {
          if (c.input.startsWith('=')) {
            const rewritten = rewriteFormulaOnAxisMutation(c.input, targetSheet.name, mutation);
            if (rewritten !== c.input) {
              rewrittenCells[k] = { ...c, input: rewritten };
              hasChanges = true;
            }
          }
        }
        if (hasChanges) {
          next.sheets[sIdx] = { ...targetSheet, cells: rewrittenCells };
        }
      }
      break;
    }
    case 'resizeAxis':
      ensure(command.axis === 'row' || command.axis === 'column', '无效的行列类型');
      if (command.axis === 'row') {
        integer(command.index, 0, source.rowCount - 1, '行'); integer(command.size, 22, 400, '行高');
        sheet.rowHeights = { ...source.rowHeights, [command.index]: command.size };
      } else {
        integer(command.index, 0, source.columnCount - 1, '列'); integer(command.size, 48, 600, '列宽');
        sheet.columnWidths = { ...source.columnWidths, [command.index]: command.size };
      }
      break;
    case 'expandSheet':
      integer(command.rowCount, source.rowCount, MAX_ROWS, '行数'); integer(command.columnCount, source.columnCount, MAX_COLUMNS, '列数');
      sheet.rowCount = command.rowCount; sheet.columnCount = command.columnCount; break;
    default: throw new WorkbookError('未知操作');
  }
  return next;
}
