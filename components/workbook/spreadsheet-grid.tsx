'use client';
import { useEffect, useMemo, useRef, useState } from 'react';
import { useVirtualizer } from '@tanstack/react-virtual';
import { bounds, cellKey, columnLabel, inSelection, selectionLabel } from '@/domain/workbook/address';
import { copyTsv, numericValue, pasteChanges } from '@/domain/workbook/clipboard';
import { DEFAULT_COLUMN_WIDTH, DEFAULT_ROW_HEIGHT, type Sheet, type Selection, type WorkbookCommand } from '@/domain/workbook/types';

type Menu = { x: number; y: number; axis: 'row'|'column'; index: number } | null;
interface Props { sheet: Sheet; onCommand: (command: WorkbookCommand) => void; report: (message: string, error?: boolean) => void; readOnly?: boolean }

export function SpreadsheetGrid({ sheet, onCommand, report, readOnly = false }: Props) {
 const scrollRef = useRef<HTMLDivElement>(null);
 const dragging = useRef(false);
 const [selection, setSelection] = useState<Selection>({anchor:{row:0,col:0},focus:{row:0,col:0}});
 const [draft, setDraft] = useState('');
 const [editing, setEditing] = useState<'cell'|'formula'|null>(null);
 const [menu, setMenu] = useState<Menu>(null);
 const [scroll, setScroll] = useState({ x: 0, y: 0 });
 const explainReadOnly = () => report('当前 Sheet 绑定了数据库，需点击上方“转为手工数据”后才能编辑', true);
 const b = bounds(selection);
 const currentInput = sheet.cells[cellKey(selection.focus.row, selection.focus.col)]?.input ?? '';
 useEffect(() => { if (!editing) setDraft(currentInput); }, [currentInput, editing, sheet.id]);
 useEffect(() => { setSelection({anchor:{row:0,col:0},focus:{row:0,col:0}}); setEditing(null); }, [sheet.id]);
 useEffect(() => {
   const close = () => { dragging.current = false; setMenu(null); };
   window.addEventListener('pointerup', close); window.addEventListener('blur', close);
   return () => { window.removeEventListener('pointerup', close); window.removeEventListener('blur', close); };
 }, []);
 const rows = useVirtualizer({ count: sheet.rowCount, getScrollElement: () => scrollRef.current, estimateSize: i => sheet.rowHeights[i] ?? DEFAULT_ROW_HEIGHT, overscan: 8 });
 const columns = useVirtualizer({ horizontal: true, count: sheet.columnCount, getScrollElement: () => scrollRef.current, estimateSize: i => sheet.columnWidths[i] ?? DEFAULT_COLUMN_WIDTH, overscan: 3 });
 const virtualRows = rows.getVirtualItems(), virtualCols = columns.getVirtualItems();
 const numeric = useMemo(() => {
   let count = 0, sum = 0;
   for (const [key, cell] of Object.entries(sheet.cells)) if (inSelection({row:Number(key.split(':')[0]),col:Number(key.split(':')[1])}, selection)) { const value=numericValue(cell.input); if(value!==null){count++;sum+=value;} }
   return { count, sum };
 }, [sheet.cells, selection]);
 const choose = (row:number,col:number,event?:React.PointerEvent) => {
   const pos={row,col}; setEditing(null);
   setSelection(previous => event?.shiftKey ? {...previous,focus:pos}:{anchor:pos,focus:pos});
   dragging.current=true; scrollRef.current?.focus();
 };
 const commit = (value=draft) => {
   setEditing(null);
   if (readOnly) return;
   if(value!==currentInput) onCommand({type:'setCells',sheetId:sheet.id,changes:[{...selection.focus,input:value}]});
 };
 const move = (dr:number,dc:number,extend=false) => {
   const pos={row:Math.max(0,Math.min(sheet.rowCount-1,selection.focus.row+dr)),col:Math.max(0,Math.min(sheet.columnCount-1,selection.focus.col+dc))};
   setSelection(previous=>extend?{...previous,focus:pos}:{anchor:pos,focus:pos}); rows.scrollToIndex(pos.row); columns.scrollToIndex(pos.col);
 };
 const onKeyDown=(event:React.KeyboardEvent)=>{
   if(readOnly){
     if(event.key==='ArrowUp'||event.key==='ArrowDown'||event.key==='ArrowLeft'||event.key==='ArrowRight'){
       event.preventDefault();move(event.key==='ArrowUp'?-1:event.key==='ArrowDown'?1:0,event.key==='ArrowLeft'?-1:event.key==='ArrowRight'?1:0,event.shiftKey);
     }
     else if(event.key==='Enter'||event.key==='F2'||event.key==='Delete'||event.key==='Backspace'||(!event.metaKey&&!event.ctrlKey&&!event.altKey&&event.key.length===1)) explainReadOnly();
     return;
   }
   if(editing){ if(event.key==='Enter'){event.preventDefault();commit();move(1,0);} if(event.key==='Escape'){setEditing(null);setDraft(currentInput);} return; }
   const key=event.key;
   if(key==='ArrowUp'||key==='ArrowDown'||key==='ArrowLeft'||key==='ArrowRight'){event.preventDefault();move(key==='ArrowUp'?-1:key==='ArrowDown'?1:0,key==='ArrowLeft'?-1:key==='ArrowRight'?1:0,event.shiftKey);}
   else if(key==='Enter'||key==='F2'){event.preventDefault();setDraft(currentInput);setEditing('cell');}
   else if(key==='Delete'||key==='Backspace'){event.preventDefault();onCommand({type:'clearCells',sheetId:sheet.id,selection});}
   else if(!event.metaKey&&!event.ctrlKey&&!event.altKey&&key.length===1){setDraft(key);setEditing('cell');}
 };
 const copy=(event:React.ClipboardEvent)=>{event.preventDefault();try{event.clipboardData.setData('text/plain',copyTsv(sheet,selection));report(`已复制 ${selectionLabel(selection)}`);}catch(e){report((e as Error).message,true);}};
 const paste=(event:React.ClipboardEvent)=>{event.preventDefault();if(readOnly){report('数据库 Sheet 由刷新数据维护；切换为手工数据后可编辑',true);return;}try{const result=pasteChanges(event.clipboardData.getData('text/plain'),selection.focus);onCommand({type:'setCells',sheetId:sheet.id,changes:result.changes});setSelection({anchor:selection.focus,focus:result.end});report(`已粘贴 ${result.changes.length.toLocaleString()} 个单元格`);}catch(e){report((e as Error).message,true);}};
 const axisCommand=(type:'insertAxis'|'deleteAxis')=>{if(!menu)return;onCommand({type,sheetId:sheet.id,axis:menu.axis,index:menu.index,count:1});setMenu(null);};
 const resize=(axis:'row'|'column',index:number,start:number,startSize:number)=>{
   let latest=startSize; const min=axis==='row'?22:48,max=axis==='row'?400:600;
   const move=(e:PointerEvent)=>{latest=Math.round(Math.max(min,Math.min(max,startSize+(axis==='row'?e.clientY:e.clientX)-start)));};
   const up=()=>{window.removeEventListener('pointermove',move);window.removeEventListener('pointerup',up);if(!readOnly)onCommand({type:'resizeAxis',sheetId:sheet.id,axis,index,size:latest});};
   window.addEventListener('pointermove',move);window.addEventListener('pointerup',up,{once:true});
 };
 return <div className="grid-area">
   <div className="formula-bar">
    <div className="name-box">{selectionLabel(selection)}</div><span className="fx">fx</span>
    <input value={draft} readOnly={readOnly} aria-label="单元格内容" placeholder={readOnly?'数据库 Sheet 为只读；转为手工数据后可编辑':'输入内容；公式计算将在公式阶段启用'} onFocus={()=>readOnly?explainReadOnly():setEditing('formula')} onChange={e=>setDraft(e.target.value)} onKeyDown={e=>{if(e.key==='Enter'){e.preventDefault();commit();}}} onBlur={()=>editing==='formula'&&commit()} />
   </div>
   <div ref={scrollRef} className="spreadsheet" tabIndex={0} role="grid" aria-rowcount={sheet.rowCount} aria-colcount={sheet.columnCount} onScroll={e=>setScroll({x:e.currentTarget.scrollLeft,y:e.currentTarget.scrollTop})} onKeyDown={onKeyDown} onCopy={copy} onPaste={paste}>
    <div className="grid-canvas" style={{width:columns.getTotalSize()+46,height:rows.getTotalSize()+31}}>
     <div className="corner" />
     {virtualCols.map(col=><div key={col.key} className={`column-head ${col.index>=b.left&&col.index<=b.right?'axis-selected':''}`} style={{left:col.start+46,top:scroll.y,width:col.size}} onPointerDown={e=>{e.preventDefault();setSelection({anchor:{row:0,col:col.index},focus:{row:sheet.rowCount-1,col:col.index}});}} onContextMenu={e=>{e.preventDefault();if(!readOnly)setMenu({x:e.clientX,y:e.clientY,axis:'column',index:col.index});}}>{columnLabel(col.index)}{!readOnly&&<button className="resize-col" aria-label={`调整 ${columnLabel(col.index)} 列宽`} onPointerDown={e=>{e.stopPropagation();resize('column',col.index,e.clientX,col.size);}}/>}</div>)}
     {virtualRows.map(row=><div key={row.key} className={`row-head ${row.index>=b.top&&row.index<=b.bottom?'axis-selected':''}`} style={{top:row.start+31,left:scroll.x,height:row.size}} onPointerDown={e=>{e.preventDefault();setSelection({anchor:{row:row.index,col:0},focus:{row:row.index,col:sheet.columnCount-1}});}} onContextMenu={e=>{e.preventDefault();if(!readOnly)setMenu({x:e.clientX,y:e.clientY,axis:'row',index:row.index});}}>{row.index+1}{!readOnly&&<button className="resize-row" aria-label={`调整第 ${row.index+1} 行高度`} onPointerDown={e=>{e.stopPropagation();resize('row',row.index,e.clientY,row.size);}}/>}</div>)}
     {virtualRows.flatMap(row=>virtualCols.map(col=>{
       const selected=inSelection({row:row.index,col:col.index},selection), focused=row.index===selection.focus.row&&col.index===selection.focus.col;
       const input=sheet.cells[cellKey(row.index,col.index)]?.input??'';
       return <div key={row.key+':'+col.key} role="gridcell" tabIndex={-1} aria-selected={selected} className={`grid-cell ${selected?'selected':''} ${focused?'focused':''}`} style={{top:row.start+31,left:col.start+46,width:col.size,height:row.size}} onPointerDown={e=>choose(row.index,col.index,e)} onPointerEnter={()=>{if(dragging.current)setSelection(previous=>({...previous,focus:{row:row.index,col:col.index}}));}} onDoubleClick={()=>{if(readOnly)explainReadOnly();else{setDraft(input);setEditing('cell');}}} title={readOnly?'数据库 Sheet 为只读；转为手工数据后可编辑':input.startsWith('=')?'公式计算将在 Milestone 3 启用':''}>
        {focused&&editing==='cell'&&!readOnly?<input className="cell-editor" autoFocus value={draft} onChange={e=>setDraft(e.target.value)} onBlur={()=>commit()}/>:<span>{input.startsWith("'")?input.slice(1):input}</span>}
       </div>;
     }))}
    </div>
   </div>
   <footer className="grid-status"><span>{selectionLabel(selection)}</span><span>行 {sheet.rowCount.toLocaleString()}</span><span>列 {sheet.columnCount.toLocaleString()}</span>{numeric.count>0&&<><span className="status-spacer"/><span>计数 {numeric.count.toLocaleString()}</span><span>合计 {numeric.sum.toLocaleString()}</span></>}</footer>
   {menu&&<div className="context-menu" style={{left:Math.min(menu.x,window.innerWidth-190),top:Math.min(menu.y,window.innerHeight-135)}}><button onClick={()=>axisCommand('insertAxis')}>在{menu.axis==='row'?'上方插入行':'左侧插入列'}</button><button onClick={()=>axisCommand('deleteAxis')}>删除{menu.axis==='row'?'第 '+(menu.index+1)+' 行':columnLabel(menu.index)+' 列'}</button><button onClick={()=>{const size=Number(prompt(menu.axis==='row'?'行高（22–400）':'列宽（48–600）'));if(Number.isFinite(size))onCommand({type:'resizeAxis',sheetId:sheet.id,axis:menu.axis,index:menu.index,size});setMenu(null);}}>设置{menu.axis==='row'?'行高':'列宽'}…</button></div>}
  </div>;
}
