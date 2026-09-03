'use client';
import { useEffect, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, Copy, Database, Ellipsis, FileSpreadsheet, FolderOpen, Loader2, Pencil, Plus, RefreshCw, Table2, Trash2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Input } from '@/components/ui/input';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from '@/components/ui/dropdown-menu';
import { applyCommand, copySheetName, nextSheetName } from '@/domain/workbook/commands';
import type { Project, ProjectSummary, WorkbookCommand } from '@/domain/workbook/types';
import { api } from '@/lib/api';
import { SpreadsheetGrid } from './spreadsheet-grid';
import { DataSourcePanel } from './data-source-panel';

type Modal = {type:'project'|'renameProject'|'sheet'|'renameSheet';title:string;value:string}|null;
export function Workspace() {
 const [projects,setProjects]=useState<ProjectSummary[]>([]),[project,setProjectState]=useState<Project|null>(null);
 const [activeSheet,setActiveSheet]=useState(''),[loading,setLoading]=useState(true),[modal,setModal]=useState<Modal>(null);
 const [notice,setNotice]=useState<{message:string;error?:boolean}|null>(null);
 const [sourcePanel,setSourcePanel]=useState(false),[refreshing,setRefreshing]=useState(false);
 const projectRef=useRef<Project|null>(null), queue=useRef(Promise.resolve());
 const setProject=(p:Project|null)=>{projectRef.current=p;setProjectState(p);};
 const report=(message:string,error=false)=>{setNotice({message,error});window.setTimeout(()=>setNotice(current=>current?.message===message?null:current),2800);};
 const refreshList=()=>api.list().then(setProjects);
 const openProject=async(id:string)=>{setLoading(true);try{const p=await api.get(id);setProject(p);setActiveSheet(p.sheets[0]?.id??'');}catch(e){report((e as Error).message,true);}finally{setLoading(false);}};
 const reloadCurrent=async()=>{const current=projectRef.current;if(!current)return;const loaded=await api.get(current.id);setProject(loaded);setProjects(list=>list.map(item=>item.id===loaded.id?{...item,name:loaded.name,sheetCount:loaded.sheets.length,updatedAt:loaded.updatedAt}:item));};
 const sourceMutation=async(work:(current:Project)=>Promise<unknown>)=>{await queue.current;const current=projectRef.current;if(!current)return;await work(current);await reloadCurrent();};
 useEffect(()=>{api.list().then(async list=>{setProjects(list);if(list[0])await openProject(list[0].id);else setLoading(false);}).catch(e=>{setLoading(false);report((e as Error).message,true);});},[]);
 const command=(cmd:WorkbookCommand)=>{
   const before=projectRef.current;if(!before)return;
   let after:Project;try{after=applyCommand(before,cmd);}catch(e){report((e as Error).message,true);return;}
   setProject(after);
   queue.current=queue.current.catch(()=>{}).then(async()=>{try{await api.command(before.id,before.revision,cmd);setProjects(list=>list.map(item=>item.id===after.id?{...item,name:after.name,sheetCount:after.sheets.length,updatedAt:after.updatedAt}:item));report('已保存');}catch(e){report((e as Error).message,true);}});
 };
 const submit=async(event:React.SyntheticEvent<HTMLFormElement>)=>{event.preventDefault();if(!modal)return;const value=modal.value.trim();if(!value)return;
   try{if(modal.type==='project'){const p=await api.create(value);setProject(p);setActiveSheet(p.sheets[0].id);await refreshList();}
   else if(modal.type==='renameProject')command({type:'renameProject',name:value});
   else if(modal.type==='sheet'){const id=crypto.randomUUID();command({type:'addSheet',id,name:value});setActiveSheet(id);}
   else command({type:'renameSheet',sheetId:activeSheet,name:value});setModal(null);}catch(e){report((e as Error).message,true);}
 };
 const sheet=project?.sheets.find(s=>s.id===activeSheet)??project?.sheets[0];
 const makeSheetEditable=()=>{if(!sheet?.dataSource)return;sourceMutation(p=>api.unbindSource(p.id,sheet.id,p.revision)).then(()=>report('已转为手工数据，现在可以编辑单元格')).catch(e=>report((e as Error).message,true));};
 const askNewProject=()=>setModal({type:'project',title:'新建通报项目',value:''});
 const deleteProject=async()=>{const p=projectRef.current;if(!p||!confirm(`确定删除“${p.name}”吗？此操作无法撤销。`))return;try{await api.delete(p.id,p.revision);setProject(null);setActiveSheet('');await refreshList();report('项目已删除');}catch(e){report((e as Error).message,true);}};
 const deleteSheet=()=>{if(!project||!sheet)return;if(project.sheets.length===1){report('项目至少需要保留一个 Sheet',true);return;}if(confirm(`确定删除“${sheet.name}”吗？`)){const oldIndex=project.sheets.findIndex(s=>s.id===sheet.id);const next=project.sheets[oldIndex+1]??project.sheets[oldIndex-1];command({type:'deleteSheet',sheetId:sheet.id});setActiveSheet(next.id);}};
 return <main className="workspace">
  <aside className="sidebar">
   <div className="brand"><span className="brand-mark"><Table2 size={21}/></span><div>通报平台<small>REPORT WORKSPACE</small></div></div>
   <div className="sidebar-label">工作空间</div><div className="nav-active"><FolderOpen size={17}/>通报项目<span>{String(projects.length).padStart(2,'0')}</span></div>
   <div className="sidebar-label projects-label">我的项目<button aria-label="新建项目" onClick={askNewProject}><Plus size={15}/></button></div>
   <div className="project-list">{projects.map(item=><button key={item.id} className={item.id===project?.id?'project-item active':'project-item'} onClick={()=>openProject(item.id)}><FileSpreadsheet size={15}/><span>{item.name}</span><small>{item.sheetCount}</small></button>)}</div>
   {!projects.length&&!loading&&<p className="sidebar-empty">新建项目，开始制作第一份通报。</p>}
   <div className="sidebar-foot"><span className="status-dot"/>数据源层<span className="milestone-tag">M2</span></div>
  </aside>
  <section className="main-surface">
   <header className="topbar"><div className="breadcrumb">工作空间 <span>/</span><strong>{project?.name??'通报项目'}</strong></div><span className="phase-label">第二阶段 · 数据源层</span></header>
   {loading?<div className="loading"><Loader2 className="spin"/>正在加载工作区</div>:project&&sheet?<><div className="workbook-bar"><div><div className="eyebrow">通报项目</div><div className="project-title"><h1>{project.name}</h1><button aria-label="修改项目名称" onClick={()=>setModal({type:'renameProject',title:'修改项目名称',value:project.name})}><Pencil size={14}/></button></div><p>{project.sheets.length} 个 Sheet · 数据自动保存</p></div><DropdownMenu><DropdownMenuTrigger render={<Button variant="outline" size="icon"/>}><Ellipsis/></DropdownMenuTrigger><DropdownMenuContent align="end"><DropdownMenuItem onClick={()=>setModal({type:'renameProject',title:'修改项目名称',value:project.name})}><Pencil/>修改项目名称</DropdownMenuItem><DropdownMenuItem variant="destructive" onClick={deleteProject}><Trash2/>删除项目</DropdownMenuItem></DropdownMenuContent></DropdownMenu></div>
    <div className="editor-shell">
     <div className="sheet-heading"><Table2 size={17}/><strong>{sheet.name}</strong><span>{sheet.dataSource?'数据库 · 只读':'手工数据 · 可编辑'}</span>{sheet.dataSource&&<small>{sheet.dataSource.lastRowCount.toLocaleString()} 行{sheet.dataSource.truncated?' · 已达查询上限':''}</small>}<div className="sheet-tools">{sheet.dataSource&&<><Button variant="outline" size="sm" onClick={makeSheetEditable}>转为手工数据</Button><Button variant="outline" size="sm" disabled={refreshing} onClick={()=>{setRefreshing(true);sourceMutation(p=>api.refreshSource(p.id,sheet.id,p.revision)).then(()=>report('数据已刷新')).catch(e=>report((e as Error).message,true)).finally(()=>setRefreshing(false));}}>{refreshing?<Loader2 className="spin"/>:<RefreshCw/>}刷新数据</Button></>}<Button variant="outline" size="sm" onClick={()=>setSourcePanel(true)}><Database/>数据源</Button></div></div>
     <SpreadsheetGrid sheet={sheet} onCommand={command} report={report} readOnly={!!sheet.dataSource}/>
     {sourcePanel&&<DataSourcePanel key={sheet.id} project={project} sheet={sheet} report={report} onClose={()=>setSourcePanel(false)} onBind={(dataSourceId,query)=>sourceMutation(p=>api.bindSource(p.id,sheet.id,p.revision,dataSourceId,query))} onRefresh={()=>sourceMutation(p=>api.refreshSource(p.id,sheet.id,p.revision))} onUnbind={()=>sourceMutation(p=>api.unbindSource(p.id,sheet.id,p.revision))}/>} 
     <div className="sheet-tabs"><button className="tab-scroll" aria-label="向左滚动 Sheet"><ChevronLeft/></button><div className="tabs-scroll">{project.sheets.map((s,index)=><div key={s.id} className={s.id===sheet.id?'sheet-tab active':'sheet-tab'} draggable onDragStart={e=>e.dataTransfer.setData('text/sheet-index',String(index))} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const from=Number(e.dataTransfer.getData('text/sheet-index'));if(Number.isInteger(from)&&from!==index)command({type:'moveSheet',sheetId:project.sheets[from].id,toIndex:index});}}><button onClick={()=>setActiveSheet(s.id)}><Table2 size={14}/>{s.name}</button>{s.id===sheet.id&&<DropdownMenu><DropdownMenuTrigger className="tab-menu" aria-label="Sheet 菜单"><Ellipsis size={14}/></DropdownMenuTrigger><DropdownMenuContent><DropdownMenuItem onClick={()=>setModal({type:'renameSheet',title:'修改 Sheet 名称',value:s.name})}><Pencil/>重命名</DropdownMenuItem><DropdownMenuItem onClick={()=>{const id=crypto.randomUUID();command({type:'duplicateSheet',sheetId:s.id,id,name:copySheetName(project,s.name)});setActiveSheet(id);}}><Copy/>复制 Sheet</DropdownMenuItem><DropdownMenuItem variant="destructive" onClick={deleteSheet}><Trash2/>删除 Sheet</DropdownMenuItem></DropdownMenuContent></DropdownMenu>}</div>)}</div><button className="tab-scroll" aria-label="向右滚动 Sheet"><ChevronRight/></button><button className="add-sheet" onClick={()=>setModal({type:'sheet',title:'新建 Sheet',value:nextSheetName(project)})}><Plus/>新建 Sheet</button>
     </div>
    </div></>:<><div className="page-heading"><div><div className="eyebrow">通报项目</div><h1>从一个 Sheet 开始</h1><p>组织数据、编辑表格，搭建你的通报工作区。</p></div><Button onClick={askNewProject}><Plus/>新建项目</Button></div><div className="empty-state"><div className="empty-icon"><Table2 size={34}/></div><h2>把通报制作，放进同一个工作区</h2><p>为每项通报建立项目，用多个 Sheet 组织你的数据。<br/>从空白表格开始，也可以直接粘贴 Excel 中的数据。</p><Button onClick={askNewProject}><Plus/>创建第一个项目</Button><div className="empty-capabilities"><span>多 Sheet 管理</span><span>单元格编辑</span><span>批量粘贴</span></div></div></>}
  </section>
  <Dialog open={!!modal} onOpenChange={open=>!open&&setModal(null)}><DialogContent><DialogHeader><DialogTitle>{modal?.title}</DialogTitle><DialogDescription>{modal?.type==='project'?'一个项目对应一项通报任务。':'名称修改后将自动保存。'}</DialogDescription></DialogHeader><form onSubmit={submit}><label className="field-label" htmlFor="entity-name">名称</label><Input id="entity-name" value={modal?.value??''} onChange={e=>setModal(m=>m&&({...m,value:e.target.value}))} autoFocus maxLength={modal?.type.includes('Sheet')||modal?.type==='sheet'?31:80} required/><div className="dialog-actions"><Button type="button" variant="outline" onClick={()=>setModal(null)}>取消</Button><Button type="submit" disabled={!modal?.value.trim()}>保存</Button></div></form></DialogContent></Dialog>
  {notice&&<output className={notice.error?'toast error':'toast'}>{notice.message}</output>}
 </main>;
}
