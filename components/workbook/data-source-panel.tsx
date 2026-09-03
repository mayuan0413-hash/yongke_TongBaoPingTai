'use client';
import { useEffect, useState } from 'react';
import { ChevronRight, Database, Loader2, Plus, RefreshCw, Settings2, Trash2, X } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import type { DataSource, DataSourceQuery, Project, QueryFilter, QueryOperator, QueryPreview, Sheet, SourceColumn, SourceTable } from '@/domain/workbook/types';
import { api } from '@/lib/api';

interface Props {
 project: Project; sheet: Sheet; onClose: () => void; report: (message:string,error?:boolean)=>void;
 onBind: (dataSourceId:string,query:DataSourceQuery)=>Promise<void>; onRefresh:()=>Promise<void>; onUnbind:()=>Promise<void>;
}
const operators: {value:QueryOperator;label:string}[]=[
 {value:'eq',label:'等于'},{value:'neq',label:'不等于'},{value:'gt',label:'大于'},{value:'gte',label:'大于等于'},
 {value:'lt',label:'小于'},{value:'lte',label:'小于等于'},{value:'contains',label:'包含'},{value:'startsWith',label:'开头是'},
 {value:'endsWith',label:'结尾是'},{value:'isNull',label:'为空'},{value:'notNull',label:'不为空'},
];
const emptyQuery:DataSourceQuery={table:'',fields:[],filters:[],orderBy:[],rowLimit:10000};

export function DataSourcePanel({project,sheet,onClose,report,onBind,onRefresh,onUnbind}:Props){
 const [sources,setSources]=useState<DataSource[]>([]),[tables,setTables]=useState<SourceTable[]>([]),[columns,setColumns]=useState<SourceColumn[]>([]);
 const [sourceId,setSourceId]=useState(sheet.dataSource?.dataSourceId??''),[query,setQuery]=useState<DataSourceQuery>(sheet.dataSource?.query??emptyQuery);
 const [sourceName,setSourceName]=useState('业务数据库'),[preview,setPreview]=useState<QueryPreview|null>(null),[busy,setBusy]=useState(''),[loadError,setLoadError]=useState('');
 const loadSources=()=>api.dataSources(project.id).then(items=>{setSources(items);setSourceId(current=>current||items[0]?.id||'');return items;});
 useEffect(()=>{Promise.all([api.dataSources(project.id).then(items=>{setSources(items);setSourceId(current=>current||items[0]?.id||'');}),api.tables('business').then(setTables)]).catch(e=>setLoadError((e as Error).message));},[project.id]);
 useEffect(()=>{if(query.table)api.columns('business',query.table).then(setColumns).catch(e=>setLoadError((e as Error).message));},[query.table]);
 const valid=query.table&&query.fields.length&&sourceId;
 const selected=new Set(query.fields);
 const updateFilter=(index:number,patch:Partial<QueryFilter>)=>setQuery(q=>({...q,filters:q.filters.map((f,i)=>i===index?{...f,...patch}:f)}));
 const run=async(kind:string,work:()=>Promise<void>)=>{setBusy(kind);setLoadError('');try{await work();}catch(e){const m=(e as Error).message;setLoadError(m);report(m,true);}finally{setBusy('');}};
 const createSource=()=>run('create',async()=>{const item=await api.createDataSource(project.id,sourceName,'business');await loadSources();setSourceId(item.id);report('数据源已创建');});
 const selectTable=(table:string)=>{setQuery({...emptyQuery,table});setColumns([]);setPreview(null);};
 const doPreview=()=>run('preview',async()=>setPreview(await api.preview(project.id,sourceId,query)));
 const save=()=>run('save',async()=>{await onBind(sourceId,query);report('查询配置已保存');});
 const saveRefresh=()=>run('refresh',async()=>{await onBind(sourceId,query);await onRefresh();report('数据已刷新');});
 return <aside className="source-panel" aria-label="数据源配置">
  <header><div><span className="panel-kicker">DATA SOURCE</span><h2><Database/>数据源配置</h2></div><button onClick={onClose} aria-label="关闭数据源配置"><X/></button></header>
  <div className="source-scroll">
   {loadError&&<div className="panel-error">{loadError}</div>}
   <section><h3><span>1</span>选择数据源</h3>
    {sources.length?<select value={sourceId} onChange={e=>setSourceId(e.target.value)}>{sources.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select>:
    <div className="create-source"><Input value={sourceName} onChange={e=>setSourceName(e.target.value)} maxLength={80}/><Button onClick={createSource} disabled={!sourceName.trim()||!!busy}>{busy==='create'?<Loader2 className="spin"/>:<Plus/>}创建</Button></div>}
    <p>连接由服务端管理，浏览器不保存数据库凭据。</p>
   </section>
   <section><h3><span>2</span>配置查询</h3>
    <div className="control-label"><span>数据表</span><select aria-label="数据表" value={query.table} onChange={e=>selectTable(e.target.value)}><option value="">请选择数据表</option>{tables.map(t=><option key={t.name} value={t.name}>{t.label}</option>)}</select></div>
    {!!query.table&&<><div className="field-title"><span>查询字段</span><button onClick={()=>setQuery(q=>({...q,fields:q.fields.length===columns.length?[]:columns.map(c=>c.name)}))}>{query.fields.length===columns.length?'清空':'全选'}</button></div>
     <div className="field-list">{columns.map(c=><label key={c.name}><input type="checkbox" checked={selected.has(c.name)} onChange={e=>setQuery(q=>({...q,fields:e.target.checked?[...q.fields,c.name]:q.fields.filter(f=>f!==c.name),filters:q.filters.filter(f=>f.field!==c.name),orderBy:q.orderBy.filter(o=>o.field!==c.name)}))}/><span>{c.label}</span><small>{c.name}{c.dataType&&` · ${c.dataType}`}</small></label>)}</div>
     <div className="condition-head"><span>查询条件（全部满足）</span><button disabled={!query.fields.length||query.filters.length>=20} onClick={()=>setQuery(q=>({...q,filters:[...q.filters,{field:q.fields[0],operator:'eq',value:''}]}))}><Plus/>添加</button></div>
     <div className="filters">{query.filters.map((f,i)=><div key={i}><select value={f.field} onChange={e=>updateFilter(i,{field:e.target.value})}>{query.fields.map(field=><option key={field}>{field}</option>)}</select><select value={f.operator} onChange={e=>updateFilter(i,{operator:e.target.value as QueryOperator})}>{operators.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}</select>{!['isNull','notNull'].includes(f.operator)&&<Input value={f.value??''} onChange={e=>updateFilter(i,{value:e.target.value})} placeholder="条件值"/>}<button aria-label="删除条件" onClick={()=>setQuery(q=>({...q,filters:q.filters.filter((_,n)=>n!==i)}))}><X/></button></div>)}</div>
     <div className="query-row"><div className="control-label"><span>排序字段</span><select aria-label="排序字段" value={query.orderBy[0]?.field??''} onChange={e=>setQuery(q=>({...q,orderBy:e.target.value?[{field:e.target.value,direction:q.orderBy[0]?.direction??'asc'}]:[]}))}><option value="">保持数据库顺序</option>{query.fields.map(f=><option key={f}>{f}</option>)}</select></div><div className="control-label"><span>方向</span><select aria-label="排序方向" value={query.orderBy[0]?.direction??'asc'} disabled={!query.orderBy.length} onChange={e=>setQuery(q=>({...q,orderBy:q.orderBy.length?[{...q.orderBy[0],direction:e.target.value as 'asc'|'desc'}]:[]}))}><option value="asc">升序</option><option value="desc">降序</option></select></div><div className="control-label"><span>最大行数</span><Input aria-label="最大行数" type="number" min={1} max={100000} value={query.rowLimit} onChange={e=>setQuery(q=>({...q,rowLimit:Number(e.target.value)}))}/></div></div>
    </>}
   </section>
   {preview&&<section className="preview-block"><h3><span>3</span>查询预览 <small>{preview.rows.length} 行{preview.hasMore?'（还有更多）':''}</small></h3><div><table><thead><tr>{preview.columns.map(c=><th key={c}>{columns.find(x=>x.name===c)?.label??c}</th>)}</tr></thead><tbody>{preview.rows.map((row,i)=><tr key={i}>{preview.columns.map(c=><td key={c}>{row[c]??''}</td>)}</tr>)}</tbody></table></div></section>}
   {sheet.dataSource&&<section className="source-state"><h3><span>✓</span>当前快照</h3><dl><div><dt>数据源</dt><dd>{sheet.dataSource.dataSourceName}</dd></div><div><dt>数据表</dt><dd>{sheet.dataSource.query.table}</dd></div><div><dt>数据行</dt><dd>{sheet.dataSource.lastRowCount.toLocaleString()}</dd></div><div><dt>最近刷新</dt><dd>{sheet.dataSource.lastRefreshedAt?new Date(sheet.dataSource.lastRefreshedAt).toLocaleString('zh-CN'):'尚未刷新'}</dd></div></dl><Button variant="outline" onClick={()=>run('unbind',onUnbind)} disabled={!!busy}><Trash2/>切换为手工数据</Button></section>}
  </div>
  <footer><Button variant="outline" disabled={!valid||!!busy} onClick={doPreview}>{busy==='preview'?<Loader2 className="spin"/>:<Settings2/>}预览查询</Button><Button variant="outline" disabled={!valid||!!busy} onClick={save}>保存配置</Button><Button disabled={!valid||!!busy} onClick={saveRefresh}>{busy==='refresh'?<Loader2 className="spin"/>:<RefreshCw/>}保存并刷新<ChevronRight/></Button></footer>
 </aside>;
}
