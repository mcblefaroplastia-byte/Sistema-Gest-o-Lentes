import React,{useEffect,useMemo,useState}from'react';import{LayoutDashboard,Boxes,History,CalendarDays,BellRing,Truck,FileText,Users,Settings,Search,Plus,Pencil,Trash2,Download,CheckCircle2,AlertTriangle,Clock3,PackageCheck,Save,X,FileSpreadsheet,FileDown,SlidersHorizontal,Eye,RotateCcw,HelpCircle,BookOpen,ChevronDown}from'lucide-react';
import * as XLSX from 'xlsx';
import {jsPDF} from 'jspdf';
import autoTable from 'jspdf-autotable';
import Login from './Login';
import {supabase} from './lib/supabase';
import LensScanner from './LensScanner';
const K='oftalmocastro_lens_stock_v2';
const seed={lenses:[{id:1,name:'AcrySof IQ',manufacturer:'Alcon',model:'SN60WF',power:'+21.0 D',type:'Monofocal',lot:'L240901',serial:'SN-000021',barcode:'',qty:8,min:3,expiry:'2028-09-30',location:'Armário A • Gaveta 1',supplier:'Alcon Brasil',status:'Disponível',createdAt:'11/09/2026 08:12',updatedAt:'11/09/2026 08:12'},{id:2,name:'Vivity',manufacturer:'Alcon',model:'DFT015',power:'+20.5 D',type:'EDOF',lot:'V260722',serial:'VV-205-778',barcode:'',qty:2,min:2,expiry:'2028-07-22',location:'Armário Premium • Gaveta 2',supplier:'Alcon Brasil',status:'Baixo estoque',createdAt:'11/09/2026 08:20',updatedAt:'11/09/2026 09:01'}],movements:[{id:101,date:'2026-09-11',time:'09:05',type:'Entrada',lens:'AcrySof IQ +21.0 D',qty:5,user:'Juliana',reason:'Compra / reposição',patient:'',surgery:''},{id:102,date:'2026-09-11',time:'13:20',type:'Saída',lens:'Vivity +20.5 D',qty:1,user:'Bruna',reason:'Utilizada em cirurgia',patient:'Paciente exemplo',surgery:'Catarata'}],reservations:[{id:201,patient:'Amélia',surgery:'Catarata',doctor:'Dr. Daniel',date:'2026-09-14',time:'13:30',lens:'AcrySof IQ',power:'+21.0 D',eye:'OD',status:'Confirmada'}],suppliers:[{id:301,name:'Alcon Brasil',contact:'Comercial',phone:'(11) 0000-0000',email:'comercial@exemplo.com',lead:'3-5 dias'}],users:[{id:401,name:'Administrador',role:'Administrador',active:true},{id:402,name:'Juliana',role:'Estoque / Cirurgias',active:true}]};
const load=()=>{try{return JSON.parse(localStorage.getItem(K))||seed}catch{return seed}};const fmt=d=>d?new Date(d+'T12:00:00').toLocaleDateString('pt-BR'):'—';const now=()=>{const d=new Date();return{date:d.toISOString().slice(0,10),time:d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}),stamp:d.toLocaleString('pt-BR')}};
const csv=(rows,name)=>{if(!rows.length)return;let keys=Object.keys(rows[0]);let esc=v=>'"'+String(v??'').replaceAll('"','""')+'"';let txt=[keys.join(';'),...rows.map(r=>keys.map(k=>esc(r[k])).join(';'))].join('\n');let b=new Blob(['\ufeff'+txt],{type:'text/csv'}),u=URL.createObjectURL(b),a=document.createElement('a');a.href=u;a.download=name;a.click();URL.revokeObjectURL(u)};

const safeName=s=>String(s||'relatorio').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9_-]+/g,'-').replace(/^-+|-+$/g,'').toLowerCase();
const inRange=(date,from,to)=>{if(!date)return true;if(from&&date<from)return false;if(to&&date>to)return false;return true};
const contains=(v,q)=>!q||String(v??'').toLowerCase().includes(String(q).toLowerCase());
const downloadCSV=(columns,rows,name)=>{if(!rows.length)return alert('Nenhum registro encontrado para exportar.');const esc=v=>'"'+String(v??'').replaceAll('"','""')+'"';const txt=[columns.map(c=>esc(c.label)).join(';'),...rows.map(r=>columns.map(c=>esc(r[c.key])).join(';'))].join('\n');const b=new Blob(['\ufeff'+txt],{type:'text/csv;charset=utf-8'}),u=URL.createObjectURL(b),a=document.createElement('a');a.href=u;a.download=name+'.csv';a.click();URL.revokeObjectURL(u)};
const downloadExcel=(columns,rows,meta,name)=>{if(!rows.length)return alert('Nenhum registro encontrado para exportar.');const exportRows=rows.map(r=>Object.fromEntries(columns.map(c=>[c.label,r[c.key]??''])));const ws=XLSX.utils.json_to_sheet(exportRows);ws['!cols']=columns.map(c=>({wch:Math.max(12,Math.min(34,c.width||18))}));ws['!autofilter']={ref:ws['!ref']};ws['!freeze']={xSplit:0,ySplit:1};const summary=XLSX.utils.aoa_to_sheet([['OFTALMOCASTRO - GESTÃO DE LENTES'],['Relatório',meta.title],['Período',meta.period],['Filtros',meta.filters||'Nenhum'],['Emitido em',meta.issuedAt],['Responsável',meta.user],['Total de registros',rows.length],[],['Observação','Arquivo gerado pelo módulo de relatórios do controle de estoque.']]);summary['!cols']=[{wch:22},{wch:70}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,summary,'Resumo');XLSX.utils.book_append_sheet(wb,ws,'Relatório');XLSX.writeFile(wb,name+'.xlsx')};
const downloadPDF=(columns,rows,meta,name)=>{if(!rows.length)return alert('Nenhum registro encontrado para exportar.');const doc=new jsPDF({orientation:columns.length>7?'landscape':'portrait',unit:'mm',format:'a4'});const pageW=doc.internal.pageSize.getWidth();doc.setFillColor(13,42,53);doc.rect(0,0,pageW,25,'F');doc.setTextColor(255,255,255);doc.setFont('helvetica','bold');doc.setFontSize(15);doc.text('OFTALMOCASTRO',14,10);doc.setFontSize(8);doc.setFont('helvetica','normal');doc.text('Gestão de Estoque de Lentes Intraoculares',14,16);doc.setTextColor(25,35,42);doc.setFont('helvetica','bold');doc.setFontSize(12);doc.text(meta.title,14,34);doc.setFont('helvetica','normal');doc.setFontSize(8);doc.text(`Período: ${meta.period}`,14,40);doc.text(`Emitido em: ${meta.issuedAt}  •  Responsável: ${meta.user}`,14,45);const filters=meta.filters||'Nenhum filtro adicional';doc.text(`Filtros: ${filters}`.slice(0,180),14,50);doc.text(`Total de registros: ${rows.length}`,14,55);autoTable(doc,{startY:61,head:[columns.map(c=>c.label)],body:rows.map(r=>columns.map(c=>String(r[c.key]??''))),styles:{fontSize:6.7,cellPadding:1.8,overflow:'linebreak'},headStyles:{fillColor:[23,63,77],textColor:255,fontStyle:'bold'},alternateRowStyles:{fillColor:[247,249,250]},margin:{left:10,right:10,bottom:13},didDrawPage:({pageNumber})=>{const h=doc.internal.pageSize.getHeight();doc.setFontSize(7);doc.setTextColor(110);doc.text(`OftalmoCastro • Página ${pageNumber}`,14,h-6);doc.text('Documento gerado eletronicamente pelo sistema de gestão de lentes.',pageW-14,h-6,{align:'right'})}});doc.save(name+'.pdf')};

const Badge=({children,tone='blue'})=><span className={'badge '+tone}>{children}</span>;const Field=({label,children})=><label className="field"><span>{label}</span>{children}</label>;function Modal({title,onClose,children,wide}){return <div className="modal-bg"><div className={'modal '+(wide?'wide':'')}><div className="modal-head"><h3>{title}</h3><button className="icon" onClick={onClose}><X size={18}/></button></div>{children}</div></div>}
function Sistema({profile}){const[data,setData]=useState(load),[page,setPage]=useState('dashboard'),[query,setQuery]=useState(''),[modal,setModal]=useState(null),[toast,setToast]=useState('');useEffect(()=>localStorage.setItem(K,JSON.stringify(data)),[data]);useEffect(()=>{if(!toast)return;let t=setTimeout(()=>setToast(''),2200);return()=>clearTimeout(t)},[toast]);const low=data.lenses.filter(l=>l.qty<=l.min),total=data.lenses.reduce((a,b)=>a+Number(b.qty||0),0),exp=data.lenses.filter(l=>(new Date(l.expiry)-new Date())/864e5<180);const baseNav=[['dashboard','Visão geral',LayoutDashboard],['stock','Estoque de lentes',Boxes],['movements','Movimentações',History],['reservations','Cirurgias / Reservas',CalendarDays],['alerts','Alertas',BellRing],['suppliers','Fornecedores',Truck],['reports','Relatórios',FileText],['help','Central de ajuda',HelpCircle]];
const nav=profile?.role==='admin'?[...baseNav.slice(0,7),['users','Usuários e acessos',Users],['settings','Configurações',Settings],...baseNav.slice(7)]:baseNav;
const currentUserName=profile?.full_name||'Usuário';
const currentRoleLabel=profile?.role==='admin'?'Administrador':'Estoque / Cirurgias';
const saveLens=f=>{const n=now();setData(d=>{let ex=d.lenses.find(x=>x.id===f.id),it={...f,qty:+f.qty,min:+f.min,createdAt:ex?.createdAt||n.stamp,updatedAt:n.stamp};return{...d,lenses:ex?d.lenses.map(x=>x.id===f.id?it:x):[{...it,id:Date.now()},...d.lenses]}});setModal(null);setToast('Lente salva com sucesso')};
const removeLens=id=>{if(confirm('Excluir esta lente?')){setData(d=>({...d,lenses:d.lenses.filter(x=>x.id!==id)}));setToast('Lente excluída')}};
const saveMov=f=>{let n=now(),q=+f.qty;setData(d=>{let lenses=d.lenses.map(l=>{if(String(l.id)!==String(f.lensId))return l;let nq=l.qty;if(f.type==='Entrada'||f.type==='Devolução')nq+=q;if(f.type==='Saída')nq=Math.max(0,nq-q);return{...l,qty:nq,status:nq<=l.min?'Baixo estoque':'Disponível',updatedAt:n.stamp}});let l=d.lenses.find(x=>String(x.id)===String(f.lensId));let m={id:Date.now(),date:f.date,time:f.time,type:f.type,lens:l?l.name+' '+l.power:'—',qty:q,user:currentUserName,reason:f.reason,patient:f.patient,surgery:f.surgery};return{...d,lenses,movements:[m,...d.movements]}});setModal(null);setToast('Movimentação registrada')};
const saveRes=f=>{setData(d=>{let ex=d.reservations.find(x=>x.id===f.id);return{...d,reservations:ex?d.reservations.map(x=>x.id===f.id?f:x):[{...f,id:Date.now()},...d.reservations]}});setModal(null);setToast('Reserva salva')};
return <div className="app"><aside><div className="brand"><div className="mark">OC</div><div><strong>OftalmoCastro</strong><span>Gestão de Lentes</span></div></div><nav>{nav.map(([id,t,I])=><button className={page===id?'active':''} onClick={()=>{setPage(id);setQuery('')}} key={id}><I size={18}/><span>{t}</span>{id==='alerts'&&low.length>0?<b>{low.length}</b>:null}</button>)}</nav><div className="side-note">Controle de estoque cirúrgico</div></aside><main><header><div><h1>{nav.find(x=>x[0]===page)?.[1]}</h1><p>Estoque, rastreabilidade, reservas e auditoria de lentes intraoculares</p></div><div className="head-actions">{page!=='reports'&&<div className="search"><Search size={16}/><input placeholder="Buscar..." value={query} onChange={e=>setQuery(e.target.value)}/></div>}<div style={{display:'flex',alignItems:'center',gap:8,padding:'7px 10px',border:'1px solid #dfe7e5',borderRadius:10,background:'#fff'}}><div style={{lineHeight:1.15}}><b style={{display:'block',fontSize:12}}>{currentUserName}</b><small style={{fontSize:10,color:'#6d7b7d'}}>{currentRoleLabel}</small></div><button className="secondary" style={{padding:'6px 9px'}} onClick={()=>supabase.auth.signOut()}>Sair</button></div><button className="primary" onClick={()=>setModal({type:'movement'})}><Plus size={16}/>Nova movimentação</button></div></header><section className="content">{page==='dashboard'&&<Dashboard data={data} low={low} total={total} exp={exp} setPage={setPage}/>} {page==='stock'&&<Stock data={data} query={query} onNew={()=>setModal({type:'lens'})} onEdit={x=>setModal({type:'lens',item:x})} onDelete={removeLens}/>} {page==='movements'&&<Movements data={data} query={query} onNew={()=>setModal({type:'movement'})}/>} {page==='reservations'&&<Reservations data={data} query={query} onNew={()=>setModal({type:'reservation'})} onEdit={x=>setModal({type:'reservation',item:x})} onDelete={id=>confirm('Excluir reserva?')&&setData(d=>({...d,reservations:d.reservations.filter(x=>x.id!==id)}))}/>} {page==='alerts'&&<Alerts low={low} exp={exp}/>} {page==='suppliers'&&<Suppliers data={data} setData={setData}/>} {page==='reports'&&<Reports data={data}/>} {page==='users'&&profile?.role==='admin'&&<UsersPage data={data} setData={setData}/>} {page==='settings'&&profile?.role==='admin'&&<SettingsPage setData={setData}/>} {page==='help'&&<HelpCenter setPage={setPage}/>}</section></main>{modal?.type==='lens'&&<LensModal item={modal.item} onClose={()=>setModal(null)} onSave={saveLens}/>} {modal?.type==='movement'&&<MovementModal lenses={data.lenses} currentUserName={currentUserName} onClose={()=>setModal(null)} onSave={saveMov}/>} {modal?.type==='reservation'&&<ReservationModal item={modal.item} lenses={data.lenses} onClose={()=>setModal(null)} onSave={saveRes}/>} {toast&&<div className="toast"><CheckCircle2 size={17}/>{toast}</div>}</div>}
function Dashboard({data,low,total,exp,setPage}){return <div className="dashboard-page">
 <div className="dashboard-kpis">
  <Stat t="Itens cadastrados" v={data.lenses.length} I={Boxes}/>
  <Stat t="Unidades em estoque" v={total} I={PackageCheck}/>
  <Stat t="Estoque baixo" v={low.length} I={AlertTriangle}/>
  <Stat t="Reservas cirúrgicas" v={data.reservations.length} I={CalendarDays}/>
 </div>
 <div className="dashboard-main-grid">
  <div className="panel dashboard-surgeries">
   <Head title="Próximas cirurgias" sub="Lentes reservadas por paciente" action={<button className="link" onClick={()=>setPage('reservations')}>Ver todas</button>}/>
   <div className="tablewrap"><table><thead><tr><th>Paciente</th><th>Data</th><th>Lente</th><th>Olho</th><th>Status</th></tr></thead><tbody>{data.reservations.slice(0,5).map(r=><tr key={r.id}><td><b>{r.patient}</b><small>{r.doctor}</small></td><td>{fmt(r.date)}<small>{r.time}</small></td><td>{r.lens}<small>{r.power}</small></td><td>{r.eye}</td><td><Badge tone="green">{r.status}</Badge></td></tr>)}</tbody></table></div>
  </div>
  <div className="panel dashboard-alerts">
   <Head title="Alertas do estoque" sub="Itens que precisam de atenção" action={<button className="link" onClick={()=>setPage('alerts')}>Abrir alertas</button>}/>
   <div className="dashboard-alert-list">{low.map(l=><div className="alert" key={'l'+l.id}><AlertTriangle/><div><b>{l.name} {l.power}</b><span>Saldo {l.qty} • mínimo {l.min}</span></div><Badge tone="orange">Repor</Badge></div>)}{exp.map(l=><div className="alert" key={'e'+l.id}><Clock3/><div><b>{l.name} {l.power}</b><span>Validade {fmt(l.expiry)}</span></div><Badge tone="red">Validade</Badge></div>)}{!low.length&&!exp.length&&<div className="empty"><CheckCircle2/><b>Tudo em ordem</b></div>}</div>
  </div>
 </div>
 <div className="panel dashboard-movements">
  <Head title="Últimas movimentações" sub="Rastreabilidade por data, horário e responsável"/>
  <MovementTable rows={data.movements.slice(0,6)}/>
 </div>
 </div>}
const Stat=({t,v,I})=><div className="stat"><div className="stat-ico"><I size={21}/></div><div><span>{t}</span><strong>{v}</strong></div></div>;const Head=({title,sub,action})=><div className="panel-head"><div><h2>{title}</h2><p>{sub}</p></div>{action}</div>;
function Stock({data,query,onNew,onEdit,onDelete}){let q=query.toLowerCase(),rows=data.lenses.filter(l=>Object.values(l).join(' ').toLowerCase().includes(q));return <div className="panel"><Head title="Cadastro e estoque" sub="Nome, potência, código, lote, série, validade, localização e fornecedor" action={<div className="actions"><button className="secondary" onClick={()=>csv(rows,'estoque-lentes.csv')}><Download size={15}/>Exportar</button><button className="primary" onClick={onNew}><Plus size={15}/>Cadastrar lente</button></div>}/><div className="tablewrap"><table><thead><tr><th>Lente</th><th>Potência / Tipo</th><th>Código</th><th>Lote / Série</th><th>Validade</th><th>Localização</th><th>Qtd.</th><th>Status</th><th></th></tr></thead><tbody>{rows.map(l=><tr key={l.id}><td><b>{l.name}</b><small>{l.manufacturer} • {l.model}</small></td><td>{l.power}<small>{l.type}</small></td><td>{l.barcode||'—'}</td><td>{l.lot}<small>{l.serial}</small></td><td>{fmt(l.expiry)}</td><td>{l.location}</td><td><b className="qty">{l.qty}</b><small>Mín. {l.min}</small></td><td><Badge tone={l.qty<=l.min?'orange':'green'}>{l.qty<=l.min?'Baixo estoque':'Disponível'}</Badge></td><td><div className="row-actions"><button onClick={()=>onEdit(l)}><Pencil size={14}/></button><button className="danger" onClick={()=>onDelete(l.id)}><Trash2 size={14}/></button></div></td></tr>)}</tbody></table></div></div>}
function MovementTable({rows}){return <div className="tablewrap"><table><thead><tr><th>Data / Hora</th><th>Tipo</th><th>Lente</th><th>Qtd.</th><th>Motivo</th><th>Paciente / Cirurgia</th><th>Responsável</th></tr></thead><tbody>{rows.map(m=><tr key={m.id}><td>{fmt(m.date)}<small>{m.time}</small></td><td><Badge tone={m.type==='Entrada'?'green':m.type==='Saída'?'red':'blue'}>{m.type}</Badge></td><td><b>{m.lens}</b></td><td>{m.qty}</td><td>{m.reason}</td><td>{m.patient||'—'}<small>{m.surgery}</small></td><td>{m.user}</td></tr>)}</tbody></table></div>}
function Movements({data,query,onNew}){let q=query.toLowerCase(),rows=data.movements.filter(m=>Object.values(m).join(' ').toLowerCase().includes(q));return <div className="panel"><Head title="Histórico de movimentações" sub="Entradas, saídas, reservas, devoluções e ajustes" action={<div className="actions"><button className="secondary" onClick={()=>csv(rows,'movimentacoes.csv')}><Download size={15}/>Exportar CSV</button><button className="primary" onClick={onNew}><Plus size={15}/>Registrar</button></div>}/><MovementTable rows={rows}/></div>}
function Reservations({data,query,onNew,onEdit,onDelete}){let q=query.toLowerCase(),rows=data.reservations.filter(r=>Object.values(r).join(' ').toLowerCase().includes(q));return <div className="panel"><Head title="Cirurgias e reservas" sub="Vincule a lente ao paciente antes da cirurgia" action={<button className="primary" onClick={onNew}><Plus size={15}/>Nova reserva</button>}/><div className="tablewrap"><table><thead><tr><th>Paciente</th><th>Cirurgia</th><th>Médico</th><th>Data / Hora</th><th>Lente</th><th>Olho</th><th>Status</th><th></th></tr></thead><tbody>{rows.map(r=><tr key={r.id}><td><b>{r.patient}</b></td><td>{r.surgery}</td><td>{r.doctor}</td><td>{fmt(r.date)}<small>{r.time}</small></td><td>{r.lens}<small>{r.power}</small></td><td>{r.eye}</td><td><Badge tone="green">{r.status}</Badge></td><td><div className="row-actions"><button onClick={()=>onEdit(r)}><Pencil size={14}/></button><button className="danger" onClick={()=>onDelete(r.id)}><Trash2 size={14}/></button></div></td></tr>)}</tbody></table></div></div>}
function Alerts({low,exp}){return <div className="grid2"><div className="panel"><Head title="Estoque baixo" sub="Itens no limite mínimo"/>{low.length?low.map(l=><div className="alert"><AlertTriangle/><div><b>{l.name} {l.power}</b><span>{l.qty} em estoque • mínimo {l.min}</span></div></div>):<div className="empty"><CheckCircle2/><b>Nenhum item abaixo do mínimo</b></div>}</div><div className="panel"><Head title="Validades" sub="Vencimentos nos próximos 180 dias"/>{exp.length?exp.map(l=><div className="alert"><Clock3/><div><b>{l.name} {l.power}</b><span>{fmt(l.expiry)}</span></div></div>):<div className="empty"><CheckCircle2/><b>Nenhuma validade próxima</b></div>}</div></div>}
function Suppliers({data,setData}){return <div className="panel"><Head title="Fornecedores" sub="Contatos e prazo médio de reposição" action={<button className="primary" onClick={()=>{let n=prompt('Nome do fornecedor:');if(n)setData(d=>({...d,suppliers:[...d.suppliers,{id:Date.now(),name:n,contact:'',phone:'',email:'',lead:''}]}))}}><Plus size={15}/>Novo fornecedor</button>}/><div className="supplier-grid">{data.suppliers.map(s=><div className="supplier" key={s.id}><Truck/><div><b>{s.name}</b><span>{s.contact||'Contato não informado'}</span><small>{s.phone} {s.email}</small><small>Prazo: {s.lead||'—'}</small></div><button className="icon danger" onClick={()=>confirm('Excluir fornecedor?')&&setData(d=>({...d,suppliers:d.suppliers.filter(x=>x.id!==s.id)}))}><Trash2 size={14}/></button></div>)}</div></div>}
function Reports({data}){
 const today=new Date().toISOString().slice(0,10),sixMonths=new Date(Date.now()+180*864e5).toISOString().slice(0,10);
 const emptyFilters={from:'',to:'',search:'',manufacturer:'',power:'',lot:'',supplier:'',movementType:'',user:'',patient:'',doctor:'',status:''};
 const [type,setType]=useState('stock'),[showFilters,setShowFilters]=useState(false),[filters,setFilters]=useState(emptyFilters);
 const setF=(k,v)=>setFilters(x=>({...x,[k]:v})),reset=()=>setFilters(emptyFilters);

 const reportDefs={
  stock:{title:'Posição atual do estoque',desc:'Saldo atual das lentes por modelo, potência, lote e validade',icon:Boxes},
  inventory:{title:'Inventário completo',desc:'Conferência técnica e física de todos os itens cadastrados',icon:PackageCheck},
  entries:{title:'Entradas de estoque',desc:'Compras, reposições e devoluções recebidas',icon:Plus},
  outputs:{title:'Saídas e consumo',desc:'Lentes retiradas ou utilizadas no período',icon:FileDown},
  movements:{title:'Movimentações completas',desc:'Histórico para auditoria do estoque',icon:History},
  surgeries:{title:'Lentes por cirurgia',desc:'Rastreabilidade por paciente, médico, olho e data',icon:CalendarDays},
  low:{title:'Estoque mínimo',desc:'Itens no limite ou abaixo da quantidade mínima',icon:AlertTriangle},
  expiring:{title:'Próximas do vencimento',desc:'Lentes que vencem nos próximos 180 dias',icon:Clock3},
  expired:{title:'Lentes vencidas',desc:'Itens cuja validade já expirou',icon:BellRing},
  consumption:{title:'Consumo por modelo e potência',desc:'Resumo das saídas por lente e numeração',icon:FileSpreadsheet},
  trace:{title:'Rastreabilidade por lote / série',desc:'Pesquisa para auditoria de lote e número de série',icon:Search}
 };

 const result=useMemo(()=>{
  const stockCols=[{key:'name',label:'Lente',width:22},{key:'manufacturer',label:'Fabricante',width:18},{key:'model',label:'Modelo',width:16},{key:'power',label:'Potência',width:14},{key:'type',label:'Tipo',width:18},{key:'lot',label:'Lote',width:16},{key:'serial',label:'Nº série',width:18},{key:'qty',label:'Qtd.',width:10},{key:'min',label:'Mín.',width:10},{key:'expiry',label:'Validade',width:14},{key:'location',label:'Localização',width:24},{key:'supplier',label:'Fornecedor',width:22},{key:'status',label:'Status',width:16}];
  const movCols=[{key:'date',label:'Data',width:14},{key:'time',label:'Hora',width:10},{key:'type',label:'Tipo',width:16},{key:'lens',label:'Lente / potência',width:28},{key:'qty',label:'Qtd.',width:10},{key:'reason',label:'Motivo',width:28},{key:'patient',label:'Paciente',width:24},{key:'surgery',label:'Cirurgia',width:18},{key:'user',label:'Responsável',width:18}];
  const resCols=[{key:'patient',label:'Paciente',width:24},{key:'surgery',label:'Cirurgia',width:18},{key:'doctor',label:'Médico',width:20},{key:'date',label:'Data',width:14},{key:'time',label:'Hora',width:10},{key:'lens',label:'Lente',width:22},{key:'power',label:'Potência',width:14},{key:'eye',label:'Olho',width:10},{key:'status',label:'Status',width:15}];
  const lensMatch=l=>contains([l.name,l.model,l.type,l.serial,l.location].join(' '),filters.search)&&contains(l.manufacturer,filters.manufacturer)&&contains(l.power,filters.power)&&contains(l.lot,filters.lot)&&contains(l.supplier,filters.supplier)&&contains(l.status,filters.status);
  const movMatch=m=>inRange(m.date,filters.from,filters.to)&&contains([m.lens,m.reason,m.surgery].join(' '),filters.search)&&contains(m.type,filters.movementType)&&contains(m.user,filters.user)&&contains(m.patient,filters.patient);
  const resMatch=r=>inRange(r.date,filters.from,filters.to)&&contains([r.lens,r.surgery,r.eye].join(' '),filters.search)&&contains(r.patient,filters.patient)&&contains(r.doctor,filters.doctor)&&contains(r.status,filters.status)&&contains(r.power,filters.power);

  if(['stock','inventory','low','expiring','expired','trace'].includes(type)){
   let rows=data.lenses.filter(lensMatch);
   if(type==='low')rows=rows.filter(l=>+l.qty<=+l.min);
   if(type==='expiring')rows=rows.filter(l=>l.expiry&&l.expiry>=today&&l.expiry<=sixMonths);
   if(type==='expired')rows=rows.filter(l=>l.expiry&&l.expiry<today);
   if(type==='trace')rows=rows.filter(l=>contains([l.lot,l.serial].join(' '),filters.search||filters.lot));
   return{columns:stockCols,rows:rows.map(l=>({...l,expiry:fmt(l.expiry)}))};
  }
  if(['entries','outputs','movements'].includes(type)){
   let rows=data.movements.filter(movMatch);
   if(type==='entries')rows=rows.filter(m=>m.type==='Entrada'||m.type==='Devolução');
   if(type==='outputs')rows=rows.filter(m=>m.type==='Saída');
   return{columns:movCols,rows:rows.map(m=>({...m,date:fmt(m.date)}))};
  }
  if(type==='surgeries')return{columns:resCols,rows:data.reservations.filter(resMatch).map(r=>({...r,date:fmt(r.date)}))};
  if(type==='consumption'){
   const map={};data.movements.filter(m=>m.type==='Saída'&&movMatch(m)).forEach(m=>{const k=m.lens||'Não informado';map[k]=(map[k]||0)+Number(m.qty||0)});
   return{columns:[{key:'lens',label:'Lente / modelo / potência',width:38},{key:'qty',label:'Quantidade consumida',width:22}],rows:Object.entries(map).map(([lens,qty])=>({lens,qty})).sort((a,b)=>b.qty-a.qty)};
  }
  return{columns:[],rows:[]};
 },[type,filters,data]);

 const def=reportDefs[type],Icon=def.icon;
 const activeFilters=[filters.from&&`de ${fmt(filters.from)}`,filters.to&&`até ${fmt(filters.to)}`,filters.search&&`busca: ${filters.search}`,filters.manufacturer&&`fabricante: ${filters.manufacturer}`,filters.power&&`potência: ${filters.power}`,filters.lot&&`lote: ${filters.lot}`,filters.supplier&&`fornecedor: ${filters.supplier}`,filters.movementType&&`tipo: ${filters.movementType}`,filters.user&&`responsável: ${filters.user}`,filters.patient&&`paciente: ${filters.patient}`,filters.doctor&&`médico: ${filters.doctor}`,filters.status&&`status: ${filters.status}`].filter(Boolean).join(' • ');
 const activeCount=Object.values(filters).filter(Boolean).length;
 const meta={title:def.title,period:filters.from||filters.to?`${filters.from?fmt(filters.from):'início'} a ${filters.to?fmt(filters.to):'hoje'}`:'Todos os registros',filters:activeFilters,issuedAt:new Date().toLocaleString('pt-BR'),user:'Administrador'};
 const file=safeName('oftalmocastro-'+def.title+'-'+new Date().toISOString().slice(0,10));

 return <div className="reports-clean">
  <div className="reports-clean-head">
   <div><h2>Relatórios</h2><p>Selecione um relatório e visualize somente as informações necessárias.</p></div>
   <div className="report-selector">
    <label>Tipo de relatório</label>
    <select value={type} onChange={e=>{setType(e.target.value);setShowFilters(false)}}>
     <optgroup label="Estoque">
      <option value="stock">Posição atual do estoque</option>
      <option value="inventory">Inventário completo</option>
      <option value="low">Estoque mínimo</option>
      <option value="expiring">Próximas do vencimento</option>
      <option value="expired">Lentes vencidas</option>
     </optgroup>
     <optgroup label="Movimentações">
      <option value="entries">Entradas de estoque</option>
      <option value="outputs">Saídas e consumo</option>
      <option value="movements">Movimentações completas</option>
      <option value="consumption">Consumo por modelo e potência</option>
     </optgroup>
     <optgroup label="Cirurgias e auditoria">
      <option value="surgeries">Lentes por cirurgia</option>
      <option value="trace">Rastreabilidade por lote / série</option>
     </optgroup>
    </select>
   </div>
  </div>

  <div className="report-focus-card">
   <div className="report-focus-top">
    <div className="report-focus-title"><div className="report-focus-icon"><Icon size={22}/></div><div><span>RELATÓRIO ATUAL</span><h3>{def.title}</h3><p>{def.desc}</p></div></div>
    <div className="report-focus-actions">
     <button className="secondary" onClick={()=>setShowFilters(x=>!x)}><SlidersHorizontal size={15}/>{showFilters?'Ocultar filtros':'Filtrar'}{activeCount>0&&<b className="filter-pill">{activeCount}</b>}</button>
    </div>
   </div>

   {showFilters&&<div className="report-filter-simple">
    <div className="filter-simple-head"><div><b>Filtros</b><span>Preencha somente o que deseja pesquisar.</span></div>{activeCount>0&&<button className="link" onClick={reset}><RotateCcw size={13}/>Limpar</button>}</div>
    <div className="formgrid nopad report-filter-simple-grid">
     <Field label="Data inicial"><input type="date" value={filters.from} onChange={e=>setF('from',e.target.value)}/></Field>
     <Field label="Data final"><input type="date" value={filters.to} onChange={e=>setF('to',e.target.value)}/></Field>
     <Field label="Busca geral"><input placeholder="Nome, lente, modelo..." value={filters.search} onChange={e=>setF('search',e.target.value)}/></Field>
     <Field label="Potência / Numeração"><input placeholder="+21.0 D" value={filters.power} onChange={e=>setF('power',e.target.value)}/></Field>
     <Field label="Fabricante"><input value={filters.manufacturer} onChange={e=>setF('manufacturer',e.target.value)}/></Field>
     <Field label="Lote"><input value={filters.lot} onChange={e=>setF('lot',e.target.value)}/></Field>
     <Field label="Fornecedor"><input value={filters.supplier} onChange={e=>setF('supplier',e.target.value)}/></Field>
     <Field label="Movimentação"><select value={filters.movementType} onChange={e=>setF('movementType',e.target.value)}><option value="">Todas</option><option>Entrada</option><option>Saída</option><option>Reserva</option><option>Devolução</option><option>Ajuste de inventário</option></select></Field>
     <Field label="Responsável"><input value={filters.user} onChange={e=>setF('user',e.target.value)}/></Field>
     <Field label="Paciente"><input value={filters.patient} onChange={e=>setF('patient',e.target.value)}/></Field>
     <Field label="Médico"><input value={filters.doctor} onChange={e=>setF('doctor',e.target.value)}/></Field>
     <Field label="Status"><input value={filters.status} onChange={e=>setF('status',e.target.value)}/></Field>
    </div>
   </div>}

   {activeCount>0&&<div className="report-active-filter"><SlidersHorizontal size={14}/><span>{activeFilters}</span><button onClick={reset}>Limpar filtros</button></div>}

   <div className="report-toolbar">
    <div><strong>{result.rows.length}</strong><span> registro(s) encontrado(s)</span></div>
    <div className="report-downloads">
     <button className="secondary" onClick={()=>downloadExcel(result.columns,result.rows,meta,file)}><FileSpreadsheet size={15}/>Excel</button>
     <button className="primary" onClick={()=>downloadPDF(result.columns,result.rows,meta,file)}><FileDown size={15}/>PDF</button>
    </div>
   </div>

   {result.rows.length?<div className="tablewrap report-clean-table"><table><thead><tr>{result.columns.map(c=><th key={c.key}>{c.label}</th>)}</tr></thead><tbody>{result.rows.slice(0,10).map((r,i)=><tr key={i}>{result.columns.map(c=><td key={c.key}>{r[c.key]??'—'}</td>)}</tr>)}</tbody></table><div className="report-preview-note">Prévia de até 10 registros. Excel e PDF incluem todos os resultados encontrados.</div></div>:<div className="empty report-clean-empty"><Eye size={30}/><b>Nenhum registro encontrado</b><span>Altere os filtros para realizar uma nova busca.</span></div>}
  </div>
 </div>
}


function HelpCenter({setPage}){
 const [term,setTerm]=useState(''),[open,setOpen]=useState('inicio');
 const sections=[
  {id:'inicio',title:'Primeiros passos',icon:LayoutDashboard,desc:'Entenda a lógica do sistema e por onde começar.',content:<>
   <p>O sistema organiza o controle das lentes intraoculares em um único lugar. O fluxo recomendado é: cadastrar as lentes, registrar entradas, reservar para as cirurgias, registrar a saída quando a lente for utilizada e acompanhar alertas e relatórios.</p>
   <div className="help-steps">
    <div><b>1</b><span><strong>Cadastre a lente</strong>Informe nome, fabricante, modelo, potência, lote, série, validade, quantidade, estoque mínimo, localização e fornecedor.</span></div>
    <div><b>2</b><span><strong>Registre as entradas</strong>Toda reposição recebida deve ser lançada em Nova movimentação como Entrada.</span></div>
    <div><b>3</b><span><strong>Faça a reserva</strong>Antes da cirurgia, vincule a lente ao paciente, médico, data, horário e olho.</span></div>
    <div><b>4</b><span><strong>Registre a utilização</strong>Após o uso, faça a movimentação de Saída para manter o saldo correto.</span></div>
   </div>
  </>},
  {id:'dashboard',title:'Visão geral',icon:LayoutDashboard,desc:'Como interpretar o painel inicial.',content:<>
   <p>A Visão geral apresenta um resumo rápido do estoque. Os indicadores mostram itens cadastrados, unidades em estoque, itens com estoque baixo e quantidade de reservas cirúrgicas.</p>
   <p>Em <b>Próximas cirurgias</b>, você acompanha as reservas mais recentes. Em <b>Alertas do estoque</b>, aparecem lentes que atingiram o estoque mínimo ou estão próximas do vencimento. Em <b>Últimas movimentações</b>, ficam os lançamentos mais recentes para rastreabilidade.</p>
  </>},
  {id:'estoque',title:'Estoque de lentes',icon:Boxes,desc:'Cadastro, edição, saldo e informações de cada lente.',content:<>
   <p>Use <b>Estoque de lentes</b> para consultar tudo que está cadastrado. O botão <b>Cadastrar lente</b> abre o formulário completo. Nome e potência/numeração são campos obrigatórios.</p>
   <p>A quantidade representa o saldo atual. O estoque mínimo define quando o sistema deve gerar um alerta. Para corrigir dados cadastrais, use o ícone de lápis. Para localizar rapidamente uma lente, use a busca no topo da tela.</p>
   <div className="help-tip"><AlertTriangle size={17}/><span><b>Atenção:</b> para entrada ou retirada de unidades, prefira registrar uma movimentação em vez de alterar manualmente a quantidade. Assim o histórico fica preservado.</span></div>
  </>},
  {id:'movimentacoes',title:'Movimentações',icon:History,desc:'Entradas, saídas, devoluções e ajustes.',content:<>
   <p>Clique em <b>Nova movimentação</b> e selecione o tipo. Em <b>Entrada</b>, a quantidade é acrescentada ao estoque. Em <b>Saída</b>, ela é descontada. O sistema também oferece Reserva, Devolução e Ajuste de inventário como tipos de registro.</p>
   <p>Preencha a lente, quantidade, responsável, data, horário e motivo. Paciente e cirurgia podem ser informados quando a movimentação estiver relacionada a um procedimento. O histórico registra data, tipo, lente, quantidade, motivo e responsável.</p>
  </>},
  {id:'reservas',title:'Cirurgias e reservas',icon:CalendarDays,desc:'Como vincular uma lente a uma cirurgia.',content:<>
   <p>Em <b>Cirurgias / Reservas</b>, clique em <b>Nova reserva</b>. Informe paciente, cirurgia, médico, data, horário, lente, olho e status.</p>
   <p>Os status disponíveis atualmente são <b>Reservada, Confirmada, Realizada e Cancelada</b>. Uma reserva existente pode ser editada pelo lápis ou excluída pelo ícone de lixeira.</p>
   <div className="help-tip"><HelpCircle size={17}/><span>A reserva organiza a programação cirúrgica. No código atual, o saldo físico é alterado pela movimentação de estoque; portanto, após a utilização, registre também a saída correspondente.</span></div>
  </>},
  {id:'alertas',title:'Alertas',icon:BellRing,desc:'Estoque mínimo e validade das lentes.',content:<>
   <p>O sistema sinaliza <b>estoque baixo</b> quando a quantidade da lente é menor ou igual ao mínimo configurado. Também acompanha lentes com vencimento dentro dos próximos <b>180 dias</b>.</p>
   <p>O número exibido ao lado de Alertas no menu corresponde aos itens que estão no limite ou abaixo do estoque mínimo.</p>
  </>},
  {id:'fornecedores',title:'Fornecedores',icon:Truck,desc:'Cadastro dos contatos de reposição.',content:<>
   <p>Nesta área ficam os fornecedores e o prazo médio de reposição. Use <b>Novo fornecedor</b> para cadastrar um nome. Os dados ajudam a equipe a identificar rapidamente de quem solicitar uma reposição.</p>
  </>},
  {id:'relatorios',title:'Relatórios',icon:FileText,desc:'Filtros, consultas, Excel e PDF.',content:<>
   <p>No campo <b>Tipo de relatório</b>, escolha o que deseja consultar. Há relatórios de posição do estoque, inventário, estoque mínimo, vencimentos, entradas, saídas, movimentações, consumo, lentes por cirurgia e rastreabilidade por lote/série.</p>
   <p>Clique em <b>Filtrar</b> para abrir somente quando precisar os filtros de data, busca, potência, fabricante, lote, fornecedor, movimentação, responsável, paciente, médico e status. A tela mostra uma prévia de até 10 registros, enquanto os arquivos exportados incluem todos os resultados encontrados.</p>
   <p>Use <b>Excel</b> para análise em planilha ou <b>PDF</b> para gerar um documento organizado do relatório selecionado.</p>
  </>},
  {id:'usuarios',title:'Usuários e acessos',icon:Users,desc:'Responsáveis e rastreabilidade.',content:<>
   <p>A área de usuários mantém os nomes e perfis usados como base para organização e rastreabilidade. É possível cadastrar um novo usuário e alternar seu status entre ativo e inativo.</p>
  </>},
  {id:'config',title:'Configurações',icon:Settings,desc:'Parâmetros gerais e demonstração.',content:<>
   <p>As configurações exibem os parâmetros gerais do sistema: estoque mínimo definido individualmente por lente, alerta de validade em 180 dias e auditoria por data, horário e responsável.</p>
   <div className="help-tip danger-help"><AlertTriangle size={17}/><span><b>Restaurar demonstração</b> substitui os dados atuais pelos dados de exemplo do sistema. Use essa opção somente quando realmente quiser reiniciar a demonstração.</span></div>
  </>},
  {id:'duvidas',title:'Dúvidas frequentes',icon:HelpCircle,desc:'Respostas rápidas para situações do dia a dia.',content:<>
   <div className="faq-list">
    <div><b>O saldo da lente ficou errado. O que faço?</b><span>Confira primeiro o histórico de movimentações. Se houve um lançamento incorreto, identifique o registro e faça o ajuste necessário de forma rastreável.</span></div>
    <div><b>Como encontro uma lente específica?</b><span>Na página Estoque de lentes, use a busca superior. Nos Relatórios, você também pode pesquisar por modelo, potência, lote e outros filtros.</span></div>
    <div><b>Como sei o que precisa ser reposto?</b><span>Abra Alertas. Os itens cuja quantidade chegou ao estoque mínimo aparecem em Estoque baixo.</span></div>
    <div><b>Como consultar um lote?</b><span>Abra Relatórios, selecione Rastreabilidade por lote / série, clique em Filtrar e informe o lote ou a busca desejada.</span></div>
    <div><b>Como exporto as informações?</b><span>Em Relatórios, selecione o relatório, aplique os filtros e use Excel ou PDF. Algumas telas também possuem exportação própria.</span></div>
    <div><b>Cadastrei uma lente com informação errada.</b><span>Abra Estoque de lentes e clique no lápis da linha correspondente para editar o cadastro.</span></div>
   </div>
  </>}
 ];
 const filtered=sections.filter(s=>!term||(`${s.title} ${s.desc}`).toLowerCase().includes(term.toLowerCase()));
 return <div className="help-page">
  <div className="help-hero">
   <div className="help-hero-copy"><div className="help-hero-icon"><BookOpen size={24}/></div><div><span>CENTRAL DE AJUDA</span><h2>Como podemos ajudar?</h2><p>Manual completo para usar o Gestão de Lentes com segurança e organização.</p></div></div>
   <div className="help-search"><Search size={17}/><input value={term} onChange={e=>setTerm(e.target.value)} placeholder="Pesquisar uma dúvida ou função..."/></div>
  </div>
  <div className="help-quick">
   <button onClick={()=>setPage('stock')}><Boxes size={18}/><span><b>Cadastrar lente</b><small>Ir para o estoque</small></span></button>
   <button onClick={()=>setPage('movements')}><History size={18}/><span><b>Movimentações</b><small>Consultar histórico</small></span></button>
   <button onClick={()=>setPage('reservations')}><CalendarDays size={18}/><span><b>Reservar lente</b><small>Ir para cirurgias</small></span></button>
   <button onClick={()=>setPage('reports')}><FileText size={18}/><span><b>Gerar relatório</b><small>Abrir relatórios</small></span></button>
  </div>
  <div className="help-list">
   {filtered.map(s=>{const I=s.icon,isOpen=open===s.id;return <div className={'help-item '+(isOpen?'open':'')} key={s.id}>
    <button className="help-item-head" onClick={()=>setOpen(isOpen?'':s.id)}>
     <div className="help-item-icon"><I size={19}/></div>
     <div><b>{s.title}</b><span>{s.desc}</span></div>
     <ChevronDown size={18} className="help-chevron"/>
    </button>
    {isOpen&&<div className="help-item-body">{s.content}</div>}
   </div>})}
   {!filtered.length&&<div className="help-no-result"><Search size={25}/><b>Nenhum tópico encontrado</b><span>Tente pesquisar usando outra palavra.</span></div>}
  </div>
 </div>
}

function UsersPage({data,setData}){return <div className="panel"><Head title="Usuários e permissões" sub="Base para rastreabilidade e acesso" action={<button className="primary" onClick={()=>{let n=prompt('Nome do usuário:');if(n)setData(d=>({...d,users:[...d.users,{id:Date.now(),name:n,role:'Estoque / Cirurgias',active:true}]}))}}><Plus size={15}/>Novo usuário</button>}/><table><thead><tr><th>Usuário</th><th>Perfil</th><th>Status</th></tr></thead><tbody>{data.users.map(u=><tr key={u.id}><td><b>{u.name}</b></td><td>{u.role}</td><td><button className="statusbtn" onClick={()=>setData(d=>({...d,users:d.users.map(x=>x.id===u.id?{...x,active:!x.active}:x)}))}><Badge tone={u.active?'green':'red'}>{u.active?'Ativo':'Inativo'}</Badge></button></td></tr>)}</tbody></table></div>}
const SettingsPage=({setData})=><div className="panel"><Head title="Configurações gerais" sub="Parâmetros e manutenção"/><div className="settings"><div><b>Estoque mínimo</b><span>Definido individualmente por lente.</span></div><div><b>Alerta de validade</b><span>180 dias antes do vencimento.</span></div><div><b>Auditoria</b><span>Data, horário e responsável em cada movimentação.</span></div><button className="dangerBtn" onClick={()=>confirm('Restaurar dados de demonstração?')&&setData(seed)}><Trash2 size={15}/>Restaurar demonstração</button></div></div>;
function LensModal({item,onClose,onSave}){
 const[f,setF]=useState(item||{name:'',manufacturer:'',model:'',power:'',type:'Monofocal',lot:'',serial:'',barcode:'',qty:1,min:2,expiry:'',location:'',supplier:'',status:'Disponível'}),s=(k,v)=>setF({...f,[k]:v});
 return <Modal title={item?'Editar lente':'Cadastrar nova lente'} onClose={onClose} wide>
  <div className="formgrid">
   <Field label="Nome da lente *"><input value={f.name} onChange={e=>s('name',e.target.value)}/></Field>
   <Field label="Fabricante"><input value={f.manufacturer} onChange={e=>s('manufacturer',e.target.value)}/></Field>
   <Field label="Modelo"><input value={f.model} onChange={e=>s('model',e.target.value)}/></Field>
   <Field label="Potência / Numeração *"><input placeholder="+21.0 D" value={f.power} onChange={e=>s('power',e.target.value)}/></Field>
   <Field label="Código de barras / QR"><input autoFocus={!item} placeholder="Escaneie ou digite o código" value={f.barcode||''} onChange={e=>s('barcode',e.target.value)}/></Field>
   <Field label="Tipo"><select value={f.type} onChange={e=>s('type',e.target.value)}><option>Monofocal</option><option>Monofocal Plus</option><option>Tórica</option><option>EDOF</option><option>Multifocal</option><option>Trifocal</option></select></Field>
   <Field label="Lote"><input value={f.lot} onChange={e=>s('lot',e.target.value)}/></Field>
   <Field label="Número de série"><input value={f.serial} onChange={e=>s('serial',e.target.value)}/></Field>
   <Field label="Validade"><input type="date" value={f.expiry} onChange={e=>s('expiry',e.target.value)}/></Field>
   <Field label="Quantidade"><input type="number" value={f.qty} onChange={e=>s('qty',e.target.value)}/></Field>
   <Field label="Estoque mínimo"><input type="number" value={f.min} onChange={e=>s('min',e.target.value)}/></Field>
   <Field label="Localização"><input value={f.location} onChange={e=>s('location',e.target.value)}/></Field>
   <Field label="Fornecedor"><input value={f.supplier} onChange={e=>s('supplier',e.target.value)}/></Field>
  </div>
  <div className="modal-actions"><button className="secondary" onClick={onClose}>Cancelar</button><button className="primary" disabled={!f.name||!f.power} onClick={()=>onSave(f)}><Save size={15}/>Salvar lente</button></div>
 </Modal>
}
function MovementModal({lenses,currentUserName,onClose,onSave}){
 let n=now();
 const[f,setF]=useState({type:'Entrada',lensId:lenses[0]?.id||'',qty:1,date:n.date,time:n.time,user:currentUserName||'Usuário',reason:'Compra / reposição',patient:'',surgery:''});
 const[unknownCode,setUnknownCode]=useState('');
 const s=(k,v)=>setF(x=>({...x,[k]:v}));
 const selectedLens=lenses.find(l=>String(l.id)===String(f.lensId));

 const selectScannedLens=(lens)=>{
  setUnknownCode('');
  setF(x=>({...x,lensId:lens.id}));
 };

 return <Modal title="Registrar movimentação" onClose={onClose} wide>
  <div className="formone">
   <LensScanner
    lenses={lenses}
    title="Escanear lente"
    onSelect={selectScannedLens}
    onUnknown={code=>setUnknownCode(code)}
   />

   {unknownCode&&<div className="help-tip" style={{marginTop:10}}>
    <AlertTriangle size={17}/>
    <span>O código <b>{unknownCode}</b> não está vinculado a nenhuma lente. Cadastre o código na lente antes de registrar a movimentação.</span>
   </div>}

   {selectedLens&&<div style={{padding:'10px 12px',border:'1px solid #dfe8e7',borderRadius:10,background:'#f7faf9'}}>
    <b>{selectedLens.name} • {selectedLens.power}</b>
    <div style={{fontSize:12,color:'#6d7b7d',marginTop:3}}>Lote {selectedLens.lot||'—'} • Código {selectedLens.barcode||'não cadastrado'} • Saldo {selectedLens.qty}</div>
   </div>}

   <Field label="Tipo"><select value={f.type} onChange={e=>s('type',e.target.value)}><option>Entrada</option><option>Saída</option><option>Reserva</option><option>Devolução</option><option>Ajuste de inventário</option></select></Field>
   <Field label="Lente / Numeração"><select value={f.lensId} onChange={e=>s('lensId',e.target.value)}>{lenses.map(l=><option key={l.id} value={l.id}>{l.name} • {l.power} • lote {l.lot} • saldo {l.qty}</option>)}</select></Field>
   <div className="formgrid nopad">
    <Field label="Quantidade"><input type="number" min="1" value={f.qty} onChange={e=>s('qty',e.target.value)}/></Field>
    <Field label="Responsável"><input value={currentUserName||f.user} readOnly/></Field>
    <Field label="Data"><input type="date" value={f.date} onChange={e=>s('date',e.target.value)}/></Field>
    <Field label="Horário"><input type="time" value={f.time} onChange={e=>s('time',e.target.value)}/></Field>
   </div>
   <Field label="Motivo / Observação"><input value={f.reason} onChange={e=>s('reason',e.target.value)}/></Field>
   <div className="formgrid nopad">
    <Field label="Paciente (opcional)"><input value={f.patient} onChange={e=>s('patient',e.target.value)}/></Field>
    <Field label="Cirurgia (opcional)"><input value={f.surgery} onChange={e=>s('surgery',e.target.value)}/></Field>
   </div>
  </div>
  <div className="modal-actions"><button className="secondary" onClick={onClose}>Cancelar</button><button className="primary" disabled={!f.lensId||Number(f.qty)<1} onClick={()=>onSave(f)}><Save size={15}/>Registrar</button></div>
 </Modal>
}
function ReservationModal({item,lenses,onClose,onSave}){const[f,setF]=useState(item||{patient:'',surgery:'Catarata',doctor:'Dr. Daniel',date:'',time:'13:30',lens:lenses[0]?.name||'',power:lenses[0]?.power||'',eye:'OD',status:'Reservada'}),s=(k,v)=>setF({...f,[k]:v});return <Modal title={item?'Editar reserva':'Nova reserva cirúrgica'} onClose={onClose}><div className="formone"><Field label="Paciente *"><input value={f.patient} onChange={e=>s('patient',e.target.value)}/></Field><div className="formgrid nopad"><Field label="Cirurgia"><input value={f.surgery} onChange={e=>s('surgery',e.target.value)}/></Field><Field label="Médico"><input value={f.doctor} onChange={e=>s('doctor',e.target.value)}/></Field><Field label="Data"><input type="date" value={f.date} onChange={e=>s('date',e.target.value)}/></Field><Field label="Horário"><input type="time" value={f.time} onChange={e=>s('time',e.target.value)}/></Field></div><Field label="Lente"><select value={f.lens+'|'+f.power} onChange={e=>{let[a,b]=e.target.value.split('|');setF({...f,lens:a,power:b})}}>{lenses.map(l=><option value={l.name+'|'+l.power}>{l.name} • {l.power} • saldo {l.qty}</option>)}</select></Field><div className="formgrid nopad"><Field label="Olho"><select value={f.eye} onChange={e=>s('eye',e.target.value)}><option>OD</option><option>OE</option><option>AO</option></select></Field><Field label="Status"><select value={f.status} onChange={e=>s('status',e.target.value)}><option>Reservada</option><option>Confirmada</option><option>Realizada</option><option>Cancelada</option></select></Field></div></div><div className="modal-actions"><button className="secondary" onClick={onClose}>Cancelar</button><button className="primary" disabled={!f.patient||!f.date} onClick={()=>onSave(f)}><Save size={15}/>Salvar reserva</button></div></Modal>}

export default function App(){
 const[session,setSession]=useState(undefined);
 const[profile,setProfile]=useState(undefined);
 const[authError,setAuthError]=useState('');

 useEffect(()=>{
  let mounted=true;

  const applySession=async(newSession)=>{
   if(!mounted)return;
   setSession(newSession||null);
   setAuthError('');

   if(!newSession){
    setProfile(null);
    return;
   }

   setProfile(undefined);

   const{data:p,error}=await supabase
    .from('profiles')
    .select('id,full_name,role,active')
    .eq('id',newSession.user.id)
    .single();

   if(!mounted)return;

   if(error||!p){
    console.error('Erro ao carregar perfil:',error);
    setAuthError('Não foi possível carregar o perfil deste usuário.');
    setProfile(null);
    return;
   }

   if(!p.active){
    setAuthError('Este usuário está inativo. Procure o administrador do sistema.');
    setProfile(null);
    await supabase.auth.signOut();
    return;
   }

   setProfile(p);
  };

  supabase.auth.getSession().then(({data,error})=>{
   if(error){
    console.error('Erro ao verificar sessão:',error);
    if(mounted){
     setSession(null);
     setProfile(null);
    }
    return;
   }
   applySession(data?.session||null);
  });

  const{data:{subscription}}=supabase.auth.onAuthStateChange((_event,newSession)=>{
   applySession(newSession||null);
  });

  return()=>{
   mounted=false;
   subscription.unsubscribe();
  };
 },[]);

 if(session===undefined||(session&&profile===undefined)){
  return <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',background:'#f5f7f7',color:'#285b64',fontFamily:'inherit'}}>Carregando sistema...</div>;
 }

 if(!session)return <Login/>;

 if(authError||!profile){
  return <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',background:'#f5f7f7',padding:24}}><div style={{maxWidth:430,width:'100%',background:'#fff',border:'1px solid #e1e8e7',borderRadius:16,padding:28,textAlign:'center'}}><h2 style={{marginTop:0,color:'#18363d'}}>Acesso não liberado</h2><p style={{color:'#68787a'}}>{authError||'Perfil de acesso não encontrado.'}</p><button className="primary" onClick={()=>supabase.auth.signOut()}>Voltar ao login</button></div></div>;
 }

 return <Sistema profile={profile}/>;
}

