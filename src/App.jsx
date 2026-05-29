import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend } from "recharts";
import { supabase, fetchPedidos, upsertPedidos, updatePedidoEnviado, updatePedidoStatus, fetchFaturamento, upsertFaturamento } from "./supabase.js";

// ── Column mappings for new spreadsheet format ────────────────────────────────
const NEW_COL_MAP = {
  "ID do Pedido":                 "idPedido",
  "Status do Pedido":             "statusPedido",
  "Data Prevista de Envio":       "dataEnvio",
  "Hora do Pagamento do Pedido":  "horaPagamento",
  "Nome do Produto":              "produto",
  "Preço Acordado":               "preco",
  "Quantidade":                   "quantidade",
  "Nome da Variação":             "variacao",
  "Nome do Destinatário":         "destinatario",
};

// ── Parse new spreadsheet ─────────────────────────────────────────────────────
function parseNewSheet(sheet, loja) {
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return raw.map(row => {
    const m = { loja, statusInterno: "", notaRevisao: "", _status: "new" };
    for (const [orig, key] of Object.entries(NEW_COL_MAP)) {
      const found = Object.keys(row).find(k => k.trim().toLowerCase() === orig.trim().toLowerCase());
      m[key] = found ? String(row[found]) : "";
    }
    // normalize preco to number
    m.preco = parseFloat(String(m.preco).replace(/[^\d.,]/g,"").replace(",",".")) || 0;
    return m;
  }).filter(r => r.idPedido);
}

// ── Deadline color ────────────────────────────────────────────────────────────
function deadlineInfo(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr.replace(" ","T"));
  if (isNaN(d)) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  const dd = new Date(d); dd.setHours(0,0,0,0);
  const diff = Math.round((dd-today)/86400000);
  if (diff<=1) return {bg:"#fde8e8",border:"#ef4444",text:"#991b1b",icon:"🔴",label:diff<=0?"Vencido!":"Amanhã",tier:"red"};
  if (diff===2) return {bg:"#fef9c3",border:"#f59e0b",text:"#78350f",icon:"🟡",label:`${diff}d`,tier:"yellow"};
  return {bg:"#dcfce7",border:"#22c55e",text:"#14532d",icon:"🟢",label:`${diff}d`,tier:"green"};
}

function fmtBRL(v) {
  return Number(v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
}

function currentMes() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
}

// ── DropZone ──────────────────────────────────────────────────────────────────
function DropZone({label, sublabel, onFile, file, color, disabled}) {
  const ref = useRef();
  const [drag, setDrag] = useState(false);
  const handle = f => {
    if (!f) return;
    const r = new FileReader();
    r.onload = e => { const wb=XLSX.read(e.target.result,{type:"binary"}); onFile(wb.Sheets[wb.SheetNames[0]], f.name); };
    r.readAsBinaryString(f);
  };
  return (
    <div onClick={()=>!disabled&&ref.current.click()}
      onDragOver={e=>{if(!disabled){e.preventDefault();setDrag(true);}}}
      onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);if(!disabled)handle(e.dataTransfer.files[0]);}}
      style={{border:`2px dashed ${drag?color:disabled?"#e2e8f0":"#cbd5e1"}`,borderRadius:12,padding:"18px 14px",cursor:disabled?"not-allowed":"pointer",background:disabled?"#f8fafc":drag?`${color}18`:"#f8fafc",textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:5,opacity:disabled?0.5:1,transition:"all 0.2s",minHeight:110}}>
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>handle(e.target.files[0])} />
      <div style={{fontSize:24}}>{disabled?"🔒":"📋"}</div>
      <div style={{fontWeight:700,fontSize:12,color:"#1e293b"}}>{label}</div>
      {sublabel&&<div style={{fontSize:10,color:"#94a3b8",maxWidth:160}}>{sublabel}</div>}
      {file?<div style={{background:color,color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:10,fontWeight:600,marginTop:2}}>✓ {file.length>22?file.slice(0,22)+"...":file}</div>
           :<div style={{fontSize:10,color:"#94a3b8"}}>Arraste ou clique</div>}
    </div>
  );
}


function StyledCheckbox({checked, onChange}) {
  return (
    <div onClick={onChange} style={{
      width:18, height:18, borderRadius:5, border:`2px solid ${checked?"#1d4ed8":"#cbd5e1"}`,
      background: checked?"#1d4ed8":"#fff",
      display:"flex", alignItems:"center", justifyContent:"center",
      cursor:"pointer", transition:"all 0.15s", flexShrink:0,
    }}>
      {checked && <svg width="10" height="8" viewBox="0 0 10 8" fill="none"><path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>}
    </div>
  );
}

function ExpandCell({value}) {
  const [expanded, setExpanded] = useState(false);
  const isLong = value && value.length > 18;
  return (
    <div style={{display:"flex", alignItems:"flex-start", gap:4, maxWidth:160}}>
      <span style={{
        fontSize:12, color:"#374151", lineHeight:1.4,
        whiteSpace: expanded?"normal":"nowrap",
        overflow: expanded?"visible":"hidden",
        textOverflow: expanded?"clip":"ellipsis",
        flex:1,
      }}>{value}</span>
      {isLong && (
        <button onClick={()=>setExpanded(v=>!v)} style={{
          background:"none", border:"none", cursor:"pointer",
          color:"#94a3b8", fontSize:14, padding:0, flexShrink:0, lineHeight:1,
        }} title={expanded?"Recolher":"Expandir"}>
          {expanded?"▲":"▼"}
        </button>
      )}
    </div>
  );
}

function Toast({msg,color}) {
  if(!msg) return null;
  return <div style={{position:"fixed",bottom:24,right:24,zIndex:9999,background:color||"#1e293b",color:"#fff",borderRadius:12,padding:"12px 20px",fontSize:13,fontWeight:600,boxShadow:"0 8px 32px rgba(0,0,0,0.2)"}}>{msg}</div>;
}

function RevisaoModal({order, onConfirm, onClose}) {
  const [nota, setNota] = useState(order.notaRevisao||"");
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
        <h3 style={{margin:"0 0 6px",fontSize:15,fontWeight:700,color:"#1e293b"}}>📋 Enviar para Revisão</h3>
        <p style={{margin:"0 0 14px",fontSize:12,color:"#64748b"}}>Pedido: <strong>{order.idPedido}</strong></p>
        <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Motivo *</label>
        <textarea value={nota} onChange={e=>setNota(e.target.value)} placeholder="Descreva o motivo da revisão..." autoFocus
          style={{width:"100%",minHeight:90,padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,resize:"vertical",outline:"none",boxSizing:"border-box",fontFamily:"inherit"}} />
        <div style={{display:"flex",gap:10,marginTop:14,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"8px 18px",border:"1px solid #e2e8f0",borderRadius:10,background:"#fff",fontSize:12,cursor:"pointer",color:"#374151"}}>Cancelar</button>
          <button onClick={()=>onConfirm(nota)} disabled={!nota.trim()} style={{padding:"8px 18px",border:"none",borderRadius:10,background:nota.trim()?"#f59e0b":"#e2e8f0",color:nota.trim()?"#fff":"#94a3b8",fontSize:12,fontWeight:700,cursor:nota.trim()?"pointer":"not-allowed"}}>Confirmar</button>
        </div>
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
const LOJAS = ["Gran Shop", "Aishael Mix"];

export default function App({user, onLogout}) {
  const [orders,       setOrders]       = useState([]);
  const [enviados,     setEnviados]     = useState([]);
  const [faturamento,  setFaturamento]  = useState([]);
  const [toast,        setToast]        = useState(null);
  const [dbLoading,    setDbLoading]    = useState(true);
  const [saving,       setSaving]       = useState(false);
  const [search,       setSearch]       = useState("");
  const [filterLoja,   setFilterLoja]   = useState("all");
  const [filterPrazo,  setFilterPrazo]  = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterProduto,setFilterProduto]= useState("all");
  const [showProdPanel,setShowProdPanel]= useState(false);
  const [showEnviados, setShowEnviados] = useState(false);
  const [revisaoModal, setRevisaoModal] = useState(null);
  const [selected,     setSelected]     = useState(new Set());
  const [copied,       setCopied]       = useState(false);
  const [uploadNames,  setUploadNames]  = useState({});
  const tableRef = useRef(null);
  const showToast = (msg,color,ms=3500)=>{setToast({msg,color});setTimeout(()=>setToast(null),ms);};

  useEffect(()=>{
    if(!supabase||!user){setDbLoading(false);return;}
    Promise.all([fetchPedidos(user.id), fetchFaturamento(user.id)]).then(([rows,fat])=>{
      setOrders(rows.filter(r=>r.statusPedido!=="ENVIADO"));
      setEnviados(rows.filter(r=>r.statusPedido==="ENVIADO"));
      setFaturamento(fat);
      setDbLoading(false);
    });
  },[user]);

  // ── Load new orders ────────────────────────────────────────────────────────
  const handleAddPedidos = async (sheet, name, loja) => {
    setUploadNames(p=>({...p,[`add_${loja}`]:name}));
    const incoming = parseNewSheet(sheet, loja).map(r=>({...r, statusPedido:"A ENVIAR"}));
    const existingIds = new Set(orders.map(r=>r.idPedido));
    const newRows = incoming.filter(r=>!existingIds.has(r.idPedido));
    const ignored = incoming.length - newRows.length;
    if(newRows.length===0){showToast(`ℹ️ Todos os ${ignored} pedidos já existem.`,"#d97706");return;}
    const updated = [...orders, ...newRows];
    setOrders(updated);
    showToast(`✅ +${newRows.length} pedidos adicionados (${ignored} ignorados)`,"#059669");
    // save to DB + faturamento
    if(supabase){
      setSaving(true);
      await upsertPedidos(newRows, user.id);
      const mes = currentMes();
      const totalValor = newRows.reduce((s,r)=>s+(parseFloat(r.preco)||0),0);
      if(totalValor>0) await upsertFaturamento(user.id, mes, loja, totalValor);
      const fat = await fetchFaturamento(user.id);
      setFaturamento(fat);
      setSaving(false);
    }
  };

  // ── Load final sheet (mark as ENVIADO) ────────────────────────────────────
  const handleFinalSheet = async (sheet, name, loja) => {
    setUploadNames(p=>({...p,[`final_${loja}`]:name}));
    const incoming = parseNewSheet(sheet, loja);
    const enviadoIds = new Set(incoming.filter(r=>r.statusPedido&&r.statusPedido.toUpperCase().includes("ENVIADO")).map(r=>r.idPedido));
    // also treat any non-"A ENVIAR" rows as enviado
    incoming.forEach(r=>{ if(r.statusPedido&&!r.statusPedido.toUpperCase().includes("A ENVIAR")) enviadoIds.add(r.idPedido); });
    if(enviadoIds.size===0){showToast("ℹ️ Nenhum pedido 'ENVIADO' identificado.","#d97706");return;}
    const nowEnviados = orders.filter(r=>enviadoIds.has(r.idPedido)).map(r=>({...r,statusPedido:"ENVIADO",_status:"enviado"}));
    const remaining  = orders.filter(r=>!enviadoIds.has(r.idPedido));
    setOrders(remaining);
    setEnviados(prev=>{
      const existIds = new Set(prev.map(e=>e.idPedido));
      return [...prev, ...nowEnviados.filter(e=>!existIds.has(e.idPedido))];
    });
    showToast(`📦 ${nowEnviados.length} pedido(s) marcados como ENVIADO`,"#059669");
    if(supabase){setSaving(true); await updatePedidoEnviado([...enviadoIds], user.id); setSaving(false);}
  };

  // ── Status: revisao / feito ───────────────────────────────────────────────
  const handleStatusChange = async (order, newStatus, nota="") => {
    const upd = orders.map(o=>o.idPedido===order.idPedido?{...o,statusInterno:newStatus,notaRevisao:nota}:o);
    setOrders(upd);
    if(supabase) await updatePedidoStatus(order.idPedido, user.id, newStatus, nota);
    showToast(newStatus==="feito"?"✅ Marcado como Feito!":"📋 Enviado para Revisão", newStatus==="feito"?"#059669":"#f59e0b");
  };

  // ── Selection ─────────────────────────────────────────────────────────────
  const toggleSelect = id => setSelected(prev=>{ const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s; });
  const toggleAll = () => selected.size===filtered.length?setSelected(new Set()):setSelected(new Set(filtered.map(r=>r.idPedido)));
  const copySelected = () => {
    navigator.clipboard.writeText([...selected].join(",")).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);showToast("📋 IDs copiados para o UpSeller!","#1d4ed8");});
  };

  // ── Filters ───────────────────────────────────────────────────────────────
  const activeOrders = orders.filter(r=>r._status!=="removed");
  const produtoMap = activeOrders.reduce((acc,r)=>{ if(r.produto)acc[r.produto]=(acc[r.produto]||0)+1; return acc; },{});
  const produtos = Object.entries(produtoMap).sort((a,b)=>b[1]-a[1]);

  const filtered = activeOrders.filter(r=>{
    const dl = deadlineInfo(r.dataEnvio);
    const q = search.toLowerCase(); const matchSearch = !search || r.idPedido.toLowerCase().includes(q)||r.produto.toLowerCase().includes(q)||r.destinatario.toLowerCase().includes(q)||r.loja.toLowerCase().includes(q)||r.variacao.toLowerCase().includes(q);
    return (
      (filterLoja==="all"   ||r.loja===filterLoja)&&
      (filterPrazo==="all"  ||(dl&&dl.tier===filterPrazo))&&
      (filterStatus==="all" ||r.statusInterno===filterStatus)&&
      (filterProduto==="all"||r.produto===filterProduto)&&
      matchSearch
    );
  }).sort((a,b)=>{
    // sort by deadline: expired/closest first, no deadline last
    const getDay = r => {
      if(!r.dataEnvio) return 9999;
      const d = new Date(r.dataEnvio.replace(" ","T"));
      if(isNaN(d)) return 9999;
      const today = new Date(); today.setHours(0,0,0,0);
      const dd = new Date(d); dd.setHours(0,0,0,0);
      return Math.round((dd-today)/86400000);
    };
    return getDay(a)-getDay(b);
  });

  const urgentCount  = activeOrders.filter(r=>deadlineInfo(r.dataEnvio)?.tier==="red").length;
  const feitoCount   = activeOrders.filter(r=>r.statusInterno==="feito").length;
  const revisaoCount = activeOrders.filter(r=>r.statusInterno==="revisao").length;
  const totalPreco   = activeOrders.reduce((s,r)=>s+(parseFloat(r.preco)||0),0);

  // ── Chart data ─────────────────────────────────────────────────────────────
  // Build chart data with faturamento + ticket médio
  // count pedidos per month+loja from orders+enviados for ticket calc
  const allOrders = [...orders, ...enviados];
  const pedidosByMesLoja = allOrders.reduce((acc,r)=>{
    if(!r.horaPagamento && !r.dataEnvio) return acc;
    const raw = r.horaPagamento || r.dataEnvio || "";
    const mes = raw.slice(0,7); // "YYYY-MM"
    if(!mes || mes.length<7) return acc;
    const key = `${mes}__${r.loja}`;
    acc[key] = (acc[key]||0)+1;
    return acc;
  },{});

  const chartData = faturamento.reduce((acc,f)=>{
    const ex = acc.find(a=>a.mes===f.mes);
    const count = pedidosByMesLoja[`${f.mes}__${f.loja}`] || 1;
    const ticket = Number(f.valor) / count;
    if(ex){
      ex[f.loja]=(ex[f.loja]||0)+Number(f.valor);
      ex[`ticket_${f.loja}`]=ticket;
    } else {
      const n={mes:f.mes};
      n[f.loja]=Number(f.valor);
      n[`ticket_${f.loja}`]=ticket;
      acc.push(n);
    }
    return acc;
  },[]).sort((a,b)=>a.mes.localeCompare(b.mes));

  if(dbLoading) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#1e3a8a,#3b82f6)",fontFamily:"sans-serif"}}>
      <div style={{color:"#fff",fontSize:16,fontWeight:600}}>📦 Carregando seus pedidos...</div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#f0f4ff",fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
      <style>{`*{box-sizing:border-box} @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}`}</style>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#1e3a8a 0%,#1d4ed8 60%,#3b82f6 100%)",padding:"20px 28px",color:"#fff"}}>
        <div style={{maxWidth:1600,margin:"0 auto",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:10,fontWeight:600,letterSpacing:2,opacity:0.7}}>GERENCIADOR DE PEDIDOS</div>
            <h1 style={{margin:"2px 0 0",fontSize:22,fontWeight:800}}>Painel de Pedidos 📦</h1>
          </div>
          <span style={{background:"rgba(255,255,255,0.15)",borderRadius:20,padding:"4px 14px",fontSize:12,display:"flex",alignItems:"center",gap:6}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:saving?"#fbbf24":"#4ade80",display:"inline-block"}}/>
            {saving?"Salvando...":"Conectado"}
          </span>
          <span style={{background:"rgba(255,255,255,0.12)",borderRadius:20,padding:"4px 14px",fontSize:12}}>👤 {user?.email}</span>
          <button onClick={onLogout} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"4px 16px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Sair →</button>
        </div>
      </div>

      <div style={{maxWidth:1600,margin:"0 auto",padding:"20px 28px"}}>

        {/* ── Upload Section ── */}
        <div style={{background:"#fff",borderRadius:18,padding:20,marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
          <div style={{fontSize:12,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:0.5,marginBottom:14}}>Carregar Planilhas</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
            {LOJAS.map(loja=>(
              <div key={`add_${loja}`} style={{border:"1px solid #e2e8f0",borderRadius:12,padding:12}}>
                <div style={{fontSize:11,fontWeight:700,color:"#1d4ed8",marginBottom:8,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{background:"#eff6ff",borderRadius:6,padding:"2px 8px"}}>{loja}</span>
                  <span style={{color:"#64748b",fontWeight:500}}>· Novos Pedidos</span>
                </div>
                <DropZone label="Adicionar Pedidos" sublabel="Pedidos com status A ENVIAR" color="#1d4ed8"
                  file={uploadNames[`add_${loja}`]}
                  onFile={(sheet,name)=>handleAddPedidos(sheet,name,loja)} />
              </div>
            ))}
            {LOJAS.map(loja=>(
              <div key={`final_${loja}`} style={{border:"1px solid #e2e8f0",borderRadius:12,padding:12}}>
                <div style={{fontSize:11,fontWeight:700,color:"#059669",marginBottom:8,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{background:"#f0fdf4",borderRadius:6,padding:"2px 8px"}}>{loja}</span>
                  <span style={{color:"#64748b",fontWeight:500}}>· Planilha Final</span>
                </div>
                <DropZone label="Marcar Enviados" sublabel="Pedidos ENVIADO saem da lista principal" color="#059669"
                  file={uploadNames[`final_${loja}`]}
                  onFile={(sheet,name)=>handleFinalSheet(sheet,name,loja)} />
              </div>
            ))}
          </div>
        </div>

        {/* ── Stats ── */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:12,marginBottom:18}}>
          {[
            {label:"Total Pedidos",    value:activeOrders.length,        color:"#1d4ed8", icon:"📦", action:()=>{setFilterStatus("all");setFilterPrazo("all");setFilterProduto("all");setFilterLoja("all");setSearch("");}},
            {label:"Feitos",           value:feitoCount,                  color:"#059669", icon:"✅", action:()=>setFilterStatus("feito")},
            {label:"Em Revisão",       value:revisaoCount,                color:"#f59e0b", icon:"📋", action:()=>setFilterStatus("revisao")},
            {label:"Urgentes (≤1d)",   value:urgentCount,                 color:"#ef4444", icon:"🔴", action:()=>{setFilterStatus("all");setFilterPrazo("red");}},
            {label:"Enviados",         value:enviados.length,             color:"#7c3aed", icon:"📬", action:()=>setShowEnviados(v=>!v)},
            {label:"Valor em Aberto",  value:fmtBRL(totalPreco),          color:"#0891b2", icon:"💰", action:()=>{}},
          ].map(s=>(
            <div key={s.label} onClick={()=>{s.action();tableRef.current?.scrollIntoView({behavior:"smooth"});}}
              style={{background:"#fff",borderRadius:14,padding:"14px 16px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",borderLeft:`4px solid ${s.color}`,cursor:"pointer",transition:"transform 0.15s,box-shadow 0.15s"}}
              onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 6px 20px rgba(0,0,0,0.12)";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,0.07)";}}>
              <div style={{fontSize:18,marginBottom:2}}>{s.icon}</div>
              <div style={{fontSize:s.label==="Valor em Aberto"?16:22,fontWeight:800,color:s.color,lineHeight:1.2}}>{s.value}</div>
              <div style={{fontSize:10,color:"#64748b",fontWeight:500,marginTop:2}}>{s.label}</div>
              <div style={{fontSize:9,color:s.color,opacity:0.7,marginTop:2}}>clique para filtrar →</div>
            </div>
          ))}
        </div>

        {/* ── Faturamento Chart ── */}
        {chartData.length>0 && (
          <div style={{background:"#fff",borderRadius:18,padding:20,marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
              <div style={{fontSize:14,fontWeight:700,color:"#1e293b"}}>💰 Faturamento Mensal por Loja</div>
              <div style={{marginLeft:"auto",fontSize:12,color:"#64748b"}}>Total acumulado: <strong style={{color:"#0891b2"}}>{fmtBRL(faturamento.reduce((s,f)=>s+Number(f.valor),0))}</strong></div>
            </div>
            {/* Ticket médio summary cards */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
              {LOJAS.map(loja=>{
                const totalFat = faturamento.filter(f=>f.loja===loja).reduce((s,f)=>s+Number(f.valor),0);
                const totalPed = allOrders.filter(r=>r.loja===loja).length || 1;
                const ticket = totalFat / totalPed;
                const lastMes = chartData[chartData.length-1];
                const mesFat = lastMes?.[loja] || 0;
                return (
                  <div key={loja} style={{background:"#f8fafc",borderRadius:12,padding:"12px 14px",border:"1px solid #e2e8f0"}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:6,textTransform:"uppercase",letterSpacing:0.5}}>{loja}</div>
                    <div style={{display:"flex",flexDirection:"column",gap:4}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontSize:11,color:"#94a3b8"}}>Total acum.</span>
                        <span style={{fontSize:13,fontWeight:800,color:"#1d4ed8"}}>{fmtBRL(totalFat)}</span>
                      </div>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontSize:11,color:"#94a3b8"}}>Pedidos</span>
                        <span style={{fontSize:13,fontWeight:700,color:"#374151"}}>{allOrders.filter(r=>r.loja===loja).length}</span>
                      </div>
                      <div style={{height:1,background:"#e2e8f0",margin:"2px 0"}} />
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontSize:11,color:"#94a3b8"}}>🎯 Ticket Médio</span>
                        <span style={{fontSize:13,fontWeight:800,color:"#059669"}}>{fmtBRL(ticket)}</span>
                      </div>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontSize:11,color:"#94a3b8"}}>Mês atual</span>
                        <span style={{fontSize:12,fontWeight:700,color:"#7c3aed"}}>{fmtBRL(mesFat)}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
              <div style={{background:"#eff6ff",borderRadius:12,padding:"12px 14px",border:"1px solid #bfdbfe"}}>
                <div style={{fontSize:20,marginBottom:4}}>💰</div>
                <div style={{fontSize:18,fontWeight:800,color:"#0891b2"}}>{fmtBRL(faturamento.reduce((s,f)=>s+Number(f.valor),0))}</div>
                <div style={{fontSize:11,color:"#64748b",fontWeight:600,marginTop:2}}>Total Geral</div>
              </div>
              <div style={{background:"#f0fdf4",borderRadius:12,padding:"12px 14px",border:"1px solid #86efac"}}>
                <div style={{fontSize:20,marginBottom:4}}>🎯</div>
                <div style={{fontSize:18,fontWeight:800,color:"#059669"}}>{fmtBRL(faturamento.reduce((s,f)=>s+Number(f.valor),0)/(allOrders.length||1))}</div>
                <div style={{fontSize:11,color:"#64748b",fontWeight:600,marginTop:2}}>Ticket Médio Geral</div>
              </div>
            </div>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={chartData} margin={{top:0,right:10,left:10,bottom:0}}>
                <XAxis dataKey="mes" tick={{fontSize:11}} />
                <YAxis yAxisId="fat" tick={{fontSize:11}} tickFormatter={v=>v>=1000?`R$${(v/1000).toFixed(1)}k`:`R$${v}`} />
                <YAxis yAxisId="ticket" orientation="right" tick={{fontSize:11}} tickFormatter={v=>`R$${v.toFixed(0)}`} />
                <Tooltip formatter={(v,n)=>{
                  if(n.startsWith("Ticket")) return [fmtBRL(v), n];
                  return [fmtBRL(v), n];
                }} />
                <Legend />
                <Bar yAxisId="fat"    dataKey="Gran Shop"              fill="#1d4ed8" radius={[4,4,0,0]} name="Gran Shop" />
                <Bar yAxisId="fat"    dataKey="Aishael Mix"            fill="#7c3aed" radius={[4,4,0,0]} name="Aishael Mix" />
                <Bar yAxisId="ticket" dataKey="ticket_Gran Shop"       fill="#93c5fd" radius={[4,4,0,0]} name="Ticket Gran Shop" opacity={0.8} />
                <Bar yAxisId="ticket" dataKey="ticket_Aishael Mix"     fill="#c4b5fd" radius={[4,4,0,0]} name="Ticket Aishael Mix" opacity={0.8} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}

        {/* ── Produto Panel ── */}
        <div style={{background:"#fff",borderRadius:18,marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
          <button onClick={()=>setShowProdPanel(v=>!v)}
            style={{width:"100%",padding:"14px 20px",background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",gap:10,textAlign:"left"}}>
            <span style={{fontSize:14,fontWeight:700,color:"#1e293b"}}>📦 Filtrar por Produto</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{produtos.length} produto{produtos.length!==1?"s":""}</span>
            {filterProduto!=="all" && <span style={{background:"#7c3aed",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>Ativo: {filterProduto.slice(0,30)}...</span>}
            <span style={{marginLeft:"auto",color:"#94a3b8",fontSize:16}}>{showProdPanel?"▲":"▼"}</span>
          </button>
          {showProdPanel && (
            <div style={{borderTop:"1px solid #f1f5f9",padding:"14px 20px"}}>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom: filterProduto!=="all"?16:0}}>
                {filterProduto!=="all" && (
                  <button onClick={()=>setFilterProduto("all")} style={{background:"#f1f5f9",color:"#374151",border:"1px solid #e2e8f0",borderRadius:10,padding:"6px 12px",fontSize:11,cursor:"pointer",fontWeight:600}}>✕ Limpar filtro</button>
                )}
                {produtos.map(([nome,qtd])=>{
                  const isActive=filterProduto===nome;
                  return (
                    <button key={nome} onClick={()=>{setFilterProduto(isActive?"all":nome);tableRef.current?.scrollIntoView({behavior:"smooth"});}} title={nome}
                      style={{display:"flex",alignItems:"center",gap:6,background:isActive?"#7c3aed":"#f8fafc",color:isActive?"#fff":"#374151",border:`1.5px solid ${isActive?"#7c3aed":"#e2e8f0"}`,borderRadius:10,padding:"6px 12px",fontSize:12,fontWeight:isActive?700:500,cursor:"pointer",transition:"all 0.15s"}}>
                      <span style={{background:isActive?"rgba(255,255,255,0.25)":"#7c3aed",color:"#fff",borderRadius:20,padding:"1px 7px",fontSize:11,fontWeight:700,flexShrink:0}}>{qtd}</span>
                      {nome.length>50?nome.slice(0,50)+"...":nome}
                    </button>
                  );
                })}
              </div>
              {/* Product order list */}
              {filterProduto!=="all" && (
                <div>
                  <div style={{fontSize:12,fontWeight:700,color:"#5b21b6",marginBottom:10}}>
                    Pedidos de: <span style={{fontWeight:500}}>{filterProduto}</span>
                    <span style={{background:"#7c3aed",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700,marginLeft:10}}>{activeOrders.filter(r=>r.produto===filterProduto).length} pedidos</span>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:5}}>
                    {activeOrders.filter(r=>r.produto===filterProduto).map((r,i)=>{
                      const dl=deadlineInfo(r.dataEnvio);
                      const si=r.statusInterno;
                      return (
                        <div key={r.idPedido} style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",background:si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#f8fafc":"#fff",borderRadius:10,padding:"8px 12px",border:"1px solid",borderColor:si==="feito"?"#86efac":si==="revisao"?"#fde68a":"#f1f5f9"}}>
                          <StyledCheckbox checked={selected.has(r.idPedido)} onChange={()=>toggleSelect(r.idPedido)} />
                          <span style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1d4ed8",flexShrink:0}}>{r.idPedido}</span>
                          <span style={{fontSize:11,color:"#64748b",flexShrink:0}}>Loja: <strong>{r.loja}</strong></span>
                          <span style={{fontSize:11,color:"#64748b",flexShrink:0}}>Qtd: <strong>{r.quantidade}</strong></span>
                          <span style={{fontSize:11,color:"#64748b",flex:1}}>{r.variacao}</span>
                          <span style={{fontSize:11,color:"#059669",fontWeight:700,flexShrink:0}}>{fmtBRL(r.preco)}</span>
                          {dl&&<span style={{background:dl.bg,color:dl.text,border:`1px solid ${dl.border}`,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:600,flexShrink:0}}>{dl.icon} {dl.label}</span>}
                          {si==="feito"&&<span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700,flexShrink:0}}>✅ Feito</span>}
                          {si==="revisao"&&<span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700,flexShrink:0}}>📋 Revisão</span>}
                          {!si&&<span style={{background:"#fde8e8",color:"#991b1b",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:600,flexShrink:0}}>A ENVIAR</span>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* ── Main Table ── */}
        <div ref={tableRef} style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden",marginBottom:18}}>
          {/* Filter bar */}
          <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <span style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>📋 Pedidos em Aberto</span>
            <input placeholder="🔍 Buscar ID, produto, destinatário..." value={search} onChange={e=>setSearch(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 12px",fontSize:12,outline:"none",flex:1,minWidth:180}} />
            <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 12px",fontSize:12,cursor:"pointer",background:"#fff"}}>
              <option value="all">Todos os estados</option>
              <option value="">⏳ Pendentes</option>
              <option value="feito">✅ Feito</option>
              <option value="revisao">📋 Revisão</option>
            </select>
            <select value={filterLoja} onChange={e=>setFilterLoja(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 12px",fontSize:12,cursor:"pointer",background:"#fff"}}>
              <option value="all">Todas as lojas</option>
              {LOJAS.map(l=><option key={l} value={l}>{l}</option>)}
            </select>
            <select value={filterPrazo} onChange={e=>setFilterPrazo(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 12px",fontSize:12,cursor:"pointer",background:"#fff"}}>
              <option value="all">Todos os prazos</option>
              <option value="red">🔴 Urgente (≤1 dia)</option>
              <option value="yellow">🟡 Atenção (2 dias)</option>
              <option value="green">🟢 OK (3+ dias)</option>
            </select>
          </div>
          {/* Selection bar */}
          <div style={{padding:"8px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"#64748b"}}><strong>{filtered.length}</strong> pedidos · <strong>{selected.size}</strong> selecionados</span>
            {selected.size>0&&(
              <>
                <button onClick={copySelected} style={{background:copied?"#059669":"#1d4ed8",color:"#fff",border:"none",borderRadius:10,padding:"5px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                  {copied?"✓ Copiado!`":`📋 Copiar ${selected.size} ID(s) → UpSeller`}
                </button>
                <button onClick={()=>setSelected(new Set())} style={{background:"none",border:"1px solid #e2e8f0",borderRadius:10,padding:"5px 10px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Limpar</button>
              </>
            )}
            <div style={{marginLeft:"auto",display:"flex",gap:6}}>
              {[{icon:"🟢",label:"3+d",bg:"#dcfce7",text:"#14532d",border:"#22c55e"},{icon:"🟡",label:"2d",bg:"#fef9c3",text:"#78350f",border:"#f59e0b"},{icon:"🔴",label:"≤1d",bg:"#fde8e8",text:"#991b1b",border:"#ef4444"}].map(l=>(
                <span key={l.label} style={{background:l.bg,color:l.text,border:`1px solid ${l.border}`,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:600}}>{l.icon} {l.label}</span>
              ))}
            </div>
          </div>
          {/* IDs preview */}
          {selected.size>0&&(
            <div style={{padding:"8px 18px",background:"#eff6ff",borderBottom:"1px solid #bfdbfe"}}>
              <div style={{fontSize:11,color:"#1d4ed8",fontWeight:600,marginBottom:3}}>IDs para colar no UpSeller:</div>
              <div style={{fontSize:12,color:"#1e40af",fontFamily:"monospace",wordBreak:"break-all",background:"#dbeafe",borderRadius:8,padding:"6px 10px"}}>{[...selected].join(",")}</div>
            </div>
          )}
          {/* Table */}
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  <th style={TH}><StyledCheckbox checked={selected.size>0&&selected.size===filtered.length} onChange={toggleAll} /></th>
                  <th style={TH}>Ações</th>
                  <th style={TH}>Status</th>
                  <th style={TH}>ID Pedido</th>
                  <th style={TH}>Destinatário</th>
                  <th style={TH}>Loja</th>
                  <th style={TH}>Produto</th>
                  <th style={TH}>Variação</th>
                  <th style={TH}>Qtd</th>
                  <th style={TH}>Preço</th>
                  <th style={TH}>Prazo Envio</th>
                  <th style={TH}>Pgto</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((row,i)=>{
                  const dl=deadlineInfo(row.dataEnvio);
                  const isSelected=selected.has(row.idPedido);
                  const si=row.statusInterno;
                  return (
                    <tr key={row.idPedido} style={{background:isSelected?"#eff6ff":si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f1f5f9"}}>
                      <td style={{...TD,textAlign:"center"}}><StyledCheckbox checked={isSelected} onChange={()=>toggleSelect(row.idPedido)} /></td>
                      {/* Action buttons - LEFT */}
                      <td style={{...TD,whiteSpace:"nowrap"}}>
                        <div style={{display:"flex",gap:4}}>
                          <button onClick={()=>setRevisaoModal(row)} style={{padding:"4px 8px",border:"none",borderRadius:7,background:si==="revisao"?"#fef3c7":"#f1f5f9",color:si==="revisao"?"#92400e":"#374151",fontSize:10,fontWeight:600,cursor:"pointer"}}>📋 Revisão</button>
                          <button onClick={()=>handleStatusChange(row,si==="feito"?"":"feito")} style={{padding:"4px 8px",border:"none",borderRadius:7,background:si==="feito"?"#d1fae5":"#f1f5f9",color:si==="feito"?"#065f46":"#374151",fontSize:10,fontWeight:600,cursor:"pointer"}}>{si==="feito"?"✅":"⬜"} Feito</button>
                        </div>
                      </td>
                      {/* Status pedido badge */}
                      <td style={{...TD,textAlign:"center"}}>
                        {si==="feito"?<span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>✅ Feito</span>
                        :si==="revisao"?<div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:1}}><span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>📋 Revisão</span>{row.notaRevisao&&<span style={{fontSize:9,color:"#92400e",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={row.notaRevisao}>{row.notaRevisao}</span>}</div>
                        :<span style={{background:"#fde8e8",color:"#991b1b",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>A ENVIAR</span>}
                      </td>
                      <td style={{...TD,fontWeight:700,color:"#1d4ed8",fontFamily:"monospace"}}>{row.idPedido}</td>
                      <td style={{...TD}}><ExpandCell value={row.destinatario||"—"} /></td>
                      <td style={TD}>{row.loja||"—"}</td>
                      <td style={{...TD,maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3}}>{row.produto||"—"}</div></td>
                      <td style={{...TD,maxWidth:140}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.variacao||"—"}</div></td>
                      <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,fontWeight:600,color:"#059669",whiteSpace:"nowrap"}}>{fmtBRL(row.preco)}</td>
                      <td style={TD}>
                        {dl?<div style={{background:dl.bg,border:`1px solid ${dl.border}`,borderRadius:8,padding:"3px 8px",display:"inline-flex",flexDirection:"column",gap:1,minWidth:80}}>
                          <span style={{fontSize:10,color:dl.text,fontWeight:700}}>{dl.icon} {dl.label}</span>
                          <span style={{fontSize:9,color:dl.text,opacity:0.75}}>{row.dataEnvio?.slice(0,10)}</span>
                        </div>:<span style={{color:"#94a3b8",fontSize:11}}>—</span>}
                      </td>
                      <td style={{...TD,fontSize:11,color:"#64748b"}}>{row.horaPagamento?.slice(0,10)||"—"}</td>
                    </tr>
                  );
                })}
                {filtered.length===0&&(
                  <tr><td colSpan={12} style={{textAlign:"center",padding:40,color:"#94a3b8",fontSize:14}}>
                    {activeOrders.length===0?"📭 Nenhum pedido ainda. Carregue uma planilha!":"Nenhum pedido encontrado com esses filtros."}
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* ── Enviados Section ── */}
        <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
          <button onClick={()=>setShowEnviados(v=>!v)}
            style={{width:"100%",padding:"14px 20px",background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",gap:10,textAlign:"left"}}>
            <span style={{fontSize:14,fontWeight:700,color:"#059669"}}>📬 Pedidos Enviados</span>
            <span style={{background:"#059669",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{enviados.length}</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>Histórico completo de envios</span>
            <span style={{marginLeft:"auto",color:"#94a3b8",fontSize:16}}>{showEnviados?"▲":"▼"}</span>
          </button>
          {showEnviados&&(
            <div style={{borderTop:"1px solid #f1f5f9",overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead>
                  <tr style={{background:"#f0fdf4"}}>
                    {["ID Pedido","Loja","Produto","Variação","Qtd","Preço","Data Envio","Destinatário"].map(h=>(
                      <th key={h} style={TH}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {enviados.map((row,i)=>(
                    <tr key={row.idPedido} style={{background:i%2===0?"#f0fdf4":"#fff",borderBottom:"1px solid #dcfce7"}}>
                      <td style={{...TD,fontWeight:700,color:"#059669",fontFamily:"monospace"}}>{row.idPedido}</td>
                      <td style={TD}>{row.loja||"—"}</td>
                      <td style={{...TD,maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3}}>{row.produto||"—"}</div></td>
                      <td style={{...TD,maxWidth:130}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.variacao||"—"}</div></td>
                      <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,fontWeight:600,color:"#059669"}}>{fmtBRL(row.preco)}</td>
                      <td style={TD}>{row.dataEnvio?.slice(0,10)||"—"}</td>
                      <td style={TD}>{row.destinatario||"—"}</td>
                    </tr>
                  ))}
                  {enviados.length===0&&(
                    <tr><td colSpan={8} style={{textAlign:"center",padding:30,color:"#94a3b8",fontSize:13}}>Nenhum pedido enviado ainda.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      {revisaoModal&&<RevisaoModal order={revisaoModal} onConfirm={nota=>{handleStatusChange(revisaoModal,"revisao",nota);setRevisaoModal(null);}} onClose={()=>setRevisaoModal(null)} />}
      {toast&&<Toast msg={toast.msg} color={toast.color} />}
    </div>
  );
}

const TH = {padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:10,color:"#64748b",letterSpacing:0.5,textTransform:"uppercase",whiteSpace:"nowrap",borderBottom:"2px solid #e2e8f0"};
const TD = {padding:"8px 12px",verticalAlign:"middle"};
