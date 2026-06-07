import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend } from "recharts";
import { supabase, fetchPedidos, upsertPedidos, updatePedidoStatus, fetchDevolucoes, upsertDevolucoes, fetchFaturamento, upsertFaturamento, deleteAllPedidos, deleteAllDevolucoes, deleteAllFaturamento } from "./supabase.js";

// ── Column maps — matched to REAL Shopee spreadsheet columns ─────────────────
const PEDIDO_COLS = {
  "ID do pedido":                "idPedido",
  "Status do pedido":            "statusPedido",
  "Data prevista de envio":      "dataEnvio",
  "Hora do pagamento do pedido": "horaPagamento",
  "Nome do Produto":             "produto",
  "Preço acordado":              "preco",
  "Quantidade":                  "quantidade",
  "Nome da variação":            "variacao",
  "Nome do destinatário":        "destinatario",
  "Observação do comprador":     "notas",
};

const DEVOLUCAO_COLS = {
  "ID do pedido":                    "id_pedido",
  "Data de criação do pedido":       "data_criacao",
  "Nome do Produto":                 "produto",
  "Nome da variação":                "variacao",
  "Preço da unidade":                "preco_unidade",
  "Status da Devolução / Reembolso": "status_devolucao",
  "Quantidade de Devoluções":        "quantidade",
  "Motivo da Devolução":             "motivo",
  "Observações da Devolução":        "observacoes",
};

// ── Normalize string for comparison (lowercase, no accents) ──────────────────
function norm(s) {
  return (s||"").toString().trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"");
}

// ── Parse spreadsheet with flexible column matching ───────────────────────────
function parseSheet(sheet, colMap, loja) {
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!raw.length) return [];
  const sheetKeys = Object.keys(raw[0]);
  const normKeys  = sheetKeys.map(k => norm(k));

  return raw.map(row => {
    const m = { loja: loja || "", _status: "new" };
    for (const [orig, key] of Object.entries(colMap)) {
      const idx = normKeys.findIndex(k => k === norm(orig));
      m[key] = idx >= 0 && row[sheetKeys[idx]] !== undefined ? String(row[sheetKeys[idx]]) : "";
    }
    if (colMap === PEDIDO_COLS) {
      m.preco = parseFloat(String(m.preco).replace(/[^\d.,]/g,"").replace(",",".")) || 0;
      m.statusInterno = ""; m.notaRevisao = "";
    } else {
      m.preco_unidade = parseFloat(String(m.preco_unidade).replace(/[^\d.,]/g,"").replace(",",".")) || 0;
    }
    return m;
  }).filter(r => r.idPedido || r.id_pedido);
}

// ── Status classification — covers all Shopee status variants ────────────────
function normStatus(s) { return norm(s||""); }

function isEnviado(r) {
  const s = normStatus(r.statusPedido);
  return s === "enviado" || s === "entregue" || s === "concluido" ||
         s.startsWith("o comprador pode pedir") || s === "completed" || s === "delivered";
}
function isCancelado(r) {
  const s = normStatus(r.statusPedido);
  return s === "cancelado" || s === "nao pago" || s === "unpaid" ||
         s === "cancelamento solicitado" || s.startsWith("cancelad");
}
function isAberto(r) {
  // Cancelled and sent must never appear in "abertos"
  if (isCancelado(r) || isEnviado(r)) return false;
  const s = normStatus(r.statusPedido);
  return s.includes("a enviar") || s === "a" || s === "pendente" || s === "order received";
}
function isHistorico(r) { return isEnviado(r) || isCancelado(r); }

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

function fmtBRL(v) { return Number(v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"}); }
function currentMes() { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; }

// ── Status badge ──────────────────────────────────────────────────────────────
function StatusBadge({status}) {
  const s = normStatus(status);
  let bg="#f1f5f9", color="#64748b", label=status||"—";
  if (isAberto({statusPedido:status}))    { bg="#fde8e8"; color="#991b1b"; label="🔴 A ENVIAR"; }
  else if (isEnviado({statusPedido:status}))  { bg="#d1fae5"; color="#065f46"; label="✅ ENVIADO"; }
  else if (isCancelado({statusPedido:status})){ bg="#f3f4f6"; color="#374151"; label="❌ CANCELADO"; }
  return <span style={{background:bg,color,borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{label}</span>;
}

// ── UI Components ─────────────────────────────────────────────────────────────
function StyledCheckbox({checked,onChange}) {
  return (
    <div onClick={onChange} style={{width:18,height:18,borderRadius:5,border:`2px solid ${checked?"#1d4ed8":"#cbd5e1"}`,background:checked?"#1d4ed8":"#fff",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",transition:"all 0.15s",flexShrink:0}}>
      {checked&&<svg width="10" height="8" viewBox="0 0 10 8" fill="none"><path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>}
    </div>
  );
}

function ExpandCell({value,maxLen=22}) {
  const [exp,setExp] = useState(false);
  const long = value && value.length > maxLen;
  return (
    <div style={{display:"flex",alignItems:"flex-start",gap:3,maxWidth:180}}>
      <span style={{fontSize:12,color:"#374151",lineHeight:1.4,whiteSpace:exp?"normal":"nowrap",overflow:exp?"visible":"hidden",textOverflow:exp?"clip":"ellipsis",flex:1}}>{value||"—"}</span>
      {long&&<button onClick={()=>setExp(v=>!v)} style={{background:"none",border:"none",cursor:"pointer",color:"#94a3b8",fontSize:12,padding:0,flexShrink:0}}>{exp?"▲":"▼"}</button>}
    </div>
  );
}

function DropZone({label,sublabel,onFile,file,color,disabled}) {
  const ref=useRef(); const [drag,setDrag]=useState(false);
  const handle=f=>{if(!f)return;const r=new FileReader();r.onload=e=>{const wb=XLSX.read(e.target.result,{type:"binary"});onFile(wb.Sheets[wb.SheetNames[0]],f.name);};r.readAsBinaryString(f);};
  return (
    <div onClick={()=>!disabled&&ref.current.click()} onDragOver={e=>{if(!disabled){e.preventDefault();setDrag(true);}}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);if(!disabled)handle(e.dataTransfer.files[0]);}}>
      <input ref={ref} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>handle(e.target.files[0])} />
      <div style={{border:`2px dashed ${drag?color:disabled?"#e2e8f0":"#cbd5e1"}`,borderRadius:12,padding:"16px 12px",cursor:disabled?"not-allowed":"pointer",background:disabled?"#f8fafc":drag?`${color}18`:"#f8fafc",textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",gap:5,opacity:disabled?0.5:1,minHeight:100}}>
        <div style={{fontSize:22}}>{disabled?"🔒":"📋"}</div>
        <div style={{fontWeight:700,fontSize:12,color:"#1e293b"}}>{label}</div>
        {sublabel&&<div style={{fontSize:10,color:"#94a3b8",maxWidth:160}}>{sublabel}</div>}
        {file?<div style={{background:color,color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:10,fontWeight:600}}>{file.length>24?file.slice(0,24)+"...":file}</div>:<div style={{fontSize:10,color:"#94a3b8"}}>Arraste ou clique</div>}
      </div>
    </div>
  );
}

function Toast({msg,color}) {
  if(!msg)return null;
  return <div style={{position:"fixed",bottom:24,right:24,zIndex:9999,background:color||"#1e293b",color:"#fff",borderRadius:12,padding:"12px 20px",fontSize:13,fontWeight:600,boxShadow:"0 8px 32px rgba(0,0,0,0.2)"}}>{msg}</div>;
}

function RevisaoModal({order,onConfirm,onClose}) {
  const [nota,setNota]=useState(order.notaRevisao||"");
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
        <h3 style={{margin:"0 0 6px",fontSize:15,fontWeight:700,color:"#1e293b"}}>📋 Revisão</h3>
        <p style={{margin:"0 0 14px",fontSize:12,color:"#64748b"}}>Pedido: <strong>{order.idPedido}</strong></p>
        <textarea value={nota} onChange={e=>setNota(e.target.value)} placeholder="Descreva o motivo..." autoFocus
          style={{width:"100%",minHeight:90,padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,resize:"vertical",outline:"none",boxSizing:"border-box",fontFamily:"inherit"}} />
        <div style={{display:"flex",gap:10,marginTop:14,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"8px 18px",border:"1px solid #e2e8f0",borderRadius:10,background:"#fff",fontSize:12,cursor:"pointer"}}>Cancelar</button>
          <button onClick={()=>onConfirm(nota)} disabled={!nota.trim()} style={{padding:"8px 18px",border:"none",borderRadius:10,background:nota.trim()?"#f59e0b":"#e2e8f0",color:nota.trim()?"#fff":"#94a3b8",fontSize:12,fontWeight:700,cursor:nota.trim()?"pointer":"not-allowed"}}>Confirmar</button>
        </div>
      </div>
    </div>
  );
}

const FINANCE_PASSWORD = "1234";
function FinanceGate({onUnlock}) {
  const [pwd,setPwd]=useState(""); const [err,setErr]=useState(false);
  const check=()=>{if(pwd===FINANCE_PASSWORD)onUnlock();else{setErr(true);setPwd("");}};
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:380,boxShadow:"0 20px 60px rgba(0,0,0,0.25)",textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:10}}>🔐</div>
        <h3 style={{margin:"0 0 6px",fontSize:16,fontWeight:700,color:"#1e293b"}}>Área Restrita</h3>
        <p style={{margin:"0 0 20px",fontSize:13,color:"#64748b"}}>Senha para acessar o financeiro</p>
        <input type="password" value={pwd} onChange={e=>{setPwd(e.target.value);setErr(false);}} onKeyDown={e=>e.key==="Enter"&&check()} placeholder="••••••••" autoFocus
          style={{width:"100%",padding:"11px 14px",border:`1.5px solid ${err?"#ef4444":"#e2e8f0"}`,borderRadius:10,fontSize:14,outline:"none",boxSizing:"border-box",textAlign:"center",letterSpacing:4,marginBottom:err?6:16}} />
        {err&&<p style={{margin:"0 0 14px",fontSize:12,color:"#ef4444",fontWeight:600}}>Senha incorreta.</p>}
        <button onClick={check} style={{width:"100%",padding:"12px",background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:700,cursor:"pointer"}}>Entrar →</button>
      </div>
    </div>
  );
}

function TabBar({tabs,active,onChange}) {
  return (
    <div style={{display:"flex",gap:4,background:"#f1f5f9",borderRadius:12,padding:4,marginBottom:20,flexWrap:"wrap"}}>
      {tabs.map(t=>(
        <button key={t.id} onClick={()=>onChange(t.id)} style={{flex:1,minWidth:90,padding:"9px 14px",border:"none",borderRadius:9,fontSize:12,fontWeight:600,cursor:"pointer",transition:"all 0.15s",background:active===t.id?"#fff":"transparent",color:active===t.id?t.color||"#1d4ed8":"#64748b",boxShadow:active===t.id?"0 1px 4px rgba(0,0,0,0.1)":"none",display:"flex",alignItems:"center",justifyContent:"center",gap:6,whiteSpace:"nowrap"}}>
          {t.icon} {t.label} {t.badge!=null&&<span style={{background:active===t.id?(t.color||"#1d4ed8"):"#94a3b8",color:"#fff",borderRadius:20,padding:"0px 7px",fontSize:10,fontWeight:700}}>{t.badge}</span>}
        </button>
      ))}
    </div>
  );
}

const LOJAS = ["Gran Shop","Aishael Mix"];
const TH={padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:10,color:"#64748b",letterSpacing:0.5,textTransform:"uppercase",whiteSpace:"nowrap",borderBottom:"2px solid #e2e8f0"};
const TD={padding:"7px 11px",verticalAlign:"middle"};

// ── MAIN APP ──────────────────────────────────────────────────────────────────
export default function App({user,onLogout}) {
  const [allPedidos,    setAllPedidos]    = useState([]);
  const [devolucoes,    setDevolucoes]    = useState([]);
  const [faturamento,   setFaturamento]   = useState([]);
  const [dbLoading,     setDbLoading]     = useState(true);
  const [saving,        setSaving]        = useState(false);
  const [toast,         setToast]         = useState(null);
  const [activeTab,     setActiveTab]     = useState("abertos");
  const [financeUnlocked,setFinanceUnlocked]=useState(false);
  const [showFinanceGate,setShowFinanceGate]=useState(false);
  const [revisaoModal,  setRevisaoModal]  = useState(null);
  const [uploadNames,   setUploadNames]   = useState({});
  const [search,        setSearch]        = useState("");
  const [filterLoja,    setFilterLoja]    = useState("all");
  const [filterPrazo,   setFilterPrazo]   = useState("all");
  const [filterSt,      setFilterSt]      = useState("all");
  const [filterProduto, setFilterProduto] = useState("all");
  const [showProdPanel, setShowProdPanel] = useState(false);
  const [selected,      setSelected]      = useState(new Set());
  const [copied,        setCopied]        = useState(false);
  const tableRef = useRef(null);

  const showToast=(msg,color,ms=3500)=>{setToast({msg,color});setTimeout(()=>setToast(null),ms);};

  // ── Load all data from Supabase on mount ───────────────────────────────────
  useEffect(()=>{
    if(!user?.id){setDbLoading(false);return;}
    if(!supabase){setDbLoading(false);return;}
    setDbLoading(true);
    Promise.all([fetchPedidos(user.id),fetchDevolucoes(user.id),fetchFaturamento(user.id)])
      .then(([p,d,f])=>{
        setAllPedidos(p); setDevolucoes(d); setFaturamento(f);
        setDbLoading(false);
      }).catch(err=>{console.error(err);setDbLoading(false);});
  },[user?.id]);

  // ── Pedido groups ──────────────────────────────────────────────────────────
  const pedidosAbertos    = allPedidos.filter(isAberto);
  const pedidosEnviados   = allPedidos.filter(isEnviado);
  const pedidosCancelados = allPedidos.filter(isCancelado);

  // ── Load planilha ──────────────────────────────────────────────────────────
  const handlePlanilha = async (sheet, name, loja) => {
    setUploadNames(p=>({...p,[`ped_${loja}`]:name}));
    const incoming = parseSheet(sheet, PEDIDO_COLS, loja);
    if (!incoming.length) { showToast("Nenhum pedido encontrado na planilha.","#d97706"); return; }

    // Merge: update existing, add new
    const map = new Map(allPedidos.map(r=>[r.idPedido,r]));
    let added=0, updated=0;
    for (const r of incoming) {
      const ex = map.get(r.idPedido);
      if (ex) {
        map.set(r.idPedido,{...ex,statusPedido:r.statusPedido,dataEnvio:r.dataEnvio,horaPagamento:r.horaPagamento,produto:r.produto,preco:r.preco,quantidade:r.quantidade,variacao:r.variacao,destinatario:r.destinatario,notas:r.notas,_status:"existing"});
        updated++;
      } else {
        map.set(r.idPedido,{...r,statusInterno:"",notaRevisao:""});
        added++;
      }
    }
    const merged = Array.from(map.values());
    setAllPedidos(merged);
    showToast(`✅ ${added} novos · ${updated} atualizados`,"#059669");

    // Save to DB
    if (supabase) {
      setSaving(true);
      await upsertPedidos(incoming, user.id);
      // Faturamento: only count "a enviar" orders
      const aEnviar = incoming.filter(isAberto);
      const totalValor = aEnviar.reduce((s,r)=>s+(parseFloat(r.preco)||0),0);
      if (totalValor > 0) await upsertFaturamento(user.id, currentMes(), loja, totalValor);
      const fat = await fetchFaturamento(user.id);
      setFaturamento(fat);
      setSaving(false);
    }
  };

  // ── Load devolução ─────────────────────────────────────────────────────────
  const handleDevolucao = async (sheet, name, loja) => {
    setUploadNames(p=>({...p,[`dev_${loja}`]:name}));
    const rows = parseSheet(sheet, DEVOLUCAO_COLS, loja);
    if (!rows.length) { showToast("Nenhuma devolução encontrada.","#d97706"); return; }
    const existIds = new Set(devolucoes.map(d=>d.id_pedido));
    const novas = rows.filter(r=>!existIds.has(r.id_pedido));
    setDevolucoes(prev=>[...prev,...novas]);
    showToast(`🔄 ${novas.length} devolução(ões) carregada(s)`,"#ef4444");
    if (supabase) { setSaving(true); await upsertDevolucoes(rows,user.id); setSaving(false); }
  };

  // ── Clear all data ────────────────────────────────────────────────────────
  const [showConfirmClear, setShowConfirmClear] = useState(false);

  const handleClearAll = async () => {
    setShowConfirmClear(false);
    setSaving(true);
    if (supabase) {
      await Promise.all([
        deleteAllPedidos(user.id),
        deleteAllDevolucoes(user.id),
        deleteAllFaturamento(user.id),
      ]);
    }
    setAllPedidos([]);
    setDevolucoes([]);
    setFaturamento([]);
    setUploadNames({});
    setSelected(new Set());
    setActiveTab("abertos");
    setSaving(false);
    showToast("🗑️ Sistema zerado com sucesso!", "#059669");
  };

  // ── Status interno ─────────────────────────────────────────────────────────
  const handleStatusChange = async (order,newSt,nota="") => {
    setAllPedidos(prev=>prev.map(o=>o.idPedido===order.idPedido?{...o,statusInterno:newSt,notaRevisao:nota}:o));
    if (supabase) await updatePedidoStatus(order.idPedido,user.id,newSt,nota);
    showToast(newSt==="feito"?"✅ Marcado como Feito!":"📋 Enviado para Revisão",newSt==="feito"?"#059669":"#f59e0b");
  };

  // ── Selection ──────────────────────────────────────────────────────────────
  const toggleSelect=id=>setSelected(p=>{const s=new Set(p);s.has(id)?s.delete(id):s.add(id);return s;});
  const toggleAll=()=>selected.size===filtered.length?setSelected(new Set()):setSelected(new Set(filtered.map(r=>r.idPedido)));
  const copyIds=()=>{navigator.clipboard.writeText([...selected].join(",")).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);showToast("📋 IDs copiados!","#1d4ed8");});};

  // ── Filters ────────────────────────────────────────────────────────────────
  const produtoMap = pedidosAbertos.reduce((a,r)=>{if(r.produto)a[r.produto]=(a[r.produto]||0)+1;return a;},{});
  const produtos   = Object.entries(produtoMap).sort((a,b)=>b[1]-a[1]);

  const filtered = pedidosAbertos.filter(r=>{
    const dl=deadlineInfo(r.dataEnvio);
    const q=search.toLowerCase();
    return (
      (filterLoja==="all"   ||r.loja===filterLoja)&&
      (filterPrazo==="all"  ||(dl&&dl.tier===filterPrazo))&&
      (filterSt==="all"     ||r.statusInterno===filterSt)&&
      (filterProduto==="all"||r.produto===filterProduto)&&
      (!search||r.idPedido.toLowerCase().includes(q)||r.produto.toLowerCase().includes(q)||r.destinatario.toLowerCase().includes(q)||r.variacao.toLowerCase().includes(q))
    );
  }).sort((a,b)=>{
    const day=r=>{if(!r.dataEnvio)return 9999;const d=new Date(r.dataEnvio.replace(" ","T"));if(isNaN(d))return 9999;const t=new Date();t.setHours(0,0,0,0);const dd=new Date(d);dd.setHours(0,0,0,0);return Math.round((dd-t)/86400000);};
    return day(a)-day(b);
  });

  const urgentCount  = pedidosAbertos.filter(r=>deadlineInfo(r.dataEnvio)?.tier==="red").length;
  const feitoCount   = pedidosAbertos.filter(r=>r.statusInterno==="feito").length;
  const revisaoCount = pedidosAbertos.filter(r=>r.statusInterno==="revisao").length;
  const valorAberto  = pedidosAbertos.reduce((s,r)=>s+(parseFloat(r.preco)||0),0);
  const totalDev     = devolucoes.reduce((s,d)=>s+(parseFloat(d.preco_unidade||0)*parseInt(d.quantidade||1)),0);

  // Chart
  const chartData = faturamento.reduce((acc,f)=>{
    const ex=acc.find(a=>a.mes===f.mes);
    if(ex) ex[f.loja]=(ex[f.loja]||0)+Number(f.valor);
    else { const n={mes:f.mes}; n[f.loja]=Number(f.valor); acc.push(n); }
    return acc;
  },[]).sort((a,b)=>a.mes.localeCompare(b.mes));

  const TABS = [
    {id:"abertos",    icon:"📋",label:"Em Aberto",  badge:pedidosAbertos.length,    color:"#1d4ed8"},
    {id:"enviados",   icon:"📦",label:"Enviados",   badge:pedidosEnviados.length,   color:"#059669"},
    {id:"cancelados", icon:"❌",label:"Cancelados", badge:pedidosCancelados.length, color:"#6b7280"},
    {id:"devolucoes", icon:"🔄",label:"Devoluções", badge:devolucoes.length,        color:"#ef4444"},
    {id:"financeiro", icon:"💰",label:"Financeiro", badge:null,                     color:"#0891b2"},
  ];

  if (dbLoading) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#1e3a8a,#3b82f6)",fontFamily:"sans-serif"}}>
      <div style={{color:"#fff",fontSize:16,fontWeight:600}}>📦 Carregando seus pedidos...</div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#f0f4ff",fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
      <style>{`*{box-sizing:border-box}`}</style>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#1e3a8a 0%,#1d4ed8 60%,#3b82f6 100%)",padding:"18px 28px",color:"#fff"}}>
        <div style={{maxWidth:1600,margin:"0 auto",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:10,fontWeight:600,letterSpacing:2,opacity:0.7}}>GERENCIADOR DE PEDIDOS</div>
            <h1 style={{margin:"2px 0 0",fontSize:20,fontWeight:800}}>Painel de Pedidos 📦</h1>
          </div>
          <span style={{background:"rgba(255,255,255,0.15)",borderRadius:20,padding:"3px 12px",fontSize:12,display:"flex",alignItems:"center",gap:5}}>
            <span style={{width:6,height:6,borderRadius:"50%",background:saving?"#fbbf24":"#4ade80",display:"inline-block"}}/>{saving?"Salvando...":"Conectado"}
          </span>
          <span style={{background:"rgba(255,255,255,0.12)",borderRadius:20,padding:"3px 12px",fontSize:12}}>👤 {user?.email}</span>
          <button onClick={onLogout} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"3px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Sair →</button>
        </div>
      </div>

      <div style={{maxWidth:1600,margin:"0 auto",padding:"20px 28px"}}>

        {/* Upload */}
        <div style={{background:"#fff",borderRadius:18,padding:18,marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
          <div style={{display:"flex",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:0.5}}>Carregar Planilhas</div>
            <button onClick={()=>setShowConfirmClear(true)} style={{marginLeft:"auto",background:"#fde8e8",color:"#991b1b",border:"1px solid #fca5a5",borderRadius:10,padding:"6px 14px",fontSize:11,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:5}}>
              🗑️ Zerar Sistema
            </button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
            {LOJAS.map(loja=>(
              <div key={`ped_${loja}`} style={{border:"1px solid #e2e8f0",borderRadius:12,padding:10}}>
                <div style={{fontSize:11,fontWeight:700,color:"#1d4ed8",marginBottom:7,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{background:"#eff6ff",borderRadius:6,padding:"1px 8px"}}>{loja}</span>
                  <span style={{color:"#64748b",fontWeight:400}}>· Pedidos</span>
                </div>
                <DropZone label="Carregar Planilha" sublabel="A Enviar / Enviado / Entregue / Cancelado" color="#1d4ed8"
                  file={uploadNames[`ped_${loja}`]} onFile={(s,n)=>handlePlanilha(s,n,loja)} />
              </div>
            ))}
            {LOJAS.map(loja=>(
              <div key={`dev_${loja}`} style={{border:"1px solid #fee2e2",borderRadius:12,padding:10}}>
                <div style={{fontSize:11,fontWeight:700,color:"#ef4444",marginBottom:7,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{background:"#fef2f2",borderRadius:6,padding:"1px 8px"}}>{loja}</span>
                  <span style={{color:"#64748b",fontWeight:400}}>· Devoluções</span>
                </div>
                <DropZone label="Planilha de Devolução" sublabel="Devolução e Reembolso" color="#ef4444"
                  file={uploadNames[`dev_${loja}`]} onFile={(s,n)=>handleDevolucao(s,n,loja)} />
              </div>
            ))}
          </div>
        </div>

        {/* Stats */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:10,marginBottom:18}}>
          {[
            {label:"Em Aberto",  value:pedidosAbertos.length,    color:"#1d4ed8",icon:"📋",tab:"abertos",  action:()=>{setFilterSt("all");setFilterPrazo("all");}},
            {label:"Feitos",     value:feitoCount,                color:"#059669",icon:"✅",tab:"abertos",  action:()=>setFilterSt("feito")},
            {label:"Em Revisão", value:revisaoCount,              color:"#f59e0b",icon:"📋",tab:"abertos",  action:()=>setFilterSt("revisao")},
            {label:"Urgentes",   value:urgentCount,               color:"#ef4444",icon:"🔴",tab:"abertos",  action:()=>{setFilterSt("all");setFilterPrazo("red");}},
            {label:"Enviados",   value:pedidosEnviados.length,    color:"#059669",icon:"📦",tab:"enviados", action:()=>{}},
            {label:"Cancelados", value:pedidosCancelados.length,  color:"#6b7280",icon:"❌",tab:"cancelados",action:()=>{}},
            {label:"Devoluções", value:devolucoes.length,         color:"#ef4444",icon:"🔄",tab:"devolucoes",action:()=>{}},
          ].map(s=>(
            <div key={s.label} onClick={()=>{s.action();setActiveTab(s.tab);tableRef.current?.scrollIntoView({behavior:"smooth"});}}
              style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",borderLeft:`4px solid ${s.color}`,cursor:"pointer",transition:"transform 0.15s,box-shadow 0.15s"}}
              onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 6px 20px rgba(0,0,0,0.12)";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,0.07)";}}>
              <div style={{fontSize:16,marginBottom:2}}>{s.icon}</div>
              <div style={{fontSize:20,fontWeight:800,color:s.color,lineHeight:1.2}}>{s.value}</div>
              <div style={{fontSize:10,color:"#64748b",fontWeight:500,marginTop:1}}>{s.label}</div>
              {s.label==="Devoluções"&&totalDev>0&&<div style={{fontSize:10,color:"#ef4444",fontWeight:700}}>-{fmtBRL(totalDev)}</div>}
              <div style={{fontSize:9,color:s.color,opacity:0.7,marginTop:1}}>clique →</div>
            </div>
          ))}
        </div>

        {/* Valor aberto card */}
        <div style={{background:"#eff6ff",borderRadius:14,padding:"14px 18px",marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:22}}>💰</span>
          <div>
            <div style={{fontSize:11,color:"#64748b",fontWeight:500}}>Valor em Aberto</div>
            <div style={{fontSize:22,fontWeight:800,color:"#0891b2"}}>{fmtBRL(valorAberto)}</div>
          </div>
        </div>

        {/* Tabs */}
        <TabBar tabs={TABS} active={activeTab} onChange={tab=>{
          if(tab==="financeiro"&&!financeUnlocked){setShowFinanceGate(true);}
          else setActiveTab(tab);
        }} />

        {/* ── TAB: ABERTOS ── */}
        {activeTab==="abertos"&&(
          <>
            {/* Product panel */}
            <div style={{background:"#fff",borderRadius:18,marginBottom:16,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
              <button onClick={()=>setShowProdPanel(v=>!v)} style={{width:"100%",padding:"13px 18px",background:"none",border:"none",cursor:"pointer",display:"flex",alignItems:"center",gap:10,textAlign:"left"}}>
                <span style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>📦 Filtrar por Produto</span>
                <span style={{fontSize:11,color:"#94a3b8"}}>{produtos.length} produto{produtos.length!==1?"s":""}</span>
                {filterProduto!=="all"&&<span style={{background:"#7c3aed",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>Ativo</span>}
                <span style={{marginLeft:"auto",color:"#94a3b8",fontSize:14}}>{showProdPanel?"▲":"▼"}</span>
              </button>
              {showProdPanel&&(
                <div style={{borderTop:"1px solid #f1f5f9",padding:"12px 18px"}}>
                  <div style={{display:"flex",flexWrap:"wrap",gap:7,marginBottom:filterProduto!=="all"?14:0}}>
                    {filterProduto!=="all"&&<button onClick={()=>setFilterProduto("all")} style={{background:"#f1f5f9",color:"#374151",border:"1px solid #e2e8f0",borderRadius:10,padding:"5px 11px",fontSize:11,cursor:"pointer",fontWeight:600}}>✕ Limpar</button>}
                    {produtos.map(([nome,qtd])=>{
                      const isA=filterProduto===nome;
                      return (<button key={nome} onClick={()=>{setFilterProduto(isA?"all":nome);tableRef.current?.scrollIntoView({behavior:"smooth"});}} title={nome}
                        style={{display:"flex",alignItems:"center",gap:5,background:isA?"#7c3aed":"#f8fafc",color:isA?"#fff":"#374151",border:`1.5px solid ${isA?"#7c3aed":"#e2e8f0"}`,borderRadius:10,padding:"5px 11px",fontSize:11,fontWeight:isA?700:500,cursor:"pointer"}}>
                        <span style={{background:isA?"rgba(255,255,255,0.25)":"#7c3aed",color:"#fff",borderRadius:20,padding:"0px 6px",fontSize:10,fontWeight:700,flexShrink:0}}>{qtd}</span>
                        {nome.length>50?nome.slice(0,50)+"...":nome}
                      </button>);
                    })}
                  </div>
                  {filterProduto!=="all"&&(
                    <div>
                      <div style={{fontSize:12,fontWeight:700,color:"#5b21b6",marginBottom:8}}>
                        Pedidos de: <span style={{fontWeight:500}}>{filterProduto}</span>
                        <span style={{background:"#7c3aed",color:"#fff",borderRadius:20,padding:"1px 9px",fontSize:11,fontWeight:700,marginLeft:6}}>{pedidosAbertos.filter(r=>r.produto===filterProduto).length}</span>
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:4}}>
                        {pedidosAbertos.filter(r=>r.produto===filterProduto).map((r,i)=>{
                          const dl=deadlineInfo(r.dataEnvio);
                          return (
                            <div key={r.idPedido} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",background:i%2===0?"#f8fafc":"#fff",borderRadius:9,padding:"7px 11px",border:"1px solid #f1f5f9"}}>
                              <StyledCheckbox checked={selected.has(r.idPedido)} onChange={()=>toggleSelect(r.idPedido)} />
                              <span style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1d4ed8"}}>{r.idPedido}</span>
                              <span style={{fontSize:11,color:"#64748b"}}>Loja: <strong>{r.loja}</strong></span>
                              <span style={{fontSize:11,color:"#064e3b",fontWeight:700}}>{fmtBRL(r.preco)}</span>
                              {dl&&<span style={{background:dl.bg,color:dl.text,border:`1px solid ${dl.border}`,borderRadius:7,padding:"1px 7px",fontSize:10,fontWeight:600}}>{dl.icon} {dl.label}</span>}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Main table */}
            <div ref={tableRef} style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
              <div style={{padding:"12px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                <input placeholder="🔍 Buscar ID, produto, destinatário..." value={search} onChange={e=>setSearch(e.target.value)}
                  style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 12px",fontSize:12,outline:"none",flex:1,minWidth:180}} />
                <select value={filterSt} onChange={e=>setFilterSt(e.target.value)} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 11px",fontSize:12,cursor:"pointer",background:"#fff"}}>
                  <option value="all">Todos os estados</option>
                  <option value="">⏳ Pendentes</option>
                  <option value="feito">✅ Feito</option>
                  <option value="revisao">📋 Revisão</option>
                </select>
                <select value={filterLoja} onChange={e=>setFilterLoja(e.target.value)} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 11px",fontSize:12,cursor:"pointer",background:"#fff"}}>
                  <option value="all">Todas as lojas</option>
                  {LOJAS.map(l=><option key={l} value={l}>{l}</option>)}
                </select>
                <select value={filterPrazo} onChange={e=>setFilterPrazo(e.target.value)} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 11px",fontSize:12,cursor:"pointer",background:"#fff"}}>
                  <option value="all">Todos os prazos</option>
                  <option value="red">🔴 Urgente</option>
                  <option value="yellow">🟡 Atenção (2d)</option>
                  <option value="green">🟢 OK (3+d)</option>
                </select>
              </div>
              <div style={{padding:"8px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                <span style={{fontSize:12,color:"#64748b"}}><strong>{filtered.length}</strong> pedidos · <strong>{selected.size}</strong> selecionados</span>
                {selected.size>0&&<>
                  <button onClick={copyIds} style={{background:copied?"#059669":"#1d4ed8",color:"#fff",border:"none",borderRadius:10,padding:"5px 13px",fontSize:12,fontWeight:700,cursor:"pointer"}}>{copied?"✓ Copiado!`":`📋 Copiar ${selected.size} ID(s)`}</button>
                  <button onClick={()=>setSelected(new Set())} style={{background:"none",border:"1px solid #e2e8f0",borderRadius:10,padding:"5px 10px",fontSize:12,cursor:"pointer",color:"#64748b"}}>Limpar</button>
                </>}
                <div style={{marginLeft:"auto",display:"flex",gap:5}}>
                  {[{icon:"🟢",label:"3+d",bg:"#dcfce7",text:"#14532d",bdr:"#22c55e"},{icon:"🟡",label:"2d",bg:"#fef9c3",text:"#78350f",bdr:"#f59e0b"},{icon:"🔴",label:"≤1d",bg:"#fde8e8",text:"#991b1b",bdr:"#ef4444"}].map(l=>(
                    <span key={l.label} style={{background:l.bg,color:l.text,border:`1px solid ${l.bdr}`,borderRadius:7,padding:"2px 7px",fontSize:10,fontWeight:600}}>{l.icon} {l.label}</span>
                  ))}
                </div>
              </div>
              {selected.size>0&&<div style={{padding:"7px 18px",background:"#eff6ff",borderBottom:"1px solid #bfdbfe"}}>
                <div style={{fontSize:10,color:"#1d4ed8",fontWeight:600,marginBottom:2}}>IDs para UpSeller:</div>
                <div style={{fontSize:11,color:"#1e40af",fontFamily:"monospace",wordBreak:"break-all",background:"#dbeafe",borderRadius:7,padding:"5px 9px"}}>{[...selected].join(",")}</div>
              </div>}
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
                      <th style={TH}>Prazo</th>
                      <th style={TH}>Notas</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((row,i)=>{
                      const dl=deadlineInfo(row.dataEnvio);
                      const isSel=selected.has(row.idPedido);
                      const si=row.statusInterno;
                      return (
                        <tr key={row.idPedido} style={{background:isSel?"#eff6ff":si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f1f5f9"}}>
                          <td style={{...TD,textAlign:"center"}}><StyledCheckbox checked={isSel} onChange={()=>toggleSelect(row.idPedido)} /></td>
                          <td style={{...TD,whiteSpace:"nowrap"}}>
                            <div style={{display:"flex",gap:4}}>
                              <button onClick={()=>setRevisaoModal(row)} style={{padding:"4px 7px",border:"none",borderRadius:7,background:si==="revisao"?"#fef3c7":"#f1f5f9",color:si==="revisao"?"#92400e":"#374151",fontSize:10,fontWeight:600,cursor:"pointer"}}>📋</button>
                              <button onClick={()=>handleStatusChange(row,si==="feito"?"":"feito")} style={{padding:"4px 7px",border:"none",borderRadius:7,background:si==="feito"?"#d1fae5":"#f1f5f9",color:si==="feito"?"#065f46":"#374151",fontSize:10,fontWeight:600,cursor:"pointer"}}>{si==="feito"?"✅":"⬜"}</button>
                            </div>
                          </td>
                          <td style={{...TD,textAlign:"center"}}>
                            {si==="feito"?<span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:700}}>✅ Feito</span>
                            :si==="revisao"?<span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:700}}>📋 Rev.</span>
                            :<StatusBadge status={row.statusPedido} />}
                          </td>
                          <td style={{...TD,fontWeight:700,color:"#1d4ed8",fontFamily:"monospace",fontSize:11}}>{row.idPedido}</td>
                          <td style={TD}><ExpandCell value={row.destinatario||"—"} /></td>
                          <td style={TD}>{row.loja||"—"}</td>
                          <td style={{...TD,maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:200}}>{row.produto||"—"}</div></td>
                          <td style={TD}><ExpandCell value={row.variacao||"—"} maxLen={20} /></td>
                          <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                          <td style={{...TD,fontWeight:600,color:"#059669",whiteSpace:"nowrap"}}>{fmtBRL(row.preco)}</td>
                          <td style={TD}>{dl?<div style={{background:dl.bg,border:`1px solid ${dl.border}`,borderRadius:7,padding:"3px 7px",display:"inline-flex",flexDirection:"column",gap:1,minWidth:74}}><span style={{fontSize:10,color:dl.text,fontWeight:700}}>{dl.icon} {dl.label}</span><span style={{fontSize:9,color:dl.text,opacity:0.75}}>{row.dataEnvio?.slice(0,10)}</span></div>:<span style={{color:"#94a3b8",fontSize:10}}>—</span>}</td>
                          <td style={TD}><ExpandCell value={row.notas||"—"} maxLen={20} /></td>
                        </tr>
                      );
                    })}
                    {filtered.length===0&&<tr><td colSpan={12} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>{pedidosAbertos.length===0?"📭 Nenhum pedido. Carregue uma planilha!":"Nenhum pedido com esses filtros."}</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {/* ── TAB: ENVIADOS ── */}
        {activeTab==="enviados"&&(
          <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#059669"}}>📦 Pedidos Enviados / Entregues</span>
              <span style={{background:"#059669",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{pedidosEnviados.length}</span>
              <span style={{marginLeft:"auto",fontSize:12,fontWeight:700,color:"#059669"}}>{fmtBRL(pedidosEnviados.reduce((s,r)=>s+(parseFloat(r.preco)||0),0))}</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#f0fdf4"}}>{["Status","ID Pedido","Destinatário","Loja","Produto","Variação","Qtd","Preço","Data Envio"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                <tbody>
                  {pedidosEnviados.map((row,i)=>(
                    <tr key={row.idPedido} style={{background:i%2===0?"#f8fffe":"#fff",borderBottom:"1px solid #f0fdf4"}}>
                      <td style={{...TD,textAlign:"center"}}><StatusBadge status={row.statusPedido} /></td>
                      <td style={{...TD,fontWeight:700,color:"#059669",fontFamily:"monospace",fontSize:11}}>{row.idPedido}</td>
                      <td style={TD}><ExpandCell value={row.destinatario||"—"} /></td>
                      <td style={TD}>{row.loja||"—"}</td>
                      <td style={{...TD,maxWidth:190}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:190}}>{row.produto||"—"}</div></td>
                      <td style={TD}><ExpandCell value={row.variacao||"—"} maxLen={20} /></td>
                      <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,fontWeight:600,color:"#059669"}}>{fmtBRL(row.preco)}</td>
                      <td style={TD}>{row.dataEnvio?.slice(0,10)||"—"}</td>
                    </tr>
                  ))}
                  {pedidosEnviados.length===0&&<tr><td colSpan={9} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>Nenhum pedido enviado ainda.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── TAB: CANCELADOS ── */}
        {activeTab==="cancelados"&&(
          <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#6b7280"}}>❌ Pedidos Cancelados</span>
              <span style={{background:"#6b7280",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{pedidosCancelados.length}</span>
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#f9fafb"}}>{["Status","ID Pedido","Destinatário","Loja","Produto","Variação","Qtd","Preço","Data"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                <tbody>
                  {pedidosCancelados.map((row,i)=>(
                    <tr key={row.idPedido} style={{background:i%2===0?"#f9fafb":"#fff",borderBottom:"1px solid #f3f4f6",opacity:0.85}}>
                      <td style={{...TD,textAlign:"center"}}><StatusBadge status={row.statusPedido} /></td>
                      <td style={{...TD,fontWeight:700,color:"#6b7280",fontFamily:"monospace",fontSize:11}}>{row.idPedido}</td>
                      <td style={TD}><ExpandCell value={row.destinatario||"—"} /></td>
                      <td style={TD}>{row.loja||"—"}</td>
                      <td style={{...TD,maxWidth:190}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:190,textDecoration:"line-through",color:"#9ca3af"}}>{row.produto||"—"}</div></td>
                      <td style={TD}><ExpandCell value={row.variacao||"—"} maxLen={20} /></td>
                      <td style={{...TD,textAlign:"center",color:"#9ca3af"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,fontWeight:600,color:"#9ca3af",textDecoration:"line-through"}}>{fmtBRL(row.preco)}</td>
                      <td style={{...TD,color:"#9ca3af"}}>{row.dataEnvio?.slice(0,10)||"—"}</td>
                    </tr>
                  ))}
                  {pedidosCancelados.length===0&&<tr><td colSpan={9} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>Nenhum pedido cancelado.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── TAB: DEVOLUCOES ── */}
        {activeTab==="devolucoes"&&(
          <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #fee2e2",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#ef4444"}}>🔄 Devoluções e Reembolsos</span>
              <span style={{background:"#ef4444",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{devolucoes.length}</span>
              {totalDev>0&&<span style={{background:"#fde8e8",color:"#991b1b",borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:800}}>-{fmtBRL(totalDev)}</span>}
            </div>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#fef2f2"}}>{["ID Pedido","Data","Produto","Variação","Qtd","Preço Unit.","Status Dev.","Motivo","Observações","Loja"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                <tbody>
                  {devolucoes.map((row,i)=>(
                    <tr key={row.id_pedido||i} style={{background:i%2===0?"#fff5f5":"#fff",borderBottom:"1px solid #fee2e2"}}>
                      <td style={{...TD,fontWeight:700,color:"#ef4444",fontFamily:"monospace",fontSize:11}}>{row.id_pedido}</td>
                      <td style={TD}>{row.data_criacao?.slice(0,10)||"—"}</td>
                      <td style={{...TD,maxWidth:180}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:180}}>{row.produto||"—"}</div></td>
                      <td style={TD}><ExpandCell value={row.variacao||"—"} maxLen={20} /></td>
                      <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,color:"#ef4444",fontWeight:600}}>-{fmtBRL(row.preco_unidade)}</td>
                      <td style={TD}><span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:600}}>{row.status_devolucao||"—"}</span></td>
                      <td style={{...TD,maxWidth:160}}><ExpandCell value={row.motivo||"—"} maxLen={25} /></td>
                      <td style={{...TD,maxWidth:160}}><ExpandCell value={row.observacoes||"—"} maxLen={25} /></td>
                      <td style={TD}>{row.loja||"—"}</td>
                    </tr>
                  ))}
                  {devolucoes.length===0&&<tr><td colSpan={10} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>Nenhuma devolução registrada.</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── TAB: FINANCEIRO ── */}
        {activeTab==="financeiro"&&financeUnlocked&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:18}}>
              {LOJAS.map(loja=>{
                const totalFat=faturamento.filter(f=>f.loja===loja).reduce((s,f)=>s+Number(f.valor),0);
                const totalPed=allPedidos.filter(r=>r.loja===loja).length||1;
                const ticket=totalFat/totalPed;
                return (
                  <div key={loja} style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:8,textTransform:"uppercase",letterSpacing:0.5}}>{loja}</div>
                    <div style={{display:"flex",flexDirection:"column",gap:4}}>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontSize:11,color:"#94a3b8"}}>Total acum.</span><span style={{fontSize:14,fontWeight:800,color:"#1d4ed8"}}>{fmtBRL(totalFat)}</span></div>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontSize:11,color:"#94a3b8"}}>Pedidos</span><span style={{fontSize:13,fontWeight:700}}>{allPedidos.filter(r=>r.loja===loja).length}</span></div>
                      <div style={{height:1,background:"#e2e8f0",margin:"2px 0"}}/>
                      <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontSize:11,color:"#94a3b8"}}>🎯 Ticket Médio</span><span style={{fontSize:14,fontWeight:800,color:"#059669"}}>{fmtBRL(ticket)}</span></div>
                    </div>
                  </div>
                );
              })}
              <div style={{background:"#eff6ff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:26,marginBottom:4}}>💰</div>
                <div style={{fontSize:22,fontWeight:800,color:"#0891b2"}}>{fmtBRL(faturamento.reduce((s,f)=>s+Number(f.valor),0))}</div>
                <div style={{fontSize:11,color:"#64748b",fontWeight:600}}>Total Geral</div>
              </div>
              <div style={{background:"#f0fdf4",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:26,marginBottom:4}}>🎯</div>
                <div style={{fontSize:22,fontWeight:800,color:"#059669"}}>{fmtBRL(faturamento.reduce((s,f)=>s+Number(f.valor),0)/(allPedidos.length||1))}</div>
                <div style={{fontSize:11,color:"#64748b",fontWeight:600}}>Ticket Médio Geral</div>
                {totalDev>0&&<div style={{fontSize:11,fontWeight:700,color:"#ef4444",marginTop:4}}>Devol.: -{fmtBRL(totalDev)}</div>}
              </div>
            </div>
            <div style={{background:"#fff",borderRadius:18,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
              <div style={{fontSize:14,fontWeight:700,color:"#1e293b",marginBottom:16}}>💰 Faturamento Mensal por Loja</div>
              {chartData.length>0?(
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={chartData} margin={{top:0,right:10,left:10,bottom:0}}>
                    <XAxis dataKey="mes" tick={{fontSize:11}} />
                    <YAxis tick={{fontSize:11}} tickFormatter={v=>v>=1000?`R$${(v/1000).toFixed(1)}k`:`R$${v}`} />
                    <Tooltip formatter={(v,n)=>[fmtBRL(v),n]} />
                    <Legend />
                    <Bar dataKey="Gran Shop"   fill="#1d4ed8" radius={[4,4,0,0]} />
                    <Bar dataKey="Aishael Mix" fill="#7c3aed" radius={[4,4,0,0]} />
                  </BarChart>
                </ResponsiveContainer>
              ):<div style={{textAlign:"center",padding:40,color:"#94a3b8",fontSize:13}}>Carregue planilhas para ver o faturamento.</div>}
            </div>
          </div>
        )}
      </div>

      {showConfirmClear&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:400,boxShadow:"0 20px 60px rgba(0,0,0,0.25)",textAlign:"center"}}>
            <div style={{fontSize:44,marginBottom:12}}>⚠️</div>
            <h3 style={{margin:"0 0 8px",fontSize:17,fontWeight:700,color:"#1e293b"}}>Zerar o Sistema?</h3>
            <p style={{margin:"0 0 24px",fontSize:13,color:"#64748b",lineHeight:1.6}}>
              Isso vai apagar <strong>todos os pedidos, devoluções e faturamento</strong> salvos no banco de dados.<br/>
              <strong style={{color:"#ef4444"}}>Essa ação não pode ser desfeita.</strong>
            </p>
            <div style={{display:"flex",gap:10,justifyContent:"center"}}>
              <button onClick={()=>setShowConfirmClear(false)} style={{padding:"10px 24px",border:"1px solid #e2e8f0",borderRadius:12,background:"#fff",fontSize:13,fontWeight:600,cursor:"pointer",color:"#374151"}}>
                Cancelar
              </button>
              <button onClick={handleClearAll} style={{padding:"10px 24px",border:"none",borderRadius:12,background:"#ef4444",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                🗑️ Sim, Zerar Tudo
              </button>
            </div>
          </div>
        </div>
      )}
      {showFinanceGate&&<FinanceGate onUnlock={()=>{setFinanceUnlocked(true);setActiveTab("financeiro");setShowFinanceGate(false);}} />}
      {revisaoModal&&<RevisaoModal order={revisaoModal} onConfirm={nota=>{handleStatusChange(revisaoModal,"revisao",nota);setRevisaoModal(null);}} onClose={()=>setRevisaoModal(null)} />}
      {toast&&<Toast msg={toast.msg} color={toast.color} />}
    </div>
  );
}
