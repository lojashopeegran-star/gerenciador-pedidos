import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { supabase, fetchPedidos, upsertPedidos, deletePedidos, updatePedidoStatus } from "./supabase.js";

const COLUMN_MAP = {
  "Nº de Pedido da Plataforma": "idPlataforma",
  "Nº do Subpedido": "subpedido",
  "Nº de Pedido": "numeroPedido",
  "Plataformas": "plataforma",
  "Nome da Loja no UpSeller": "loja",
  "Estado do Pedido": "estado",
  "Hora do Pedido": "horaPedido",
  "Hora do Pagamento": "horaPagamento",
  "Prazo de Envio": "prazoEnvio",
  "Notas do Comprador": "notas",
  "Nome do Anúncio": "produto",
  "Variação": "variacao",
  "Qtd. do Produto": "quantidade",
  "Unidade*": "unidade",
  "Nome de Comprador": "nomeComprador",
  "ID do Comprador": "idComprador",
  "Nome do Destinatário": "destinatario",
};

const DISPLAY_COLS = [
  { key: "idPlataforma",  label: "ID Plataforma" },
  { key: "loja",          label: "Loja" },
  { key: "produto",       label: "Produto" },
  { key: "variacao",      label: "Variação" },
  { key: "quantidade",    label: "Qtd" },
  { key: "prazoEnvio",    label: "Prazo Envio" },
  { key: "notas",         label: "Notas Comprador" },
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function deadlineInfo(prazoStr) {
  if (!prazoStr) return null;
  const deadline = new Date(prazoStr.replace(" ", "T"));
  if (isNaN(deadline)) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  const d = new Date(deadline); d.setHours(0,0,0,0);
  const diff = Math.round((d - today) / 86400000);
  if (diff <= 1) return { bg:"#fde8e8", border:"#ef4444", text:"#991b1b", icon:"🔴", label: diff<=0?"Vencido!":"Amanhã", tier:"red" };
  if (diff === 2) return { bg:"#fef9c3", border:"#f59e0b", text:"#78350f", icon:"🟡", label:`${diff}d`, tier:"yellow" };
  return { bg:"#dcfce7", border:"#22c55e", text:"#14532d", icon:"🟢", label:`${diff}d`, tier:"green" };
}

function parseRows(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return raw.map(row => {
    const m = {};
    for (const [orig, key] of Object.entries(COLUMN_MAP)) {
      m[key] = row[orig] !== undefined ? String(row[orig]) : "";
      if (!m[key]) {
        const found = Object.keys(row).find(k => k.trim().toLowerCase() === orig.trim().toLowerCase());
        if (found) m[key] = String(row[found]);
      }
    }
    m.statusInterno = ""; m.notaRevisao = ""; m._status = "new";
    return m;
  });
}

function mergeOrders(existing, incoming) {
  const map = new Map();
  const stats = { added: 0, ignored: 0 };
  for (const r of existing) map.set(r.idPlataforma, { ...r, _status:"existing" });
  for (const r of incoming) {
    if (map.has(r.idPlataforma)) stats.ignored++;
    else { map.set(r.idPlataforma, { ...r, _status:"new" }); stats.added++; }
  }
  return { rows: Array.from(map.values()), stats };
}

function applyFinalSheet(current, finalRows) {
  const toRemove = new Set(finalRows.filter(r => r.estado.trim().toLowerCase() === "retirada").map(r => r.idPlataforma));
  return {
    kept: current.filter(r => !toRemove.has(r.idPlataforma)),
    removedRows: current.filter(r => toRemove.has(r.idPlataforma)),
    removedCount: toRemove.size,
  };
}

function exportToExcel(rows) {
  const clean = rows.map(r => { const o={}; for(const [orig,key] of Object.entries(COLUMN_MAP)) o[orig]=r[key]||""; return o; });
  const ws = XLSX.utils.json_to_sheet(clean);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pedidos");
  XLSX.writeFile(wb, `pedidos_${Date.now()}.xlsx`);
}

// ── DropZone ──────────────────────────────────────────────────────────────────
function DropZone({ label, sublabel, onFile, file, color, disabled }) {
  const inputRef = useRef();
  const [drag, setDrag] = useState(false);
  const handle = f => {
    if (!f) return;
    const reader = new FileReader();
    reader.onload = e => { const wb = XLSX.read(e.target.result,{type:"binary"}); onFile(parseRows(wb.Sheets[wb.SheetNames[0]]),f.name); };
    reader.readAsBinaryString(f);
  };
  return (
    <div onClick={() => !disabled && inputRef.current.click()}
      onDragOver={e=>{if(!disabled){e.preventDefault();setDrag(true);}}}
      onDragLeave={()=>setDrag(false)}
      onDrop={e=>{e.preventDefault();setDrag(false);if(!disabled)handle(e.dataTransfer.files[0]);}}
      style={{border:`2px dashed ${drag?color:disabled?"#e2e8f0":"#cbd5e1"}`,borderRadius:14,padding:"22px 16px",cursor:disabled?"not-allowed":"pointer",background:disabled?"#f8fafc":drag?`${color}18`:"#f8fafc",textAlign:"center",minHeight:130,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:6,opacity:disabled?0.5:1,transition:"all 0.2s"}}>
      <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>handle(e.target.files[0])} />
      <div style={{fontSize:28}}>{disabled?"🔒":"📋"}</div>
      <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{label}</div>
      {sublabel && <div style={{fontSize:11,color:"#94a3b8",maxWidth:190}}>{sublabel}</div>}
      {file ? <div style={{background:color,color:"#fff",borderRadius:20,padding:"3px 12px",fontSize:11,fontWeight:600}}>✓ {file}</div>
             : <div style={{fontSize:11,color:"#94a3b8"}}>Arraste ou clique</div>}
    </div>
  );
}

// ── Toast ─────────────────────────────────────────────────────────────────────
function Toast({msg,color}) {
  if(!msg) return null;
  return <div style={{position:"fixed",bottom:28,right:28,zIndex:9999,background:color||"#1e293b",color:"#fff",borderRadius:12,padding:"12px 20px",fontSize:13,fontWeight:600,boxShadow:"0 8px 32px rgba(0,0,0,0.2)"}}>{msg}</div>;
}

// ── Revisão Modal ─────────────────────────────────────────────────────────────
function RevisaoModal({ order, onConfirm, onClose }) {
  const [nota, setNota] = useState(order.notaRevisao || "");
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:460,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
        <h3 style={{margin:"0 0 6px",fontSize:16,fontWeight:700,color:"#1e293b"}}>📋 Enviar para Revisão</h3>
        <p style={{margin:"0 0 16px",fontSize:12,color:"#64748b"}}>Pedido: <strong>{order.idPlataforma}</strong></p>
        <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Motivo da revisão *</label>
        <textarea
          value={nota} onChange={e=>setNota(e.target.value)}
          placeholder="Descreva o motivo pelo qual este pedido precisa de revisão..."
          style={{width:"100%",minHeight:100,padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,resize:"vertical",outline:"none",boxSizing:"border-box",fontFamily:"inherit"}}
          autoFocus
        />
        <div style={{display:"flex",gap:10,marginTop:16,justifyContent:"flex-end"}}>
          <button onClick={onClose} style={{padding:"9px 20px",border:"1px solid #e2e8f0",borderRadius:10,background:"#fff",fontSize:13,cursor:"pointer",color:"#374151",fontWeight:500}}>Cancelar</button>
          <button onClick={()=>onConfirm(nota)} disabled={!nota.trim()} style={{padding:"9px 20px",border:"none",borderRadius:10,background:nota.trim()?"#f59e0b":"#e2e8f0",color:nota.trim()?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:nota.trim()?"pointer":"not-allowed"}}>Confirmar Revisão</button>
        </div>
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App({ user, onLogout }) {
  const [orders,        setOrders]        = useState([]);
  const [planilha2Name, setPlanilha2Name] = useState(null);
  const [finalName,     setFinalName]     = useState(null);
  const [mergeStats,    setMergeStats]    = useState(null);
  const [finalStats,    setFinalStats]    = useState(null);
  const [toast,         setToast]         = useState(null);
  const [dbLoading,     setDbLoading]     = useState(true);
  const [saving,        setSaving]        = useState(false);
  const [search,        setSearch]        = useState("");
  const [filterStatus,  setFilterStatus]  = useState("all");
  const [filterLoja,    setFilterLoja]    = useState("all");
  const [filterPrazo,   setFilterPrazo]   = useState("all");
  const [filterInterno, setFilterInterno] = useState("all");
  const [filterProduto,  setFilterProduto]  = useState("all");
  const tableRef = useRef(null);
  const [selected,      setSelected]      = useState(new Set());
  const [copied,        setCopied]        = useState(false);
  const [revisaoModal,  setRevisaoModal]  = useState(null); // order object

  const showToast = (msg, color, ms=3500) => { setToast({msg,color}); setTimeout(()=>setToast(null),ms); };

  // Load from DB on mount
  useEffect(() => {
    if (!supabase || !user) { setDbLoading(false); return; }
    fetchPedidos(user.id).then(rows => { setOrders(rows); setDbLoading(false); });
  }, [user]);

  // ── Merge new sheet into existing orders ──────────────────────────────────
  const handleMerge = async (rows, name) => {
    setPlanilha2Name(name);
    const { rows: merged, stats } = mergeOrders(orders, rows);
    setOrders(merged);
    setMergeStats(stats);
    showToast(`✅ +${stats.added} novos pedidos, ${stats.ignored} ignorados`, "#059669");
    if (supabase && stats.added > 0) {
      setSaving(true);
      await upsertPedidos(merged.filter(r=>r._status==="new"), user.id);
      setSaving(false);
    }
  };

  // ── Apply final sheet (remove retirada) ───────────────────────────────────
  const handleFinalSheet = async (rows, name) => {
    setFinalName(name);
    const { kept, removedRows, removedCount } = applyFinalSheet(orders, rows);
    if (removedCount === 0) { showToast("ℹ️ Nenhum pedido 'Retirada' encontrado.", "#d97706"); return; }
    const withFlash = [...kept, ...removedRows.map(r=>({...r,_status:"removed"}))];
    setOrders(withFlash);
    setFinalStats({ removedCount, keptCount: kept.length });
    showToast(`🗑️ ${removedCount} pedido(s) removidos`, "#ef4444", 3000);
    setTimeout(() => setOrders(kept), 2500);
    if (supabase) { setSaving(true); await deletePedidos(removedRows.map(r=>r.idPlataforma), user.id); setSaving(false); }
  };

  // ── Status buttons: revisão / feito ──────────────────────────────────────
  const handleStatusChange = async (order, newStatus, nota="") => {
    const updated = orders.map(o => o.idPlataforma === order.idPlataforma ? {...o, statusInterno:newStatus, notaRevisao:nota} : o);
    setOrders(updated);
    if (supabase) await updatePedidoStatus(order.idPlataforma, user.id, newStatus, nota);
    showToast(newStatus==="feito" ? "✅ Pedido marcado como Feito!" : "📋 Pedido enviado para Revisão", newStatus==="feito"?"#059669":"#f59e0b");
  };

  // ── Selection ─────────────────────────────────────────────────────────────
  const toggleSelect = id => {
    setSelected(prev => { const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s; });
  };
  const toggleAll = () => {
    if (selected.size === filtered.length) setSelected(new Set());
    else setSelected(new Set(filtered.map(r=>r.idPlataforma)));
  };
  const copySelected = () => {
    const ids = [...selected].join(",");
    navigator.clipboard.writeText(ids).then(() => { setCopied(true); setTimeout(()=>setCopied(false),2000); showToast("📋 IDs copiados!", "#1d4ed8"); });
  };

  // ── Filters ───────────────────────────────────────────────────────────────
  const lojas = ["all", ...new Set(orders.map(r=>r.loja).filter(Boolean))];

  // Product list with counts (based on current non-removed orders)
  const activeOrders = orders.filter(r=>r._status!=="removed");
  const produtoMap = activeOrders.reduce((acc,r)=>{
    if(r.produto){ acc[r.produto]=(acc[r.produto]||0)+1; }
    return acc;
  },{});
  const produtos = Object.entries(produtoMap).sort((a,b)=>b[1]-a[1]); // sorted by count desc
  const filtered = orders.filter(r => {
    const dl = deadlineInfo(r.prazoEnvio);
    const matchSearch = !search || r.idPlataforma.toLowerCase().includes(search.toLowerCase()) || r.produto.toLowerCase().includes(search.toLowerCase()) || r.loja.toLowerCase().includes(search.toLowerCase());
    return (
      r._status !== "removed" &&
      (filterStatus==="all" || r._status===filterStatus) &&
      (filterLoja==="all"   || r.loja===filterLoja) &&
      (filterPrazo==="all"  || (dl && dl.tier===filterPrazo)) &&
      (filterInterno==="all"|| r.statusInterno===filterInterno) &&
      (filterProduto==="all" || r.produto===filterProduto) &&
      matchSearch
    );
  });

  const urgentCount = orders.filter(r=>{ const d=deadlineInfo(r.prazoEnvio); return d?.tier==="red" && r._status!=="removed"; }).length;
  const feitoCount  = orders.filter(r=>r.statusInterno==="feito").length;
  const revisaoCount= orders.filter(r=>r.statusInterno==="revisao").length;

  if (dbLoading) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#1e3a8a,#3b82f6)",fontFamily:"sans-serif"}}>
      <div style={{color:"#fff",fontSize:16,fontWeight:600}}>📦 Carregando seus pedidos...</div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#f0f4ff",fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}} * { box-sizing: border-box; }`}</style>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#1e3a8a 0%,#1d4ed8 60%,#3b82f6 100%)",padding:"24px 32px",color:"#fff"}}>
        <div style={{maxWidth:1500,margin:"0 auto",display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:11,fontWeight:600,letterSpacing:2,opacity:0.7}}>GERENCIADOR DE PEDIDOS</div>
            <h1 style={{margin:"2px 0 0",fontSize:24,fontWeight:800}}>Unificador de Planilhas 📦</h1>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{background:"rgba(255,255,255,0.15)",borderRadius:20,padding:"4px 14px",fontSize:12,display:"flex",alignItems:"center",gap:6}}>
              <span style={{width:7,height:7,borderRadius:"50%",background:saving?"#fbbf24":"#4ade80",display:"inline-block"}}/>
              {saving?"Salvando...":"Conectado"}
            </span>
            <span style={{background:"rgba(255,255,255,0.12)",borderRadius:20,padding:"4px 14px",fontSize:12}}>👤 {user?.email}</span>
            <button onClick={onLogout} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"4px 16px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Sair →</button>
          </div>
        </div>
      </div>

      <div style={{maxWidth:1500,margin:"0 auto",padding:"24px 32px"}}>

        {/* ── Upload cards ── */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,marginBottom:20}}>
          {/* Nova planilha */}
          <div style={{background:"#fff",borderRadius:16,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#1d4ed8",marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>📥 Adicionar Pedidos</div>
            <DropZone label="Nova Planilha" sublabel="Novos pedidos serão adicionados (duplicatas ignoradas)" color="#1d4ed8" file={planilha2Name} onFile={handleMerge} />
          </div>
          {/* Planilha final */}
          <div style={{background:"#fff",borderRadius:16,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#dc2626",marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>🗑️ Remover Retiradas</div>
            <DropZone label="Planilha Final" sublabel={'Pedidos com "Estado = Retirada" são removidos'} color="#dc2626" file={finalName} onFile={handleFinalSheet} />
          </div>
          {/* Export */}
          <div style={{background:"#fff",borderRadius:16,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
            <div style={{fontSize:12,fontWeight:700,color:"#059669",marginBottom:10,textTransform:"uppercase",letterSpacing:0.5}}>⬇️ Exportar</div>
            <div style={{display:"flex",flexDirection:"column",gap:8,height:"calc(100% - 30px)",justifyContent:"center"}}>
              <button onClick={()=>exportToExcel(orders.filter(r=>r._status!=="removed"))} style={{background:"#059669",color:"#fff",border:"none",borderRadius:10,padding:"11px",fontSize:13,fontWeight:700,cursor:"pointer",width:"100%"}}>
                📥 Exportar todos os pedidos
              </button>
              <button onClick={()=>exportToExcel(orders.filter(r=>r.statusInterno==="feito"))} style={{background:"#f0fdf4",color:"#065f46",border:"1px solid #86efac",borderRadius:10,padding:"11px",fontSize:13,fontWeight:600,cursor:"pointer",width:"100%"}}>
                ✅ Exportar apenas Feitos
              </button>
            </div>
          </div>
        </div>

        {/* ── Stats (clickable) ── */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:20}}>
          {[
            {label:"Total de Pedidos", value:activeOrders.length, color:"#1d4ed8", icon:"📦", action:()=>{setFilterInterno("all");setFilterPrazo("all");setFilterProduto("all");setSearch("");}},
            {label:"Novos (última carga)", value:mergeStats?.added??0, color:"#7c3aed", icon:"✨", action:()=>{setFilterInterno("all");setFilterPrazo("all");setFilterProduto("all");setFilterStatus("new");}},
            {label:"Feitos", value:feitoCount, color:"#059669", icon:"✅", action:()=>{setFilterInterno("feito");setFilterPrazo("all");setFilterProduto("all");setSearch("");}},
            {label:"Em Revisão", value:revisaoCount, color:"#f59e0b", icon:"📋", action:()=>{setFilterInterno("revisao");setFilterPrazo("all");setFilterProduto("all");setSearch("");}},
            {label:"Urgentes (≤1 dia)", value:urgentCount, color:"#ef4444", icon:"🔴", action:()=>{setFilterInterno("all");setFilterPrazo("red");setFilterProduto("all");setSearch("");}},
          ].map(s=>(
            <div key={s.label} onClick={()=>{s.action();tableRef.current?.scrollIntoView({behavior:"smooth"});}} style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",borderLeft:`4px solid ${s.color}`,cursor:"pointer",transition:"transform 0.15s,box-shadow 0.15s"}}
              onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 6px 20px rgba(0,0,0,0.12)";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,0.07)";}}>
              <div style={{fontSize:18,marginBottom:2}}>{s.icon}</div>
              <div style={{fontSize:24,fontWeight:800,color:s.color}}>{s.value}</div>
              <div style={{fontSize:11,color:"#64748b",fontWeight:500}}>{s.label}</div>
              <div style={{fontSize:10,color:s.color,marginTop:3,opacity:0.7}}>clique para filtrar →</div>
            </div>
          ))}
        </div>

        {/* ── Product Panel ── */}
        <div style={{background:"#fff",borderRadius:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20,overflow:"hidden"}}>
          <div style={{padding:"14px 20px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:14,fontWeight:700,color:"#1e293b"}}>📦 Filtrar por Produto</span>
            <span style={{fontSize:11,color:"#94a3b8"}}>{produtos.length} produto{produtos.length!==1?"s":""} encontrado{produtos.length!==1?"s":""}</span>
            {filterProduto!=="all" && (
              <button onClick={()=>setFilterProduto("all")} style={{marginLeft:"auto",background:"none",border:"1px solid #7c3aed",color:"#7c3aed",borderRadius:8,padding:"3px 12px",fontSize:11,cursor:"pointer",fontWeight:600}}>
                ✕ Limpar filtro
              </button>
            )}
          </div>
          <div style={{padding:"14px 20px",display:"flex",flexWrap:"wrap",gap:8}}>
            {produtos.map(([nome,qtd])=>{
              const isActive = filterProduto===nome;
              const short = nome.length>55 ? nome.slice(0,55)+"..." : nome;
              return (
                <button
                  key={nome}
                  onClick={()=>{ setFilterProduto(isActive?"all":nome); tableRef.current?.scrollIntoView({behavior:"smooth"}); }}
                  title={nome}
                  style={{
                    display:"flex", alignItems:"center", gap:6,
                    background: isActive ? "#7c3aed" : "#f8fafc",
                    color: isActive ? "#fff" : "#374151",
                    border: isActive ? "1.5px solid #7c3aed" : "1.5px solid #e2e8f0",
                    borderRadius:10, padding:"7px 12px",
                    fontSize:12, fontWeight: isActive?700:500,
                    cursor:"pointer", transition:"all 0.15s",
                    textAlign:"left",
                  }}
                  onMouseEnter={e=>{ if(!isActive){ e.currentTarget.style.borderColor="#7c3aed"; e.currentTarget.style.color="#7c3aed"; }}}
                  onMouseLeave={e=>{ if(!isActive){ e.currentTarget.style.borderColor="#e2e8f0"; e.currentTarget.style.color="#374151"; }}}
                >
                  <span style={{
                    background: isActive?"rgba(255,255,255,0.25)":"#7c3aed",
                    color:"#fff", borderRadius:20,
                    padding:"1px 7px", fontSize:11, fontWeight:700, flexShrink:0,
                  }}>{qtd}</span>
                  {short}
                </button>
              );
            })}
            {produtos.length===0 && (
              <span style={{fontSize:12,color:"#94a3b8",padding:"8px 0"}}>Nenhum pedido carregado ainda.</span>
            )}
          </div>

          {/* When a product is selected — show the order list for that product */}
          {filterProduto!=="all" && (
            <div style={{borderTop:"1px solid #f1f5f9",padding:"0 20px 16px"}}>
              <div style={{padding:"12px 0 10px",display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:13,fontWeight:700,color:"#5b21b6"}}>📋 Pedidos de:</span>
                <span style={{fontSize:12,color:"#4c1d95",fontStyle:"italic",flex:1}}>{filterProduto}</span>
                <span style={{background:"#7c3aed",color:"#fff",borderRadius:20,padding:"3px 14px",fontSize:12,fontWeight:700}}>
                  {orders.filter(r=>r.produto===filterProduto&&r._status!=="removed").length} pedido{orders.filter(r=>r.produto===filterProduto&&r._status!=="removed").length!==1?"s":""}
                </span>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {orders.filter(r=>r.produto===filterProduto&&r._status!=="removed").map((r,i)=>{
                  const dl = deadlineInfo(r.prazoEnvio);
                  const si = r.statusInterno;
                  return (
                    <div key={r.idPlataforma} style={{
                      display:"flex", alignItems:"center", gap:10, flexWrap:"wrap",
                      background: si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#f8fafc":"#fff",
                      borderRadius:10, padding:"9px 14px",
                      border:"1px solid",
                      borderColor: si==="feito"?"#86efac":si==="revisao"?"#fde68a":"#f1f5f9",
                    }}>
                      <input type="checkbox" checked={selected.has(r.idPlataforma)} onChange={()=>toggleSelect(r.idPlataforma)}
                        style={{cursor:"pointer",width:14,height:14,flexShrink:0}} />
                      <span style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1d4ed8",flexShrink:0}}>{r.idPlataforma}</span>
                      <span style={{fontSize:11,color:"#64748b",flexShrink:0}}>Loja: <strong>{r.loja||"—"}</strong></span>
                      <span style={{fontSize:11,color:"#64748b",flexShrink:0}}>Qtd: <strong>{r.quantidade||"—"}</strong></span>
                      <span style={{fontSize:11,color:"#64748b",flex:1}}>{r.variacao||""}</span>
                      {dl && (
                        <span style={{background:dl.bg,color:dl.text,border:`1px solid ${dl.border}`,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:600,flexShrink:0}}>
                          {dl.icon} {dl.label}
                        </span>
                      )}
                      {si==="feito" && <span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700,flexShrink:0}}>✅ Feito</span>}
                      {si==="revisao" && <span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 8px",fontSize:11,fontWeight:700,flexShrink:0}}>📋 Revisão</span>}
                      {!si && <span style={{background:"#f1f5f9",color:"#64748b",borderRadius:20,padding:"2px 8px",fontSize:11,flexShrink:0}}>⏳ Pendente</span>}
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* ── Filters & Table ── */}
        <div ref={tableRef} style={{background:"#fff",borderRadius:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
          {/* Filter bar */}
          <div style={{padding:"16px 20px",borderBottom:"1px solid #f1f5f9",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <input placeholder="🔍 Buscar por ID, produto, loja..." value={search} onChange={e=>setSearch(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"8px 12px",fontSize:12,outline:"none",flex:1,minWidth:200}} />
            <select value={filterInterno} onChange={e=>setFilterInterno(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"8px 12px",fontSize:12,cursor:"pointer",background:"#fff"}}>
              <option value="all">Todos os estados</option>
              <option value="">⏳ Pendentes</option>
              <option value="feito">✅ Feito</option>
              <option value="revisao">📋 Revisão</option>
            </select>
            <select value={filterLoja} onChange={e=>setFilterLoja(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"8px 12px",fontSize:12,cursor:"pointer",background:"#fff"}}>
              {lojas.map(l=><option key={l} value={l}>{l==="all"?"Todas as lojas":l}</option>)}
            </select>
            <select value={filterPrazo} onChange={e=>setFilterPrazo(e.target.value)}
              style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"8px 12px",fontSize:12,cursor:"pointer",background:"#fff"}}>
              <option value="all">Todos os prazos</option>
              <option value="red">🔴 Urgente (≤1 dia)</option>
              <option value="yellow">🟡 Atenção (2 dias)</option>
              <option value="green">🟢 OK (3+ dias)</option>
            </select>
{/* produto filter handled by panel below */}
          </div>

          {/* Selection bar */}
          <div style={{padding:"10px 20px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"#64748b"}}>
              <strong>{filtered.length}</strong> pedidos &nbsp;·&nbsp;
              <strong>{selected.size}</strong> selecionados
            </span>
            {selected.size > 0 && (
              <button onClick={copySelected} style={{
                background: copied?"#059669":"#1d4ed8", color:"#fff", border:"none",
                borderRadius:10, padding:"6px 16px", fontSize:12, fontWeight:700, cursor:"pointer",
                display:"flex", alignItems:"center", gap:6,
              }}>
                {copied ? "✓ Copiado!" : `📋 Copiar ${selected.size} ID(s) para UpSeller`}
              </button>
            )}
            {selected.size > 0 && (
              <button onClick={()=>setSelected(new Set())} style={{background:"none",border:"1px solid #e2e8f0",borderRadius:10,padding:"6px 12px",fontSize:12,cursor:"pointer",color:"#64748b"}}>
                Limpar seleção
              </button>
            )}
            {/* Legend */}
            <div style={{marginLeft:"auto",display:"flex",gap:8,flexWrap:"wrap"}}>
              {[{icon:"🟢",label:"3+d",bg:"#dcfce7",text:"#14532d",border:"#22c55e"},{icon:"🟡",label:"2d",bg:"#fef9c3",text:"#78350f",border:"#f59e0b"},{icon:"🔴",label:"≤1d",bg:"#fde8e8",text:"#991b1b",border:"#ef4444"}].map(l=>(
                <span key={l.label} style={{background:l.bg,color:l.text,border:`1px solid ${l.border}`,borderRadius:8,padding:"2px 8px",fontSize:11,fontWeight:600}}>{l.icon} {l.label}</span>
              ))}
            </div>
          </div>

          {/* Copied IDs preview */}
          {selected.size > 0 && (
            <div style={{padding:"10px 20px",background:"#eff6ff",borderBottom:"1px solid #bfdbfe"}}>
              <div style={{fontSize:11,color:"#1d4ed8",fontWeight:600,marginBottom:4}}>IDs selecionados (para colar no UpSeller):</div>
              <div style={{fontSize:12,color:"#1e40af",fontFamily:"monospace",wordBreak:"break-all",background:"#dbeafe",borderRadius:8,padding:"8px 12px"}}>
                {[...selected].join(",")}
              </div>
            </div>
          )}

          {/* Table */}
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  <th style={{...TH,width:36}}>
                    <input type="checkbox" checked={selected.size>0&&selected.size===filtered.length} onChange={toggleAll}
                      style={{cursor:"pointer",width:14,height:14}} />
                  </th>
                  <th style={TH}>Estado</th>
                  {DISPLAY_COLS.map(c=><th key={c.key} style={TH}>{c.label}</th>)}
                  <th style={TH}>Ações</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((row,i)=>{
                  const dl = deadlineInfo(row.prazoEnvio);
                  const isSelected = selected.has(row.idPlataforma);
                  const si = row.statusInterno;
                  return (
                    <tr key={row.idPlataforma} style={{background:isSelected?"#eff6ff":si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f1f5f9",transition:"background 0.15s"}}>
                      {/* Checkbox */}
                      <td style={{...TD,textAlign:"center"}}>
                        <input type="checkbox" checked={isSelected} onChange={()=>toggleSelect(row.idPlataforma)}
                          style={{cursor:"pointer",width:14,height:14}} />
                      </td>
                      {/* Estado interno badge */}
                      <td style={{...TD,textAlign:"center"}}>
                        {si==="feito" ? (
                          <span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>✅ Feito</span>
                        ) : si==="revisao" ? (
                          <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                            <span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:700}}>📋 Revisão</span>
                            {row.notaRevisao && <span style={{fontSize:10,color:"#92400e",maxWidth:120,textAlign:"center",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={row.notaRevisao}>{row.notaRevisao}</span>}
                          </div>
                        ) : (
                          <span style={{background:"#f1f5f9",color:"#64748b",borderRadius:20,padding:"3px 10px",fontSize:11,fontWeight:500}}>⏳ Pendente</span>
                        )}
                      </td>
                      {/* Data cols */}
                      {DISPLAY_COLS.map(c=>{
                        const isPrazo = c.key==="prazoEnvio";
                        const dl2 = isPrazo ? deadlineInfo(row[c.key]) : null;
                        return (
                          <td key={c.key} style={{...TD, fontWeight:c.key==="idPlataforma"?700:400, color:c.key==="idPlataforma"?"#1d4ed8":"#374151"}}>
                            {isPrazo && dl2 ? (
                              <div style={{background:dl2.bg,border:`1px solid ${dl2.border}`,borderRadius:8,padding:"4px 8px",display:"inline-flex",flexDirection:"column",gap:1,minWidth:90}}>
                                <span style={{fontSize:11,color:dl2.text,fontWeight:700}}>{dl2.icon} {dl2.label}</span>
                                <span style={{fontSize:10,color:dl2.text,opacity:0.75}}>{row[c.key].slice(0,10)}</span>
                              </div>
                            ) : isPrazo ? (
                              <span style={{color:"#94a3b8",fontSize:11}}>—</span>
                            ) : (
                              <div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:c.key==="produto"?"normal":"nowrap",maxWidth:c.key==="produto"?200:150}}>
                                {row[c.key]||"—"}
                              </div>
                            )}
                          </td>
                        );
                      })}
                      {/* Action buttons */}
                      <td style={{...TD,whiteSpace:"nowrap"}}>
                        <div style={{display:"flex",gap:6}}>
                          <button
                            onClick={()=>setRevisaoModal(row)}
                            style={{padding:"5px 12px",border:"none",borderRadius:8,background:si==="revisao"?"#fef3c7":"#f1f5f9",color:si==="revisao"?"#92400e":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                            📋 Revisão
                          </button>
                          <button
                            onClick={()=>handleStatusChange(row, si==="feito"?"":"feito")}
                            style={{padding:"5px 12px",border:"none",borderRadius:8,background:si==="feito"?"#d1fae5":"#f1f5f9",color:si==="feito"?"#065f46":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                            {si==="feito"?"✅ Feito":"⬜ Feito"}
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {filtered.length===0 && (
                  <tr><td colSpan={DISPLAY_COLS.length+3} style={{textAlign:"center",padding:40,color:"#94a3b8",fontSize:14}}>
                    {orders.length===0 ? "📭 Nenhum pedido ainda. Carregue uma planilha para começar!" : "Nenhum pedido encontrado com esses filtros."}
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Revisão Modal */}
      {revisaoModal && (
        <RevisaoModal
          order={revisaoModal}
          onConfirm={nota=>{ handleStatusChange(revisaoModal,"revisao",nota); setRevisaoModal(null); }}
          onClose={()=>setRevisaoModal(null)}
        />
      )}

      {toast && <Toast msg={toast.msg} color={toast.color} />}
    </div>
  );
}

const TH = {padding:"10px 14px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748b",letterSpacing:0.5,textTransform:"uppercase",whiteSpace:"nowrap",borderBottom:"2px solid #e2e8f0"};
const TD = {padding:"9px 12px",verticalAlign:"middle"};
