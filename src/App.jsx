import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase, fetchPedidos, upsertPedidos, deletePedidos } from "./supabase.js";

// ── Column mapping ────────────────────────────────────────────────────────────
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
  { key: "numeroPedido",  label: "Nº Pedido" },
  { key: "idPlataforma",  label: "ID Plataforma" },
  { key: "loja",          label: "Loja" },
  { key: "estado",        label: "Estado" },
  { key: "produto",       label: "Produto" },
  { key: "variacao",      label: "Variação" },
  { key: "quantidade",    label: "Qtd" },
  { key: "horaPedido",    label: "Data Pedido" },
  { key: "prazoEnvio",    label: "Prazo Envio" },
  { key: "notas",         label: "Notas" },
];

// ── Deadline color ────────────────────────────────────────────────────────────
function deadlineInfo(prazoStr) {
  if (!prazoStr || prazoStr === "—") return null;
  const deadline = new Date(prazoStr.replace(" ", "T"));
  if (isNaN(deadline)) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  const d = new Date(deadline); d.setHours(0,0,0,0);
  const diff = Math.round((d - today) / 86400000);
  if (diff <= 1) return { bg:"#fde8e8", border:"#ef4444", text:"#991b1b", icon:"🔴", label: diff<=0?"Vencido!":"Vence amanhã", tier:"red" };
  if (diff === 2) return { bg:"#fef9c3", border:"#f59e0b", text:"#78350f", icon:"🟡", label:`${diff}d restantes`, tier:"yellow" };
  return              { bg:"#dcfce7", border:"#22c55e", text:"#14532d", icon:"🟢", label:`${diff}d restantes`, tier:"green" };
}

// ── Parse xlsx sheet → array of mapped objects ────────────────────────────────
function parseRows(sheet) {
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return raw.map((row) => {
    const m = {};
    for (const [orig, key] of Object.entries(COLUMN_MAP)) {
      m[key] = row[orig] !== undefined ? String(row[orig]) : "";
      if (!m[key]) {
        const found = Object.keys(row).find(k => k.trim().toLowerCase() === orig.trim().toLowerCase());
        if (found) m[key] = String(row[found]);
      }
    }
    return m;
  });
}

// ── Merge: existing base + new sheet → skip duplicates ───────────────────────
function mergeOrders(existing, incoming) {
  const map = new Map();
  const stats = { added: 0, ignored: 0 };
  for (const r of existing) map.set(r.idPlataforma, { ...r, _status: "existing" });
  for (const r of incoming) {
    if (map.has(r.idPlataforma)) stats.ignored++;
    else { map.set(r.idPlataforma, { ...r, _status: "new" }); stats.added++; }
  }
  return { rows: Array.from(map.values()), stats };
}

// ── Apply final sheet: remove orders with estado = "retirada" ────────────────
function applyFinalSheet(current, finalRows) {
  const toRemove = new Set(
    finalRows
      .filter(r => r.estado.trim().toLowerCase() === "retirada")
      .map(r => r.idPlataforma)
  );
  return {
    kept:         current.filter(r => !toRemove.has(r.idPlataforma)),
    removedRows:  current.filter(r =>  toRemove.has(r.idPlataforma)),
    removedCount: toRemove.size,
  };
}

// ── Export to xlsx ────────────────────────────────────────────────────────────
function exportToExcel(rows, filename = `pedidos_unificados_${Date.now()}.xlsx`) {
  const clean = rows.map(r => {
    const out = {};
    for (const [orig, key] of Object.entries(COLUMN_MAP)) out[orig] = r[key] || "";
    return out;
  });
  const ws = XLSX.utils.json_to_sheet(clean);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pedidos");
  XLSX.writeFile(wb, filename);
}

// ── Status badge config ───────────────────────────────────────────────────────
const STATUS_CFG = {
  new:      { bg:"#d1fae5", text:"#065f46", label:"Novo" },
  existing: { bg:"#f3f4f6", text:"#374151", label:"Existente" },
  removed:  { bg:"#fde8e8", text:"#991b1b", label:"Removido ✕" },
};

// ── DropZone ──────────────────────────────────────────────────────────────────
function DropZone({ label, sublabel, onFile, file, color, disabled }) {
  const inputRef = useRef();
  const [drag, setDrag] = useState(false);
  const handle = f => {
    if (!f) return;
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type:"binary" });
      onFile(parseRows(wb.Sheets[wb.SheetNames[0]]), f.name);
    };
    reader.readAsBinaryString(f);
  };
  return (
    <div
      onClick={() => !disabled && inputRef.current.click()}
      onDragOver={e => { if (!disabled) { e.preventDefault(); setDrag(true); }}}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); if (!disabled) handle(e.dataTransfer.files[0]); }}
      style={{
        border:`2px dashed ${drag ? color : disabled ? "#e2e8f0" : "#cbd5e1"}`,
        borderRadius:16, padding:"28px 20px",
        cursor: disabled ? "not-allowed" : "pointer",
        background: disabled ? "#f8fafc" : drag ? `${color}18` : "#f8fafc",
        transition:"all 0.2s", textAlign:"center", minHeight:150,
        display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:8,
        opacity: disabled ? 0.5 : 1,
      }}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{display:"none"}}
        onChange={e => handle(e.target.files[0])} />
      <div style={{fontSize:32}}>{disabled ? "🔒" : "📋"}</div>
      <div style={{fontWeight:700, fontSize:14, color:"#1e293b"}}>{label}</div>
      {sublabel && <div style={{fontSize:11, color:"#94a3b8", maxWidth:200}}>{sublabel}</div>}
      {file
        ? <div style={{background:color, color:"#fff", borderRadius:20, padding:"4px 14px", fontSize:11, fontWeight:600}}>✓ {file}</div>
        : <div style={{fontSize:11, color:"#94a3b8"}}>Arraste ou clique para selecionar</div>
      }
    </div>
  );
}

// ── Toast notification ────────────────────────────────────────────────────────
function Toast({ msg, color }) {
  if (!msg) return null;
  return (
    <div style={{
      position:"fixed", bottom:32, right:32, zIndex:9999,
      background: color || "#1e293b", color:"#fff",
      borderRadius:14, padding:"14px 22px", fontSize:14, fontWeight:600,
      boxShadow:"0 8px 32px rgba(0,0,0,0.18)",
      animation:"fadeIn 0.3s ease",
    }}>
      {msg}
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [planilha1,     setPlanilha1]     = useState(null);
  const [planilha1Name, setPlanilha1Name] = useState(null);
  const [planilha2,     setPlanilha2]     = useState(null);
  const [planilha2Name, setPlanilha2Name] = useState(null);
  const [result,        setResult]        = useState(null);
  const [mergeStats,    setMergeStats]    = useState(null);
  const [finalStats,    setFinalStats]    = useState(null);
  const [finalName,     setFinalName]     = useState(null);
  const [toast,         setToast]         = useState(null);
  const [dbLoading,     setDbLoading]     = useState(!!supabase);
  const [saving,        setSaving]        = useState(false);
  const [search,        setSearch]        = useState("");
  const [filterStatus,  setFilterStatus]  = useState("all");
  const [filterLoja,    setFilterLoja]    = useState("all");
  const [filterPrazo,   setFilterPrazo]   = useState("all");

  useEffect(() => {
    if (!supabase) return;
    fetchPedidos().then(rows => {
      if (rows.length > 0) { setResult(rows); setMergeStats({ added: 0, ignored: 0 }); }
      setDbLoading(false);
    });
  }, []);

  const showToast = (msg, color, ms = 4000) => {
    setToast({ msg, color });
    setTimeout(() => setToast(null), ms);
  };

  // Step 1: merge planilha base + nova
  const handleMerge = async () => {
    if (!planilha1 || !planilha2) return;
    const base = result ? result.filter(r => r._status !== "removed") : [];
    const { rows, stats } = mergeOrders(base.length ? base : planilha1, planilha2);
    setResult(rows);
    setMergeStats(stats);
    setFinalStats(null);
    setFinalName(null);
    showToast(`✅ Unificado! +${stats.added} novos, ${stats.ignored} ignorados`, "#059669");
    if (supabase && stats.added > 0) {
      setSaving(true);
      const newRows = rows.filter(r => r._status === "new");
      await upsertPedidos(newRows);
      setSaving(false);
    }
  };

  // Step 2: apply final sheet (remove "retirada")
  const handleFinalSheet = async (rows, name) => {
    if (!result) return;
    const activeRows = result.filter(r => r._status !== "removed");
    const { kept, removedRows, removedCount } = applyFinalSheet(activeRows, rows);
    if (removedCount === 0) {
      showToast("ℹ️ Nenhum pedido com estado 'Retirada' encontrado.", "#d97706");
      return;
    }
    const withFlash = [...kept, ...removedRows.map(r => ({ ...r, _status:"removed" }))];
    setResult(withFlash);
    setFinalName(name);
    setFinalStats({ removedCount, keptCount: kept.length });
    showToast(`🗑️ ${removedCount} pedido(s) marcados como Retirada serão removidos.`, "#ef4444", 3000);
    setTimeout(() => setResult(kept), 2500);
    if (supabase && removedRows.length > 0) {
      setSaving(true);
      await deletePedidos(removedRows.map(r => r.idPlataforma));
      setSaving(false);
    }
  };

  const lojas = result ? ["all", ...new Set(result.map(r => r.loja).filter(Boolean))] : [];

  const filtered = result ? result.filter(r => {
    const dl = deadlineInfo(r.prazoEnvio);
    return (
      (filterStatus === "all" || r._status === filterStatus) &&
      (filterLoja   === "all" || r.loja === filterLoja) &&
      (filterPrazo  === "all" || (dl && dl.tier === filterPrazo)) &&
      (!search || Object.values(r).some(v => String(v).toLowerCase().includes(search.toLowerCase())))
    );
  }) : [];

  const urgentCount = result ? result.filter(r => { const d = deadlineInfo(r.prazoEnvio); return d?.tier === "red"; }).length : 0;
  const activeResult = result ? result.filter(r => r._status !== "removed") : [];

  return (
    <div style={{minHeight:"100vh", background:"#f0f4ff", fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
      <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}} @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}`}</style>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{background:"linear-gradient(135deg,#1e3a8a 0%,#1d4ed8 60%,#3b82f6 100%)", padding:"30px 40px 26px", color:"#fff"}}>
        <div style={{maxWidth:1400, margin:"0 auto"}}>
          <div style={{fontSize:12, fontWeight:600, letterSpacing:2, opacity:0.7, marginBottom:5}}>GERENCIADOR DE PEDIDOS</div>
          <h1 style={{margin:0, fontSize:28, fontWeight:800, letterSpacing:-0.5}}>Unificador de Planilhas 📦</h1>
          <p style={{margin:"6px 0 0", opacity:0.8, fontSize:13}}>Une planilhas automaticamente · Remove pedidos retirados · Alertas de prazo de envio</p>
          <div style={{marginTop:10, display:"flex", alignItems:"center", gap:8, fontSize:12}}>
            {supabase ? (
              <span style={{background:"rgba(255,255,255,0.15)", borderRadius:20, padding:"3px 12px", display:"flex", alignItems:"center", gap:6}}>
                <span style={{width:7, height:7, borderRadius:"50%", background: dbLoading?"#fbbf24":"#4ade80", display:"inline-block"}}/>
                {dbLoading ? "Conectando ao banco..." : saving ? "Salvando..." : "Banco de dados conectado"}
              </span>
            ) : (
              <span style={{background:"rgba(255,255,255,0.1)", borderRadius:20, padding:"3px 12px", opacity:0.7}}>
                ⚠️ Modo local (configure .env para persistência)
              </span>
            )}
          </div>
        </div>
      </div>

      <div style={{maxWidth:1400, margin:"0 auto", padding:"28px 40px"}}>

        {/* ── Section 1: Upload & Merge ── */}
        <div style={{background:"#fff", borderRadius:20, padding:26, marginBottom:20, boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
          <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:18}}>
            <div style={{background:"#1d4ed8", color:"#fff", borderRadius:8, width:26, height:26, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800, fontSize:13}}>1</div>
            <h2 style={{margin:0, fontSize:15, fontWeight:700, color:"#1e293b"}}>Unificação de Planilhas</h2>
          </div>
          <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:18}}>
            <DropZone label="Planilha Base (atual)" sublabel="Seus pedidos já salvos no sistema"
              color="#1d4ed8" file={planilha1Name}
              onFile={(rows, name) => { setPlanilha1(rows); setPlanilha1Name(name); setResult(null); setMergeStats(null); setFinalStats(null); }} />
            <DropZone label="Planilha Nova (atualização)" sublabel="Exportação mais recente para adicionar pedidos novos"
              color="#7c3aed" file={planilha2Name}
              onFile={(rows, name) => { setPlanilha2(rows); setPlanilha2Name(name); }} />
          </div>
          <div style={{display:"flex", alignItems:"center", gap:12, flexWrap:"wrap"}}>
            {planilha1 && planilha2 && (
              <span style={{fontSize:12, color:"#64748b"}}>
                Base: <strong>{planilha1.length}</strong> linhas · Nova: <strong>{planilha2.length}</strong> linhas
              </span>
            )}
            <button onClick={handleMerge} disabled={!planilha1 || !planilha2} style={{
              marginLeft:"auto",
              background: planilha1 && planilha2 ? "linear-gradient(135deg,#1d4ed8,#7c3aed)" : "#e2e8f0",
              color: planilha1 && planilha2 ? "#fff" : "#94a3b8",
              border:"none", borderRadius:12, padding:"11px 26px",
              fontSize:14, fontWeight:700, cursor: planilha1 && planilha2 ? "pointer" : "not-allowed",
              transition:"all 0.2s",
            }}>⚡ Unificar Planilhas</button>
          </div>
          <div style={{marginTop:14, background:"#f0f4ff", borderRadius:10, padding:"10px 14px", fontSize:12, color:"#3b82f6"}}>
            🔑 Pedidos com o mesmo <strong>Nº de Pedido da Plataforma</strong> já existente são ignorados (sem duplicatas).
          </div>
        </div>

        {/* ── Section 2: Final Sheet (remove retirada) ── */}
        <div style={{
          background:"#fff", borderRadius:20, padding:26, marginBottom:20,
          boxShadow:"0 1px 3px rgba(0,0,0,0.07)",
          border: result ? "2px solid #fef3c7" : "2px solid #f1f5f9",
        }}>
          <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:16}}>
            <div style={{background: result ? "#d97706" : "#94a3b8", color:"#fff", borderRadius:8, width:26, height:26, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800, fontSize:13}}>2</div>
            <h2 style={{margin:0, fontSize:15, fontWeight:700, color:"#1e293b"}}>Planilha Final — Remover Pedidos Retirados</h2>
            {!result && <span style={{fontSize:11, color:"#94a3b8", fontWeight:500}}>(disponível após a unificação)</span>}
          </div>

          <div style={{display:"grid", gridTemplateColumns:"1fr auto", gap:16, alignItems:"start"}}>
            <DropZone
              label="Planilha Final (retiradas)"
              sublabel='Pedidos com "Estado do Pedido = Retirada" serão removidos da lista atual'
              color="#dc2626"
              file={finalName}
              disabled={!result}
              onFile={handleFinalSheet}
            />
            <div style={{
              background:"#fff7ed", border:"1px solid #fed7aa", borderRadius:14,
              padding:"18px 20px", minWidth:200,
            }}>
              <div style={{fontSize:12, fontWeight:700, color:"#92400e", marginBottom:10}}>Como funciona:</div>
              {[
                ["📥", "Carregue a planilha final"],
                ["🔍", 'Sistema identifica "Estado = Retirada"'],
                ["🗑️", "Esses pedidos são removidos da lista"],
                ["✅", "Base atualizada automaticamente"],
              ].map(([icon, text]) => (
                <div key={text} style={{display:"flex", gap:8, alignItems:"flex-start", marginBottom:6, fontSize:11, color:"#78350f"}}>
                  <span>{icon}</span><span>{text}</span>
                </div>
              ))}
            </div>
          </div>

          {finalStats && (
            <div style={{
              marginTop:14, background:"#fde8e8", border:"1px solid #fca5a5",
              borderRadius:10, padding:"12px 16px",
              display:"flex", gap:24, alignItems:"center",
            }}>
              <span style={{fontSize:13, fontWeight:700, color:"#991b1b"}}>
                🗑️ {finalStats.removedCount} pedido(s) removidos (estado: Retirada)
              </span>
              <span style={{fontSize:13, color:"#065f46", fontWeight:600}}>
                ✅ {finalStats.keptCount} pedidos ativos restantes
              </span>
            </div>
          )}
        </div>

        {/* ── Stats cards ── */}
        {mergeStats && result && (
          <div style={{display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:14, marginBottom:20}}>
            {[
              { label:"Pedidos Ativos",       value: activeResult.length,  color:"#1d4ed8", icon:"📦" },
              { label:"Novos Adicionados",    value: mergeStats.added,     color:"#059669", icon:"✅" },
              { label:"Ignorados",            value: mergeStats.ignored,   color:"#6b7280", icon:"⏭️" },
              { label:"Urgentes (≤1 dia)",    value: urgentCount,          color:"#ef4444", icon:"🔴" },
            ].map(s => (
              <div key={s.label} style={{
                background:"#fff", borderRadius:14, padding:"18px 20px",
                boxShadow:"0 1px 3px rgba(0,0,0,0.07)", borderLeft:`4px solid ${s.color}`,
              }}>
                <div style={{fontSize:20, marginBottom:4}}>{s.icon}</div>
                <div style={{fontSize:26, fontWeight:800, color:s.color}}>{s.value}</div>
                <div style={{fontSize:11, color:"#64748b", fontWeight:500}}>{s.label}</div>
              </div>
            ))}
          </div>
        )}

        {/* ── Table ── */}
        {result && (
          <div style={{background:"#fff", borderRadius:20, boxShadow:"0 1px 3px rgba(0,0,0,0.07)", overflow:"hidden"}}>
            {/* Filters */}
            <div style={{padding:"18px 22px", borderBottom:"1px solid #f1f5f9", display:"flex", gap:10, alignItems:"center", flexWrap:"wrap"}}>
              <div style={{display:"flex", alignItems:"center", gap:8, marginRight:4}}>
                <div style={{background:"#1d4ed8", color:"#fff", borderRadius:8, width:26, height:26, display:"flex", alignItems:"center", justifyContent:"center", fontWeight:800, fontSize:13}}>3</div>
                <h2 style={{margin:0, fontSize:15, fontWeight:700, color:"#1e293b"}}>Visualizar e Exportar</h2>
              </div>
              <input placeholder="🔍 Buscar..." value={search} onChange={e => setSearch(e.target.value)}
                style={{border:"1px solid #e2e8f0", borderRadius:10, padding:"8px 13px", fontSize:12, outline:"none", flex:1, minWidth:160}} />
              <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
                style={{border:"1px solid #e2e8f0", borderRadius:10, padding:"8px 13px", fontSize:12, cursor:"pointer", background:"#fff"}}>
                <option value="all">Todos os status</option>
                <option value="new">✅ Novos</option>
                <option value="existing">Existentes</option>
                <option value="removed">🗑️ Removidos</option>
              </select>
              <select value={filterLoja} onChange={e => setFilterLoja(e.target.value)}
                style={{border:"1px solid #e2e8f0", borderRadius:10, padding:"8px 13px", fontSize:12, cursor:"pointer", background:"#fff"}}>
                {lojas.map(l => <option key={l} value={l}>{l === "all" ? "Todas as lojas" : l}</option>)}
              </select>
              <select value={filterPrazo} onChange={e => setFilterPrazo(e.target.value)}
                style={{border:"1px solid #e2e8f0", borderRadius:10, padding:"8px 13px", fontSize:12, cursor:"pointer", background:"#fff"}}>
                <option value="all">Todos os prazos</option>
                <option value="red">🔴 Urgente (≤1 dia)</option>
                <option value="yellow">🟡 Atenção (2 dias)</option>
                <option value="green">🟢 OK (3+ dias)</option>
              </select>
              <button onClick={() => exportToExcel(activeResult)}
                style={{background:"#059669", color:"#fff", border:"none", borderRadius:10, padding:"8px 18px", fontSize:12, fontWeight:700, cursor:"pointer"}}>
                ⬇️ Exportar .xlsx
              </button>
            </div>

            {/* Legend */}
            <div style={{padding:"8px 22px", background:"#f8fafc", borderBottom:"1px solid #f1f5f9", display:"flex", alignItems:"center", gap:16, flexWrap:"wrap"}}>
              <span style={{fontSize:11, color:"#64748b"}}>
                Exibindo <strong>{filtered.length}</strong> de <strong>{result.length}</strong> pedidos
              </span>
              <span style={{fontSize:11, fontWeight:700, color:"#64748b", textTransform:"uppercase", letterSpacing:0.5}}>Prazo:</span>
              {[
                {icon:"🟢", label:"3+ dias",         bg:"#dcfce7", text:"#14532d", border:"#22c55e"},
                {icon:"🟡", label:"2 dias",           bg:"#fef9c3", text:"#78350f", border:"#f59e0b"},
                {icon:"🔴", label:"1 dia / vencido",  bg:"#fde8e8", text:"#991b1b", border:"#ef4444"},
              ].map(l => (
                <span key={l.label} style={{background:l.bg, color:l.text, border:`1px solid ${l.border}`, borderRadius:8, padding:"2px 9px", fontSize:11, fontWeight:600}}>
                  {l.icon} {l.label}
                </span>
              ))}
            </div>

            {/* Table */}
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%", borderCollapse:"collapse", fontSize:12}}>
                <thead>
                  <tr style={{background:"#f8fafc"}}>
                    <th style={TH}>Status</th>
                    {DISPLAY_COLS.map(c => <th key={c.key} style={TH}>{c.label}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((row, i) => {
                    const s = STATUS_CFG[row._status] || STATUS_CFG.existing;
                    const isRemoving = row._status === "removed";
                    return (
                      <tr key={i} style={{
                        background: isRemoving ? "#fff5f5" : i%2===0 ? "#fff" : "#fafafa",
                        borderBottom:"1px solid #f1f5f9",
                        opacity: isRemoving ? 0.6 : 1,
                        animation: isRemoving ? "pulse 0.8s ease infinite" : "none",
                        transition:"all 0.3s",
                      }}>
                        <td style={{...TD, textAlign:"center"}}>
                          <span style={{background:s.bg, color:s.text, borderRadius:20, padding:"2px 10px", fontWeight:600, fontSize:11, whiteSpace:"nowrap"}}>
                            {s.label}
                          </span>
                        </td>
                        {DISPLAY_COLS.map(c => {
                          const isPrazo = c.key === "prazoEnvio";
                          const dl = isPrazo ? deadlineInfo(row[c.key]) : null;
                          return (
                            <td key={c.key} style={{
                              ...TD,
                              padding: isPrazo ? "6px 10px" : TD.padding,
                              fontWeight: c.key === "numeroPedido" ? 700 : 400,
                              color: c.key === "numeroPedido" ? "#1d4ed8" : "#374151",
                            }}>
                              {isPrazo && dl ? (
                                <div style={{background:dl.bg, border:`1px solid ${dl.border}`, borderRadius:8, padding:"5px 10px", display:"inline-flex", flexDirection:"column", gap:2, minWidth:130}}>
                                  <span style={{fontSize:11, color:dl.text, fontWeight:700}}>{dl.icon} {dl.label}</span>
                                  <span style={{fontSize:10, color:dl.text, opacity:0.75}}>{row[c.key].slice(0,10)}</span>
                                </div>
                              ) : isPrazo ? (
                                <span style={{color:"#94a3b8", fontSize:11}}>Sem prazo</span>
                              ) : (
                                <div style={{overflow:"hidden", textOverflow:"ellipsis", whiteSpace: c.key==="produto"?"normal":"nowrap", maxWidth: c.key==="produto"?220:160}}>
                                  {row[c.key] || "—"}
                                </div>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                  {filtered.length === 0 && (
                    <tr>
                      <td colSpan={DISPLAY_COLS.length+1} style={{textAlign:"center", padding:40, color:"#94a3b8", fontSize:14}}>
                        Nenhum pedido encontrado com os filtros selecionados
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── Instructions when no data ── */}
        {!result && (
          <div style={{background:"#fff", borderRadius:20, padding:26, boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
            <h3 style={{margin:"0 0 14px", fontSize:14, fontWeight:700, color:"#1e293b"}}>📖 Fluxo completo em 3 etapas</h3>
            <div style={{display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14}}>
              {[
                {icon:"📂", title:"1. Unificar", desc:"Carregue a planilha base e a nova. Pedidos com ID já existente são ignorados; os novos são adicionados."},
                {icon:"🗑️", title:"2. Planilha Final", desc:'Carregue a planilha final. Pedidos com "Retirada" são identificados e removidos da lista ativa.'},
                {icon:"⬇️", title:"3. Exportar", desc:"Visualize os pedidos com alertas de prazo coloridos e exporte a planilha final limpa."},
              ].map(item => (
                <div key={item.title} style={{background:"#f8fafc", borderRadius:12, padding:"18px 16px"}}>
                  <div style={{fontSize:26, marginBottom:8}}>{item.icon}</div>
                  <div style={{fontWeight:700, fontSize:13, color:"#1e293b", marginBottom:5}}>{item.title}</div>
                  <div style={{fontSize:11, color:"#64748b", lineHeight:1.6}}>{item.desc}</div>
                </div>
              ))}
            </div>
            <div style={{marginTop:16, background:"#f0fdf4", borderRadius:10, padding:"12px 16px", fontSize:12, color:"#166534"}}>
              📅 <strong>Cores do Prazo de Envio:</strong> &nbsp;
              🟢 3+ dias = verde &nbsp;|&nbsp; 🟡 2 dias = amarelo &nbsp;|&nbsp; 🔴 1 dia ou vencido = vermelho
            </div>
          </div>
        )}
      </div>

      {/* Toast */}
      {toast && <Toast msg={toast.msg} color={toast.color} />}
    </div>
  );
}

const TH = {
  padding:"10px 14px", textAlign:"left", fontWeight:700, fontSize:11,
  color:"#64748b", letterSpacing:0.5, textTransform:"uppercase",
  whiteSpace:"nowrap", borderBottom:"2px solid #e2e8f0",
};
const TD = { padding:"9px 14px", verticalAlign:"middle" };
