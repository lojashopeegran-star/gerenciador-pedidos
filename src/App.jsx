import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend, CartesianGrid } from "recharts";
import { supabase, fetchPedidos, upsertPedidos, updatePedidoStatus, fetchDevolucoes, upsertDevolucoes, fetchFaturamento, upsertFaturamento, deleteAllPedidos, deleteAllDevolucoes, deleteAllFaturamento, fetchConfig, saveConfig, fetchMembro, fetchOrganizacao, fetchMembrosDaOrganizacao, criarFuncionario, atualizarPermissoesMembro, removerMembro, fetchProdutividade, resetPassword, removerFuncionarioCompleto, registrarAuditoria, fetchAuditoria, fetchGoogleCalendarStatus, desconectarGoogleCalendar, fetchAlarmeHorarios, salvarAlarmeHorarios } from "./supabase.js";

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
  "Nome de usuário (comprador)": "nomeUsuario",
  "Data de criação do pedido":   "dataCriacao",
  "Cancelar Motivo":             "motivoCancelamento",
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

// Calcula o mês de referência (formato "2026-06") a partir da Data de criação do pedido
function mesReferencia(dataCriacao) {
  if (!dataCriacao) return "";
  // Aceita formatos "2026-06-01 00:29" ou "2026-06-01"
  const m = String(dataCriacao).match(/^(\d{4})-(\d{2})/);
  return m ? `${m[1]}-${m[2]}` : "";
}

function mesLabel(mesRef) {
  if (!mesRef) return "—";
  const [ano, mes] = mesRef.split("-");
  const nomes = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  return `${nomes[parseInt(mes,10)-1]}/${ano}`;
}

// Cores fortes e bem visíveis de destaque para cliente recorrente, por loja.
// lojas[0] = laranja forte, lojas[1] = rosa forte (pink vibrante).
function corRecorrencia(loja, lojas) {
  if (lojas && loja === lojas[0]) return { bg:"#fb923c", bgLight:"#fed7aa", color:"#7c2d12", border:"#ea580c" };
  if (lojas && loja === lojas[1]) return { bg:"#ec4899", bgLight:"#fbcfe8", color:"#831843", border:"#db2777" };
  return { bg:"#2dd4bf", bgLight:"#99f6e4", color:"#134e4a", border:"#0d9488" };
}

// Simplifica e categoriza o texto do motivo de cancelamento da Shopee
function formatMotivoCancelamento(motivo) {
  if (!motivo) return null;
  const isComprador = /cancelado pelo comprador/i.test(motivo);
  const isSistema    = /cancelado automaticamente/i.test(motivo);
  // Extrai só a parte depois de "Motivo :"
  const m = motivo.match(/Motivo\s*:\s*(.+)$/i);
  const texto = m ? m[1].trim() : motivo;
  return {
    texto,
    origem: isComprador ? "Comprador" : isSistema ? "Sistema Shopee" : "Outro",
    cor: isComprador ? { bg:"#fef3c7", color:"#92400e" } : isSistema ? { bg:"#fee2e2", color:"#991b1b" } : { bg:"#f1f5f9", color:"#64748b" },
  };
}

function isEnviado(r) {
  const s = normStatus(r.statusPedido);
  return s === "enviado" || s === "entregue" || s === "concluido" ||
         s.startsWith("o comprador pode pedir") || s === "completed" || s === "delivered" ||
         s === "order received";
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
  return s.includes("a enviar") || s === "a" || s === "pendente";
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

// Botão "Feito" que usa a cor do funcionário que marcou (sistema de times)
function FeitoButton({isFeito,cor,nome,onClick,small}) {
  const bg = isFeito ? (cor || "#059669") : "#f1f5f9";
  const fg = isFeito ? "#fff" : "#374151";
  return (
    <button
      onClick={onClick}
      title={isFeito && nome ? `Marcado por ${nome}` : ""}
      style={{
        padding: small ? "4px 7px" : "4px 9px",
        border:"none", borderRadius:7,
        background:bg, color:fg,
        fontSize:10, fontWeight:600, cursor:"pointer",
      }}>
      {isFeito ? (small ? "✅" : "✅ Feito") : (small ? "⬜" : "⬜ Feito")}
    </button>
  );
}

// Botão "Revisão" que também usa a cor do funcionário que enviou para revisão
function RevisaoButton({isRevisao,cor,nome,onClick,small}) {
  const bg = isRevisao ? (cor || "#f59e0b") : "#f1f5f9";
  const fg = isRevisao ? "#fff" : "#374151";
  return (
    <button
      onClick={onClick}
      title={isRevisao && nome ? `Enviado para revisão por ${nome}` : ""}
      style={{
        padding: small ? "4px 7px" : "4px 9px",
        border:"none", borderRadius:7,
        background:bg, color:fg,
        fontSize:10, fontWeight:600, cursor:"pointer",
      }}>
      {small ? "📋" : "📋 Revisão"}
    </button>
  );
}

// Botão "Vazio" — indica pedido com embalagem vazia/produto indisponível
function VazioButton({isVazio,cor,nome,onClick,small}) {
  const bg = isVazio ? (cor || "#6366f1") : "#f1f5f9";
  const fg = isVazio ? "#fff" : "#374151";
  return (
    <button
      onClick={onClick}
      title={isVazio && nome ? `Marcado como Vazio por ${nome}` : ""}
      style={{
        padding: small ? "4px 7px" : "4px 9px",
        border:"none", borderRadius:7,
        background:bg, color:fg,
        fontSize:10, fontWeight:600, cursor:"pointer",
      }}>
      {isVazio ? (small ? "📦" : "📦 Vazio") : (small ? "📦" : "📦 Vazio")}
    </button>
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

function RevisaoModal({order,onConfirm,onClose,onDesmarcar}) {
  const [nota,setNota]=useState(order.notaRevisao||"");
  const jaEmRevisao = order.statusInterno === "revisao";
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
        <h3 style={{margin:"0 0 6px",fontSize:15,fontWeight:700,color:"#1e293b"}}>📋 Revisão</h3>
        <p style={{margin:"0 0 14px",fontSize:12,color:"#64748b"}}>Pedido: <strong>{order.idPedido}</strong>{order.feitoPorNome&&<span> · enviado por {order.feitoPorNome}</span>}</p>
        <textarea value={nota} onChange={e=>setNota(e.target.value)} placeholder="Descreva o motivo..." autoFocus
          style={{width:"100%",minHeight:90,padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,resize:"vertical",outline:"none",boxSizing:"border-box",fontFamily:"inherit"}} />
        <div style={{display:"flex",gap:10,marginTop:14,justifyContent:jaEmRevisao?"space-between":"flex-end"}}>
          {jaEmRevisao&&(
            <button onClick={()=>onDesmarcar(order)} style={{padding:"8px 14px",border:"1px solid #fecaca",borderRadius:10,background:"#fff",color:"#ef4444",fontSize:12,fontWeight:600,cursor:"pointer"}}>Desmarcar revisão</button>
          )}
          <div style={{display:"flex",gap:10}}>
            <button onClick={onClose} style={{padding:"8px 18px",border:"1px solid #e2e8f0",borderRadius:10,background:"#fff",fontSize:12,cursor:"pointer"}}>Cancelar</button>
            <button onClick={()=>onConfirm(nota)} disabled={!nota.trim()} style={{padding:"8px 18px",border:"none",borderRadius:10,background:nota.trim()?"#f59e0b":"#e2e8f0",color:nota.trim()?"#fff":"#94a3b8",fontSize:12,fontWeight:700,cursor:nota.trim()?"pointer":"not-allowed"}}>{jaEmRevisao?"Salvar":"Confirmar"}</button>
          </div>
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

const TH={padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:10,color:"#64748b",letterSpacing:0.5,textTransform:"uppercase",whiteSpace:"nowrap",borderBottom:"2px solid #e2e8f0"};
const TD={padding:"7px 11px",verticalAlign:"middle"};

// ── MAIN APP ──────────────────────────────────────────────────────────────────
export default function App({user,onLogout}) {
  const [allPedidos,    setAllPedidos]    = useState([]);
  const [devolucoes,    setDevolucoes]    = useState([]);
  const [faturamento,   setFaturamento]   = useState([]);
  const [dbLoading,     setDbLoading]     = useState(true);
  const [lojas,           setLojas]           = useState(["Loja 1","Loja 2"]);
  const [showConfigLojas, setShowConfigLojas] = useState(false);
  const [configLoja1,     setConfigLoja1]     = useState("");
  const [configLoja2,     setConfigLoja2]     = useState("");
  const [saving,          setSaving]          = useState(false);
  const [uploadProgress,  setUploadProgress]  = useState(null); // null = idle, 0-100 = %
  const [toast,         setToast]         = useState(null);
  const [activeTab,     setActiveTab]     = useState("abertos");
  const [financeUnlocked,setFinanceUnlocked]=useState(false);
  const [showFinanceGate,setShowFinanceGate]=useState(false);
  const [revisaoModal,      setRevisaoModal]      = useState(null);
  const [confirmarDesmarcar, setConfirmarDesmarcar] = useState(null); // {order, novoSt, label}
  const [uploadNames,   setUploadNames]   = useState({});
  const [search,        setSearch]        = useState("");
  const [filterLoja,    setFilterLoja]    = useState("all");
  const [filterPrazo,   setFilterPrazo]   = useState("all");
  const [filterData,    setFilterData]    = useState("all"); // specific date filter
  const [searchEnviados, setSearchEnviados] = useState("");
  const [mesClientesFiltro, setMesClientesFiltro] = useState("all");
  const [filterMotivo, setFilterMotivo] = useState("all");
  const [filterDataEnviados, setFilterDataEnviados] = useState("all");
  const [filterDataCancelados, setFilterDataCancelados] = useState("all");
  const [filterMotivoDevolucao, setFilterMotivoDevolucao] = useState("all");
  const [clienteExpandido, setClienteExpandido] = useState(null); // chave "nomeUsuario__loja"
  const [filterSt,      setFilterSt]      = useState("all");
  const [filterProduto, setFilterProduto] = useState("all");
  const [showProdPanel, setShowProdPanel] = useState(false);
  const [selected,      setSelected]      = useState(new Set());
  const [copied,        setCopied]        = useState(false);
  const [copiedFull,    setCopiedFull]    = useState(false);
  // ── Sistema de times ───────────────────────────────────────────────────────
  const [membro,        setMembro]        = useState(null); // current member record (with permissions)
  const [organizacao,   setOrganizacao]   = useState(null); // current organization
  const [membrosEquipe, setMembrosEquipe] = useState([]);   // all members in org (for admin)
  const [produtividade, setProdutividade] = useState([]);   // raw productivity data
  const [auditoria,     setAuditoria]     = useState([]);   // audit log entries
  const [googleCalStatus, setGoogleCalStatus] = useState(null); // { id, ativo, criado_em } | null
  const [conectandoGoogle, setConectandoGoogle] = useState(false);
  const [alarmeHorarios, setAlarmeHorarios] = useState(["09:30","11:00","13:00","14:30"]);
  const [salvandoHorarios, setSalvandoHorarios] = useState(false);
  const [showEquipeTab, setShowEquipeTab] = useState(false);
  const [showNovoFuncionario, setShowNovoFuncionario] = useState(false);
  const [novoFuncNome,     setNovoFuncNome]     = useState("");
  const [novoFuncEmail,    setNovoFuncEmail]    = useState("");
  const [novoFuncSenha,    setNovoFuncSenha]    = useState("");
  const [novoFuncCor,      setNovoFuncCor]      = useState("#ef4444");
  const [novoFuncPerms,    setNovoFuncPerms]    = useState({
    pode_ver_financeiro: false, pode_zerar_sistema: false,
    pode_carregar_planilha: true, pode_editar_status: true,
  });
  const [criandoFuncionario, setCriandoFuncionario] = useState(false);
  const tableRef  = useRef(null);
  const notifRef  = useRef(null);

  // ── Notificações in-app (declarados ANTES do useEffect que os usa) ──────────
  const [showNotifPanel,  setShowNotifPanel]  = useState(false);
  const [lidas,           setLidas]           = useState(() => {
    try { return new Set(JSON.parse(localStorage.getItem("notif_lidas")||"[]")); } catch { return new Set(); }
  });

  const showToast=(msg,color,ms=3500)=>{setToast({msg,color});setTimeout(()=>setToast(null),ms);};

  // Fechar painel de notificações ao clicar fora
  useEffect(() => {
    if (!showNotifPanel) return;
    const handler = (e) => {
      if (notifRef.current && !notifRef.current.contains(e.target)) {
        setShowNotifPanel(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [showNotifPanel]);


  // ── Load all data from Supabase on mount ───────────────────────────────────
  useEffect(()=>{
    if(!user?.id){setDbLoading(false);return;}
    if(!supabase){setDbLoading(false);return;}
    setDbLoading(true);

    // First, find out which organization this user belongs to (if any)
    fetchMembro(user.id).then(async (m) => {
      setMembro(m);
      const orgId = m?.organizacao_id || null;

      if (orgId) {
        const org = await fetchOrganizacao(orgId);
        setOrganizacao(org);
        // Todos os membros precisam da lista da equipe para resolver as cores
        // de quem marcou cada pedido como feito — não só o admin.
        fetchMembrosDaOrganizacao(orgId).then(setMembrosEquipe);
        if (m?.is_admin) {
          fetchProdutividade(orgId).then(setProdutividade);
          fetchAuditoria(orgId, true).then(setAuditoria);
          fetchGoogleCalendarStatus(orgId).then(setGoogleCalStatus);
          fetchAlarmeHorarios(orgId).then(setAlarmeHorarios);
        }
      }

      Promise.all([
        fetchPedidos(user.id, orgId),
        fetchDevolucoes(user.id, orgId),
        fetchFaturamento(user.id, orgId),
        fetchConfig(user.id, orgId),
      ]).then(([p,d,f,cfg])=>{
          setAllPedidos(p); setDevolucoes(d); setFaturamento(f);
          if (cfg?.loja1 && cfg?.loja2) {
            setLojas([cfg.loja1, cfg.loja2]);
            setConfigLoja1(cfg.loja1);
            setConfigLoja2(cfg.loja2);
          } else if (!orgId || m?.is_admin) {
            // First time — show config modal (only for admins/standalone users)
            setShowConfigLojas(true);
          }
          setDbLoading(false);
        }).catch(err=>{console.error(err);setDbLoading(false);});
    }).catch(err=>{console.error(err);setDbLoading(false);});
  },[user?.id]);

  // ── Permission helpers (default to true/permissive when no team system active) ──
  const podeVerFinanceiro    = membro ? membro.pode_ver_financeiro    : true;
  const podeZerarSistema     = membro ? membro.pode_zerar_sistema     : true;
  const podeCarregarPlanilha = membro ? membro.pode_carregar_planilha : true;
  const podeEditarStatus     = membro ? membro.pode_editar_status     : true;
  const isAdminEquipe        = membro ? membro.is_admin : true;
  const orgId                = membro?.organizacao_id || null;
  const minhaCor              = membro?.cor || "#059669";
  const meuNome                = membro?.nome || user?.email?.split("@")[0] || "Você";

  // ── Mapa nome → cor (sistema de times) ──────────────────────────────────────
  const corPorNome = membrosEquipe.reduce((acc,m)=>{ acc[m.nome]=m.cor; return acc; },{});

  // ── Total histórico de compras por (nomeUsuario+loja) — usado para destacar
  // clientes recorrentes em QUALQUER aba (Em Aberto, Enviados, etc) ────────────
  const totalHistoricoCliente = {};
  for (const p of allPedidos) {
    if (!p.nomeUsuario) continue;
    const k = `${p.nomeUsuario}__${p.loja}`;
    totalHistoricoCliente[k] = (totalHistoricoCliente[k] || 0) + 1;
  }

  const pedidosComCor = allPedidos.map(p => {
    const corFeito = p.feitoPorNome && corPorNome[p.feitoPorNome] ? { feitoPorCor: corPorNome[p.feitoPorNome] } : {};
    const chaveCliente = p.nomeUsuario ? `${p.nomeUsuario}__${p.loja}` : null;
    const totalCliente = chaveCliente ? (totalHistoricoCliente[chaveCliente] || 0) : 0;
    return { ...p, ...corFeito, clienteRecorrente: totalCliente > 1, totalComprasCliente: totalCliente };
  });

  // ── Pedido groups ──────────────────────────────────────────────────────────
  const pedidosAbertos    = pedidosComCor.filter(isAberto);
  const pedidosEnviados   = pedidosComCor.filter(isEnviado);
  const pedidosCancelados = pedidosComCor.filter(isCancelado);

  // ── Load planilha ──────────────────────────────────────────────────────────
  const handlePlanilha = async (sheet, name, loja) => {
    setUploadNames(p=>({...p,[`ped_${loja}`]:name}));
    const incoming = parseSheet(sheet, PEDIDO_COLS, loja).map(r => ({
      ...r,
      mesReferencia: mesReferencia(r.dataCriacao),
    }));
    if (!incoming.length) { showToast("Nenhum pedido encontrado na planilha.","#d97706"); return; }

    // Merge: update existing status, add new ones
    // Same ID = update fields from spreadsheet but PRESERVE statusInterno and notaRevisao
    const map = new Map(allPedidos.map(r=>[r.idPedido,r]));
    let added=0, updated=0;
    for (const r of incoming) {
      const ex = map.get(r.idPedido);
      if (ex) {
        map.set(r.idPedido, {
          ...ex,
          statusPedido:  r.statusPedido,
          dataEnvio:     r.dataEnvio,
          horaPagamento: r.horaPagamento,
          produto:       r.produto,
          preco:         r.preco,
          quantidade:    r.quantidade,
          variacao:      r.variacao,
          destinatario:  r.destinatario,
          notas:         r.notas,
          loja:          r.loja,
          nomeUsuario:   r.nomeUsuario,
          dataCriacao:   r.dataCriacao,
          mesReferencia: r.mesReferencia,
          motivoCancelamento: r.motivoCancelamento,
          // PRESERVE internal flags — never reset from spreadsheet load
          statusInterno: ex.statusInterno || "",
          notaRevisao:   ex.notaRevisao   || "",
          _status: "existing",
        });
        updated++;
      } else {
        map.set(r.idPedido, {...r, statusInterno:"", notaRevisao:""});
        added++;
      }
    }
    const merged = Array.from(map.values());
    setAllPedidos(merged);
    showToast(`✅ ${added} novos · ${updated} atualizados`,"#059669");

    // Save to DB
    if (supabase) {
      setSaving(true);
      setUploadProgress(0);
      await upsertPedidos(incoming, user.id, orgId, (pct) => setUploadProgress(pct));
      setUploadProgress(100);
      // Faturamento: only count "a enviar" orders
      const aEnviar = incoming.filter(isAberto);
      const totalValor = aEnviar.reduce((s,r)=>s+(parseFloat(r.preco)||0),0);
      if (totalValor > 0) await upsertFaturamento(user.id, currentMes(), loja, totalValor, orgId);
      const fat = await fetchFaturamento(user.id);
      setFaturamento(fat);
      setSaving(false);
      setTimeout(() => setUploadProgress(null), 800);
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
    if (supabase) { setSaving(true); await upsertDevolucoes(rows,user.id,orgId); setSaving(false); }
  };

  // ── Clear all data ────────────────────────────────────────────────────────
  const [showConfirmClear, setShowConfirmClear] = useState(false);


  const handleClearAll = async () => {
    setShowConfirmClear(false);
    setSaving(true);
    const totalAntes = allPedidos.length;
    if (supabase) {
      await Promise.all([
        deleteAllPedidos(user.id, orgId),
        deleteAllDevolucoes(user.id, orgId),
        deleteAllFaturamento(user.id, orgId),
      ]);
      registrarAuditoria(orgId, user.id, meuNome, "zerou_sistema", `Apagou ${totalAntes} pedido(s), todas as devoluções e faturamento.`);
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
    if (!podeEditarStatus) { showToast("🔒 Você não tem permissão para editar status.","#991b1b"); return; }
    // Bloqueia alterar/desmarcar um pedido (Feito ou Revisão) já marcado por outra pessoa (exceto admin)
    const jaMarcadoPorOutro = order.feitoPorNome && order.feitoPorNome !== meuNome &&
      (order.statusInterno === "feito" || order.statusInterno === "revisao");
    if (jaMarcadoPorOutro && !isAdminEquipe) {
      const acao = order.statusInterno === "feito" ? "marcado como Feito" : "enviado para Revisão";
      showToast(`🔒 Este pedido foi ${acao} por ${order.feitoPorNome}. Só ${order.feitoPorNome} ou um admin pode alterar.`,"#991b1b");
      return;
    }
    setAllPedidos(prev=>prev.map(o=>o.idPedido===order.idPedido?{...o,statusInterno:newSt,notaRevisao:nota,feitoPorNome:newSt?meuNome:"",feitoPorCor:newSt?minhaCor:""}:o));
    if (supabase) await updatePedidoStatus(order.idPedido,user.id,newSt,nota,orgId,meuNome);
    showToast(newSt==="feito"?"✅ Marcado como Feito!":newSt==="vazio"?"📦 Marcado como Vazio!":newSt==="revisao"?"📋 Enviado para Revisão":"↩️ Desmarcado",newSt==="feito"?"#059669":newSt==="vazio"?"#6366f1":newSt==="revisao"?"#f59e0b":"#64748b");
  };

  const tentarAbrirRevisao = (order) => {
    if (!podeEditarStatus) { showToast("🔒 Você não tem permissão para editar status.","#991b1b"); return; }
    const jaMarcadoPorOutro = order.feitoPorNome && order.feitoPorNome !== meuNome &&
      (order.statusInterno === "feito" || order.statusInterno === "revisao");
    if (jaMarcadoPorOutro && !isAdminEquipe) {
      const acao = order.statusInterno === "feito" ? "marcado como Feito" : "enviado para Revisão";
      showToast(`🔒 Este pedido foi ${acao} por ${order.feitoPorNome}. Só ${order.feitoPorNome} ou um admin pode alterar.`,"#991b1b");
      return;
    }
    setRevisaoModal(order);
  };

  // Intercepta clique de DESMARCAR feito ou revisão — pede confirmação
  const tentarDesmarcar = (order, novoSt, label) => {
    setConfirmarDesmarcar({ order, novoSt, label });
  };

  // ── Selection ──────────────────────────────────────────────────────────────
  const toggleSelect=id=>setSelected(p=>{const s=new Set(p);s.has(id)?s.delete(id):s.add(id);return s;});

  const markAllFeito = async () => {
    if (!podeEditarStatus) { showToast("🔒 Você não tem permissão para editar status.","#991b1b"); return; }
    const ids = [...selected];
    // Update all selected as feito
    setAllPedidos(prev => prev.map(o => ids.includes(o.idPedido) ? {...o, statusInterno:"feito", notaRevisao:"", feitoPorNome:meuNome, feitoPorCor:minhaCor} : o));
    setSelected(new Set());
    showToast(`✅ ${ids.length} pedido(s) marcados como Feito!`, "#059669");
    if (supabase) {
      for (const id of ids) {
        await updatePedidoStatus(id, user.id, "feito", "", orgId, meuNome);
      }
    }
  };
  // ── Exportar dados para Excel ──────────────────────────────────────────────
  const exportToExcel = (rows, filename, columnsMap) => {
    if (!rows.length) { showToast("Nada para exportar.","#d97706"); return; }
    const data = rows.map(r => {
      const obj = {};
      for (const [label, key] of Object.entries(columnsMap)) {
        obj[label] = typeof key === "function" ? key(r) : (r[key] ?? "");
      }
      return obj;
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dados");
    XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast(`📥 ${rows.length} linha(s) exportada(s)!`,"#059669");
  };

  const copyIds=()=>{
    navigator.clipboard.writeText([...selected].join(",")).then(()=>{
      setCopied(true);
      setTimeout(()=>setCopied(false),2000);
      showToast("📋 IDs copiados para o UpSeller!","#1d4ed8");
    });
  };

  const copyFullData=()=>{
    const lines = [...selected].map(id=>{
      const r = allPedidos.find(p=>p.idPedido===id);
      if(!r) return id;
      const prazoDate = r.dataEnvio ? r.dataEnvio.slice(0,10).split("-").reverse().join("-") : "";
      return [r.idPedido, r.destinatario||"", r.loja||"", r.variacao||"", prazoDate].join("-");
    });
    navigator.clipboard.writeText(lines.join("\n")).then(()=>{
      setCopiedFull(true);
      setTimeout(()=>setCopiedFull(false),2000);
      showToast("Dados completos copiados!","#7c3aed");
    });
  };

  // ── Filters ────────────────────────────────────────────────────────────────
  const produtoMap = pedidosAbertos.reduce((a,r)=>{if(r.produto)a[r.produto]=(a[r.produto]||0)+1;return a;},{});
  const produtos   = Object.entries(produtoMap).sort((a,b)=>b[1]-a[1]);

  // Unique deadline dates sorted ascending
  const datasUnicas = [...new Set(
    pedidosAbertos
      .map(r => r.dataEnvio ? r.dataEnvio.slice(0,10) : null)
      .filter(Boolean)
  )].sort();

  // Support multi-ID search: "ID1,ID2,ID3"
  const searchTerms = search.split(",").map(s=>s.trim().toLowerCase()).filter(Boolean);
  const matchesSearch = (r) => {
    if (!search) return true;
    // If multiple terms separated by comma, match any ID exactly
    if (searchTerms.length > 1) {
      return searchTerms.some(t => r.idPedido.toLowerCase() === t);
    }
    // Single term — search across all fields
    const q = searchTerms[0];
    return r.idPedido.toLowerCase().includes(q) ||
           r.produto.toLowerCase().includes(q) ||
           r.destinatario.toLowerCase().includes(q) ||
           r.variacao.toLowerCase().includes(q);
  };

  const filtered = pedidosAbertos.filter(r=>{
    const dl=deadlineInfo(r.dataEnvio);
    return (
      (filterLoja==="all"   ||r.loja===filterLoja)&&
      (filterPrazo==="all"  ||(dl&&dl.tier===filterPrazo))&&
      (filterData==="all"   ||r.dataEnvio?.slice(0,10)===filterData)&&
      (filterSt==="all"     ||(filterSt==="meus_feitos" ? (r.statusInterno==="feito"&&r.feitoPorNome===meuNome) : r.statusInterno===filterSt))&&
      (filterProduto==="all"||r.produto===filterProduto)&&
      matchesSearch(r)
    );
  }).sort((a,b)=>{
    const day=r=>{if(!r.dataEnvio)return 9999;const d=new Date(r.dataEnvio.replace(" ","T"));if(isNaN(d))return 9999;const t=new Date();t.setHours(0,0,0,0);const dd=new Date(d);dd.setHours(0,0,0,0);return Math.round((dd-t)/86400000);};
    return day(a)-day(b);
  });

  // toggleAll usa 'filtered', por isso é declarado DEPOIS dele (evita TDZ)
  const toggleAll=()=>selected.size===filtered.length?setSelected(new Set()):setSelected(new Set(filtered.map(r=>r.idPedido)));

  const urgentAll    = pedidosAbertos.filter(r=>deadlineInfo(r.dataEnvio)?.tier==="red");
  const urgentCount  = urgentAll.length;
  const urgentFeitos = urgentAll.filter(r=>r.statusInterno==="feito").length;
  const feitoCount   = pedidosAbertos.filter(r=>r.statusInterno==="feito").length;
  const revisaoCount = pedidosAbertos.filter(r=>r.statusInterno==="revisao").length;
  const vazioCount   = pedidosAbertos.filter(r=>r.statusInterno==="vazio").length;
  const valorAberto  = pedidosAbertos.reduce((s,r)=>s+(parseFloat(r.preco)||0),0);
  const totalDev     = devolucoes.reduce((s,d)=>s+(parseFloat(d.preco_unidade||0)*parseInt(d.quantidade||1)),0);

  // ── Notificações computadas ─────────────────────────────────────────────────
  // ── Notificações computadas ────────────────────────────────────────────────
  function buildNotifs() {
    var prazos = [];
    var revisoes = [];
    var revisaoIds = new Set();
    for (var i = 0; i < pedidosAbertos.length; i++) {
      var r = pedidosAbertos[i];
      var dl = deadlineInfo(r.dataEnvio);
      if (!dl || dl.tier !== "red") continue;
      var desc = "Pedido " + r.idPedido + (r.destinatario ? " · " + r.destinatario : "");
      if (r.statusInterno === "revisao") {
        revisoes.push({ id: "revisao__" + r.idPedido, tipo: "revisao", icon: "📋", titulo: "Revisão com prazo crítico", desc: desc, cor: "#f59e0b" });
        revisaoIds.add("prazo__" + r.idPedido);
      }
      if (r.statusInterno !== "feito") {
        prazos.push({ id: "prazo__" + r.idPedido, tipo: "prazo", icon: "🔴", titulo: dl.label === "Vencido!" ? "Prazo vencido!" : "Prazo: amanhã", desc: desc, cor: "#ef4444" });
      }
    }
    var prazosFiltrados = prazos.filter(function(n) { return !revisaoIds.has(n.id); });
    return revisoes.concat(prazosFiltrados);
  }
  var todasNotifs = buildNotifs();
  var notifNaoLidas = todasNotifs.filter(function(n) { return !lidas.has(n.id); });

  function marcarTodasLidas() {
    var ids = todasNotifs.map(function(n) { return n.id; });
    var novas = new Set([...lidas, ...ids]);
    setLidas(novas);
    try { localStorage.setItem("notif_lidas", JSON.stringify(Array.from(novas))); } catch(e) {}
  }
  function marcarLida(id) {
    var novas = new Set([...lidas, id]);
    setLidas(novas);
    try { localStorage.setItem("notif_lidas", JSON.stringify(Array.from(novas))); } catch(e) {}
  }
  function notifAcao(n) {
    if (n.tipo === "revisao") { setActiveTab("abertos"); setFilterSt("revisao"); setFilterPrazo("red"); }
    else { setActiveTab("abertos"); setFilterPrazo("red"); setFilterSt("all"); }
    setShowNotifPanel(false);
  }

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
    {id:"clientes",   icon:"👤",label:"Clientes",   badge:null,                     color:"#0d9488"},
    {id:"financeiro", icon:"💰",label:"Financeiro", badge:null,                     color:"#0891b2"},
    ...(isAdminEquipe&&organizacao ? [{id:"equipe", icon:"👥", label:"Equipe", badge:membrosEquipe.length||null, color:"#7c3aed"}] : []),
  ];


  const handleSaveConfig = async () => {
    const l1 = configLoja1.trim() || "Loja 1";
    const l2 = configLoja2.trim() || "Loja 2";
    setLojas([l1, l2]);
    setShowConfigLojas(false);
    if (supabase) await saveConfig(user.id, l1, l2, orgId);
    showToast("✅ Lojas configuradas com sucesso!", "#059669");
  };

  // ── Google Calendar — conectar / desconectar ─────────────────────────────────
  const GOOGLE_CLIENT_ID = "786408703311-jic96ggrroh26usc5jb4n0pemnm2tmqd.apps.googleusercontent.com";
  const GOOGLE_REDIRECT_URI = typeof window !== "undefined" ? window.location.origin : "";

  const conectarGoogleCalendar = () => {
    const params = new URLSearchParams({
      client_id: GOOGLE_CLIENT_ID,
      redirect_uri: GOOGLE_REDIRECT_URI,
      response_type: "code",
      scope: "https://www.googleapis.com/auth/calendar.events",
      access_type: "offline",
      prompt: "consent",
      state: "google_calendar_connect",
    });
    window.location.href = `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
  };

  // Detecta retorno do Google (code na URL) e finaliza a conexão
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const code = params.get("code");
    const state = params.get("state");
    if (code && state === "google_calendar_connect" && supabase && orgId) {
      setConectandoGoogle(true);
      supabase.functions.invoke("google-calendar-connect", {
        body: { code, redirectUri: GOOGLE_REDIRECT_URI },
      }).then(({ data, error }) => {
        setConectandoGoogle(false);
        // Limpa o "code" da URL para não tentar reconectar de novo
        window.history.replaceState({}, "", window.location.pathname);
        if (error || data?.error) {
          showToast(`❌ ${data?.error || "Erro ao conectar Google Calendar."}`, "#991b1b");
        } else {
          showToast("✅ Google Calendar conectado com sucesso!", "#059669");
          fetchGoogleCalendarStatus(orgId).then(setGoogleCalStatus);
          registrarAuditoria(orgId, user.id, meuNome, "conectou_google_calendar", "Conectou o Google Calendar para alarmes de prazo.");
        }
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [orgId]);

  const handleDesconectarGoogle = async () => {
    if (!window.confirm("Desconectar o Google Calendar? Os alarmes automáticos de prazo vão parar de ser criados.")) return;
    await desconectarGoogleCalendar(orgId);
    setGoogleCalStatus(null);
    registrarAuditoria(orgId, user.id, meuNome, "desconectou_google_calendar", "Desconectou o Google Calendar.");
    showToast("🔌 Google Calendar desconectado.", "#64748b");
  };

  const handleCriarFuncionario = async () => {
    if (!novoFuncNome.trim() || !novoFuncEmail.trim() || !novoFuncSenha.trim()) {
      showToast("Preencha nome, e-mail e senha.","#991b1b"); return;
    }
    if (novoFuncSenha.length < 6) { showToast("A senha precisa ter pelo menos 6 caracteres.","#991b1b"); return; }
    setCriandoFuncionario(true);
    const res = await criarFuncionario(orgId, {
      nome: novoFuncNome.trim(), email: novoFuncEmail.trim(), password: novoFuncSenha,
      cor: novoFuncCor, permissoes: novoFuncPerms,
    });
    setCriandoFuncionario(false);
    if (res.error) { showToast(`❌ ${res.error}`,"#991b1b"); return; }
    registrarAuditoria(orgId, user.id, meuNome, "criou_funcionario", `Criou o login de "${novoFuncNome.trim()}" (${novoFuncEmail.trim()}).`);
    showToast(`✅ Funcionário ${novoFuncNome} criado com sucesso!`,"#059669");
    setShowNovoFuncionario(false);
    setNovoFuncNome(""); setNovoFuncEmail(""); setNovoFuncSenha("");
    setNovoFuncCor("#ef4444");
    setNovoFuncPerms({ pode_ver_financeiro:false, pode_zerar_sistema:false, pode_carregar_planilha:true, pode_editar_status:true });
    if (orgId) fetchMembrosDaOrganizacao(orgId).then(setMembrosEquipe);
  };

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
          {uploadProgress !== null ? (
            <span style={{background:"rgba(255,255,255,0.15)",borderRadius:20,padding:"4px 12px",fontSize:12,display:"flex",alignItems:"center",gap:8,minWidth:180}}>
              <span style={{fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>⬆️ Salvando {uploadProgress}%</span>
              <span style={{flex:1,height:6,background:"rgba(255,255,255,0.2)",borderRadius:99,overflow:"hidden",minWidth:80}}>
                <span style={{display:"block",height:"100%",width:`${uploadProgress}%`,background:"#4ade80",borderRadius:99,transition:"width 0.3s ease"}}/>
              </span>
            </span>
          ) : (
            <span style={{background:"rgba(255,255,255,0.15)",borderRadius:20,padding:"3px 12px",fontSize:12,display:"flex",alignItems:"center",gap:5}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:saving?"#fbbf24":"#4ade80",display:"inline-block"}}/>{saving?"Salvando...":"Conectado"}
            </span>
          )}
          <span style={{background:"rgba(255,255,255,0.12)",borderRadius:20,padding:"3px 12px",fontSize:12}}>👤 {user?.email}</span>
          <button onClick={()=>{setConfigLoja1(lojas[0]);setConfigLoja2(lojas[1]);setShowConfigLojas(true);}} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"3px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>⚙️ Lojas</button>
          {/* ── Sino de Notificações ── */}
          <div ref={notifRef} style={{position:"relative"}}>
            <button onClick={()=>setShowNotifPanel(v=>!v)} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"3px 14px",fontSize:12,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
              🔔
              {notifNaoLidas.length>0&&(
                <span style={{background:"#ef4444",color:"#fff",borderRadius:"50%",width:18,height:18,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,lineHeight:1}}>
                  {notifNaoLidas.length>9?"9+":notifNaoLidas.length}
                </span>
              )}
            </button>
            {showNotifPanel&&(
              <div style={{position:"absolute",right:0,top:"calc(100% + 10px)",width:340,background:"#fff",borderRadius:16,boxShadow:"0 8px 32px rgba(0,0,0,0.18)",zIndex:9000,overflow:"hidden",border:"1px solid #e2e8f0"}}>
                <div style={{padding:"14px 16px 10px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#1e293b"}}>🔔 Notificações</div>
                  <div style={{display:"flex",gap:8,alignItems:"center"}}>
                    {notifNaoLidas.length>0&&(
                      <button onClick={marcarTodasLidas} style={{background:"none",border:"none",fontSize:11,color:"#3b82f6",cursor:"pointer",fontWeight:600,padding:"2px 6px"}}>Marcar todas como lidas</button>
                    )}
                    <button onClick={()=>setShowNotifPanel(false)} style={{background:"#f1f5f9",border:"none",borderRadius:8,width:24,height:24,cursor:"pointer",fontSize:14,color:"#64748b",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
                  </div>
                </div>
                <div style={{maxHeight:380,overflowY:"auto"}}>
                  {todasNotifs.length===0?(
                    <div style={{padding:"32px 16px",textAlign:"center",color:"#94a3b8"}}>
                      <div style={{fontSize:32,marginBottom:8}}>✅</div>
                      <div style={{fontSize:13,fontWeight:500}}>Tudo em dia!</div>
                      <div style={{fontSize:12,marginTop:4}}>Nenhum prazo urgente no momento.</div>
                    </div>
                  ):(
                    todasNotifs.map(n=>{
                      const isLida = lidas.has(n.id);
                      return (
                        <div key={n.id} style={{padding:"12px 16px",borderBottom:"1px solid #f8fafc",background:isLida?"#fff":"#fafbff",display:"flex",gap:12,alignItems:"flex-start",cursor:"pointer",transition:"background 0.15s"}}
                          onMouseEnter={e=>e.currentTarget.style.background="#f8fafc"}
                          onMouseLeave={e=>e.currentTarget.style.background=isLida?"#fff":"#fafbff"}
                          onClick={()=>{marcarLida(n.id);notifAcao(n);}}>
                          <div style={{width:36,height:36,borderRadius:10,background:n.cor+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{n.icon}</div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:isLida?500:700,color:isLida?"#64748b":"#1e293b",marginBottom:2}}>{n.titulo}</div>
                            <div style={{fontSize:11,color:"#94a3b8",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{n.desc}</div>
                          </div>
                          {!isLida&&<div style={{width:8,height:8,borderRadius:"50%",background:n.cor,flexShrink:0,marginTop:4}}/>}
                        </div>
                      );
                    })
                  )}
                </div>
              </div>
            )}
          </div>
          <button onClick={onLogout} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"3px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Sair →</button>
        </div>
      </div>

      <div style={{maxWidth:1600,margin:"0 auto",padding:"20px 28px"}}>

        {/* Upload */}
        <div style={{background:"#fff",borderRadius:18,padding:18,marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
          <div style={{display:"flex",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:0.5}}>Carregar Planilhas</div>
            {podeZerarSistema&&<button onClick={()=>setShowConfirmClear(true)} style={{marginLeft:"auto",background:"#fde8e8",color:"#991b1b",border:"1px solid #fca5a5",borderRadius:10,padding:"6px 14px",fontSize:11,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:5}}>
              🗑️ Zerar Sistema
            </button>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
            {lojas.map(loja=>(
              <div key={`ped_${loja}`} style={{border:"1px solid #e2e8f0",borderRadius:12,padding:10}}>
                <div style={{fontSize:11,fontWeight:700,color:"#1d4ed8",marginBottom:7,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{background:"#eff6ff",borderRadius:6,padding:"1px 8px"}}>{loja}</span>
                  <span style={{color:"#64748b",fontWeight:400}}>· Pedidos</span>
                </div>
                <DropZone label="Carregar Planilha" sublabel="A Enviar / Enviado / Entregue / Cancelado" color="#1d4ed8"
                  file={uploadNames[`ped_${loja}`]} onFile={(s,n)=>handlePlanilha(s,n,loja)} disabled={!podeCarregarPlanilha} />
              </div>
            ))}
            {lojas.map(loja=>(
              <div key={`dev_${loja}`} style={{border:"1px solid #fee2e2",borderRadius:12,padding:10}}>
                <div style={{fontSize:11,fontWeight:700,color:"#ef4444",marginBottom:7,display:"flex",alignItems:"center",gap:5}}>
                  <span style={{background:"#fef2f2",borderRadius:6,padding:"1px 8px"}}>{loja}</span>
                  <span style={{color:"#64748b",fontWeight:400}}>· Devoluções</span>
                </div>
                <DropZone label="Planilha de Devolução" sublabel="Devolução e Reembolso" color="#ef4444"
                  file={uploadNames[`dev_${loja}`]} onFile={(s,n)=>handleDevolucao(s,n,loja)} disabled={!podeCarregarPlanilha} />
              </div>
            ))}
          </div>
        </div>

        {/* ── Alerta de prazos vencendo hoje ── */}
        {(() => {
          const vencendoHoje = pedidosAbertos.filter(r => deadlineInfo(r.dataEnvio)?.tier === "red" && r.statusInterno !== "feito");
          if (vencendoHoje.length === 0) return null;
          return (
            <div style={{
              background: "linear-gradient(135deg,#dc2626,#ef4444)", borderRadius: 14, padding: "14px 20px",
              marginBottom: 18, display: "flex", alignItems: "center", gap: 14, boxShadow: "0 4px 14px rgba(239,68,68,0.3)",
              animation: "pulseAlert 2s ease-in-out infinite",
            }}>
              <style>{`@keyframes pulseAlert{0%,100%{box-shadow:0 4px 14px rgba(239,68,68,0.3)}50%{box-shadow:0 4px 22px rgba(239,68,68,0.55)}}`}</style>
              <span style={{ fontSize: 26 }}>🚨</span>
              <div style={{ flex: 1 }}>
                <div style={{ color: "#fff", fontWeight: 800, fontSize: 14 }}>
                  {vencendoHoje.length} pedido{vencendoHoje.length !== 1 ? "s" : ""} vencendo hoje ou já vencido{vencendoHoje.length !== 1 ? "s" : ""}!
                </div>
                <div style={{ color: "rgba(255,255,255,0.85)", fontSize: 12 }}>Priorize esses pedidos para não atrasar o envio.</div>
              </div>
              <button onClick={() => { setActiveTab("abertos"); setFilterSt("all"); setFilterPrazo("red"); }}
                style={{ background: "#fff", color: "#dc2626", border: "none", borderRadius: 10, padding: "8px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer", flexShrink: 0 }}>
                Ver agora →
              </button>
            </div>
          );
        })()}

        {/* Stats */}
        <div style={{display:"grid",gridTemplateColumns:organizacao?"repeat(8,1fr)":"repeat(7,1fr)",gap:10,marginBottom:18}}>
          {[
            {label:"Em Aberto",  value:pedidosAbertos.length,    color:"#1d4ed8",icon:"📋",tab:"abertos",  action:()=>{setFilterSt("all");setFilterPrazo("all");}},
            {label:"Feitos",     value:feitoCount,                color:"#059669",icon:"✅",tab:"abertos",  action:()=>setFilterSt("feito")},
            ...(organizacao ? [{label:"Meus Feitos", value:pedidosAbertos.filter(r=>r.feitoPorNome===meuNome).length, color:minhaCor,icon:"⭐",tab:"abertos",action:()=>setFilterSt("meus_feitos")}] : []),
            {label:"Em Revisão", value:revisaoCount,              color:"#f59e0b",icon:"📋",tab:"abertos",  action:()=>setFilterSt("revisao")},
            {label:"Vazios",      value:vazioCount,                color:"#6366f1",icon:"📦",tab:"abertos",  action:()=>setFilterSt("vazio")},
            {label:"Urgentes",   value:`${urgentFeitos}/${urgentCount}`, color:"#ef4444",icon:"🔴",tab:"abertos",  action:()=>{setFilterSt("all");setFilterPrazo("red");}},
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

        {/* Valor aberto card — visível só para o admin */}
        {isAdminEquipe&&(
        <div style={{background:"#eff6ff",borderRadius:14,padding:"14px 18px",marginBottom:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:22}}>💰</span>
          <div>
            <div style={{fontSize:11,color:"#64748b",fontWeight:500}}>Valor em Aberto</div>
            <div style={{fontSize:22,fontWeight:800,color:"#0891b2"}}>{fmtBRL(valorAberto)}</div>
          </div>
        </div>
        )}

        {/* Tabs */}
        <TabBar tabs={TABS} active={activeTab} onChange={tab=>{
          if(tab==="financeiro"){
            if(!podeVerFinanceiro){ showToast("🔒 Você não tem permissão para ver o Financeiro.","#991b1b"); return; }
            if(!financeUnlocked){setShowFinanceGate(true);return;}
          }
          setActiveTab(tab);
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
                  {filterProduto!=="all"&&(()=>{
                    const prodPedidos = pedidosAbertos.filter(r=>r.produto===filterProduto);
                    const totalProd   = prodPedidos.length;
                    const feitosProd  = prodPedidos.filter(r=>r.statusInterno==="feito").length;
                    const revisaoProd = prodPedidos.filter(r=>r.statusInterno==="revisao").length;
                    return (
                    <div>
                      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,flexWrap:"wrap"}}>
                        <span style={{fontSize:12,fontWeight:700,color:"#5b21b6"}}>Pedidos de: <span style={{fontWeight:500}}>{filterProduto}</span></span>
                        <span style={{background:"#7c3aed",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>Total: {totalProd}</span>
                        <span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>✅ Feitos: {feitosProd}/{totalProd}</span>
                        {revisaoProd>0&&<span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>📋 Em Revisão: {revisaoProd}</span>}
                      </div>
                      <div style={{display:"flex",flexDirection:"column",gap:5}}>
                        {prodPedidos.map((r,i)=>{
                          const dl=deadlineInfo(r.dataEnvio);
                          const si=r.statusInterno;
                          const corCliente = r.clienteRecorrente ? corRecorrencia(r.loja, lojas) : null;
                          return (
                            <div key={r.idPedido} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",background:si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#f8fafc":"#fff",borderRadius:9,padding:"8px 12px",border:"1px solid",borderColor:si==="feito"?"#86efac":si==="revisao"?"#fde68a":"#f1f5f9",borderLeft:corCliente?`5px solid ${corCliente.color}`:undefined}}>
                              <StyledCheckbox checked={selected.has(r.idPedido)} onChange={()=>toggleSelect(r.idPedido)} />
                              <span style={{fontFamily:"monospace",fontSize:12,fontWeight:700,color:"#1d4ed8",flexShrink:0}}>{r.idPedido}</span>
                              <span style={{fontSize:11,color:"#64748b",minWidth:120}}>👤 <strong>{r.destinatario||"—"}</strong> {r.clienteRecorrente&&<span title={`Cliente recorrente: ${r.totalComprasCliente}x`} style={{background:corCliente.color,color:"#fff",borderRadius:20,padding:"1px 6px",fontSize:9,fontWeight:800,marginLeft:3}}>⭐{r.totalComprasCliente}x</span>}</span>
                              <span style={{fontSize:11,color:"#64748b"}}>🏪 <strong>{r.loja||"—"}</strong></span>
                              <span style={{fontSize:11,color:"#374151",flex:1,minWidth:140}}>📦 {r.produto||"—"}</span>
                              <span style={{fontSize:11,color:"#7c3aed",fontWeight:600,minWidth:100}}>🎨 {r.variacao||"—"}</span>
                              <span style={{fontSize:11,color:"#064e3b",fontWeight:700}}>{fmtBRL(r.preco)}</span>
                              {dl&&<span style={{background:dl.bg,color:dl.text,border:`1px solid ${dl.border}`,borderRadius:7,padding:"1px 7px",fontSize:10,fontWeight:600,flexShrink:0}}>{dl.icon} {dl.label}</span>}
                              {/* Botões feito / revisão */}
                              <div style={{display:"flex",gap:4,flexShrink:0,marginLeft:"auto"}}>
                                <RevisaoButton isRevisao={si==="revisao"} cor={r.feitoPorCor||minhaCor} nome={r.feitoPorNome} onClick={()=>tentarAbrirRevisao(r)} />
                                <FeitoButton isFeito={si==="feito"} cor={r.feitoPorCor||minhaCor} nome={r.feitoPorNome} onClick={()=>si==="feito"?tentarDesmarcar(r,"","Feito"):handleStatusChange(r,"feito")} />
                                <VazioButton isVazio={si==="vazio"} cor={r.feitoPorCor||minhaCor} nome={r.feitoPorNome} onClick={()=>si==="vazio"?tentarDesmarcar(r,"","Vazio"):handleStatusChange(r,"vazio")} />
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    );
                  })()}
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
                  <option value="vazio">📦 Vazio</option>
                </select>
                <select value={filterLoja} onChange={e=>setFilterLoja(e.target.value)} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 11px",fontSize:12,cursor:"pointer",background:"#fff"}}>
                  <option value="all">Todas as lojas</option>
                  {lojas.map(l=><option key={l} value={l}>{l}</option>)}
                </select>
                <select value={filterPrazo} onChange={e=>{setFilterPrazo(e.target.value);setFilterData("all");}} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 11px",fontSize:12,cursor:"pointer",background:"#fff"}}>
                  <option value="all">Todos os prazos</option>
                  <option value="red">🔴 Urgente</option>
                  <option value="yellow">🟡 Atenção (2d)</option>
                  <option value="green">🟢 OK (3+d)</option>
                </select>
                <button onClick={()=>exportToExcel(filtered, "pedidos_em_aberto", {
                  "ID Pedido":"idPedido","Status":"statusPedido","Destinatário":"destinatario","Loja":"loja",
                  "Produto":"produto","Variação":"variacao","Qtd":"quantidade","Preço":r=>r.preco,
                  "Prazo":"dataEnvio","Notas":"notas",
                })} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 14px",fontSize:12,cursor:"pointer",background:"#fff",fontWeight:600,color:"#374151",display:"flex",alignItems:"center",gap:5}}>
                  📥 Exportar
                </button>
              </div>
              {/* Date filter chips */}
              {datasUnicas.length>0&&(
                <div style={{padding:"8px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                  <span style={{fontSize:11,fontWeight:700,color:"#64748b",marginRight:4,whiteSpace:"nowrap"}}>📅 Filtrar por data:</span>
                  <button onClick={()=>setFilterData("all")} style={{padding:"4px 12px",border:`1.5px solid ${filterData==="all"?"#1d4ed8":"#e2e8f0"}`,borderRadius:20,background:filterData==="all"?"#1d4ed8":"#fff",color:filterData==="all"?"#fff":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                    Todos
                  </button>
                  {datasUnicas.map(data=>{
                    const dl = deadlineInfo(data+" 00:00:00");
                    const isActive = filterData===data;
                    const bg = isActive?(dl?.tier==="red"?"#ef4444":dl?.tier==="yellow"?"#f59e0b":"#059669"):"#fff";
                    const border = isActive?bg:(dl?.tier==="red"?"#fca5a5":dl?.tier==="yellow"?"#fde68a":"#86efac");
                    const color = isActive?"#fff":(dl?.tier==="red"?"#991b1b":dl?.tier==="yellow"?"#78350f":"#14532d");
                    const [yyyy,mm,dd] = data.split("-");
                    const pedidosDia = pedidosAbertos.filter(r=>r.dataEnvio?.slice(0,10)===data);
                    const count = pedidosDia.length;
                    const feitosDia = pedidosDia.filter(r=>r.statusInterno==="feito").length;
                    return (
                      <button key={data} onClick={()=>{setFilterData(isActive?"all":data);setFilterPrazo("all");}}
                        style={{padding:"4px 12px",border:`1.5px solid ${border}`,borderRadius:20,background:bg,color,fontSize:11,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",gap:5,whiteSpace:"nowrap"}}>
                        {dl?.tier==="red"?"🔴":dl?.tier==="yellow"?"🟡":"🟢"} {dd}/{mm}
                        <span style={{background:isActive?"rgba(255,255,255,0.3)":"rgba(0,0,0,0.08)",borderRadius:20,padding:"0 5px",fontSize:10,fontWeight:700}}>{feitosDia}/{count}</span>
                      </button>
                    );
                  })}
                </div>
              )}

              <div style={{padding:"8px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                <span style={{fontSize:12,color:"#64748b"}}><strong>{filtered.length}</strong> pedidos · <strong>{selected.size}</strong> selecionados</span>
                {selected.size>0&&<>
                  <button onClick={copyIds} style={{background:copied?"#059669":"#1d4ed8",color:"#fff",border:"none",borderRadius:10,padding:"5px 13px",fontSize:12,fontWeight:700,cursor:"pointer"}}>{copied?"✓ Copiado!`":`📋 Copiar ${selected.size} ID(s)`}</button>
                  {selected.size>1&&<button onClick={markAllFeito} style={{background:"#059669",color:"#fff",border:"none",borderRadius:10,padding:"5px 13px",fontSize:12,fontWeight:700,cursor:"pointer"}}>✅ Marcar {selected.size} como Feito</button>}
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
                      const corCliente = row.clienteRecorrente ? corRecorrencia(row.loja, lojas) : null;
                      return (
                        <tr key={row.idPedido} style={{background:isSel?"#eff6ff":si==="feito"?"#f0fdf4":si==="revisao"?"#fffbeb":i%2===0?"#fff":"#fafafa",borderBottom:"1px solid #f1f5f9",borderLeft:corCliente?`5px solid ${corCliente.color}`:"5px solid transparent"}}>
                          <td style={{...TD,textAlign:"center"}}><StyledCheckbox checked={isSel} onChange={()=>toggleSelect(row.idPedido)} /></td>
                          <td style={{...TD,whiteSpace:"nowrap"}}>
                            <div style={{display:"flex",gap:4}}>
                              <RevisaoButton isRevisao={si==="revisao"} cor={row.feitoPorCor||minhaCor} nome={row.feitoPorNome} onClick={()=>tentarAbrirRevisao(row)} small />
                              <FeitoButton isFeito={si==="feito"} cor={row.feitoPorCor||minhaCor} nome={row.feitoPorNome} onClick={()=>si==="feito"?tentarDesmarcar(row,"","Feito"):handleStatusChange(row,"feito")} small />
                              <VazioButton isVazio={si==="vazio"} cor={row.feitoPorCor||minhaCor} nome={row.feitoPorNome} onClick={()=>si==="vazio"?tentarDesmarcar(row,"","Vazio"):handleStatusChange(row,"vazio")} small />
                            </div>
                          </td>
                          <td style={{...TD,textAlign:"center"}}>
                            {si==="feito"?<span style={{background:"#d1fae5",color:"#065f46",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:700}}>✅ Feito</span>
                            :si==="revisao"?<span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:700}}>📋 Rev.</span>
                            :<StatusBadge status={row.statusPedido} />}
                          </td>
                          <td style={{...TD,fontWeight:700,color:"#1d4ed8",fontFamily:"monospace",fontSize:11}}>{row.idPedido}</td>
                          <td style={TD}>
                            <div style={{display:"flex",alignItems:"center",gap:5}}>
                              <ExpandCell value={row.destinatario||"—"} />
                              {row.clienteRecorrente&&<span title={`Cliente recorrente: ${row.totalComprasCliente}x no histórico`} style={{background:corCliente.color,color:"#fff",borderRadius:20,padding:"1px 7px",fontSize:9,fontWeight:800,flexShrink:0}}>⭐ {row.totalComprasCliente}x</span>}
                            </div>
                          </td>
                          <td style={TD}>{row.loja||"—"}</td>
                          <td style={{...TD,maxWidth:200}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:200}}>{row.produto||"—"}</div></td>
                          <td style={{...TD,whiteSpace:"normal",maxWidth:200}}>{row.variacao||"—"}</td>
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
        {activeTab==="enviados"&&(()=>{
          const qE = searchEnviados.trim().toLowerCase();
          const datasEnviados = [...new Set(pedidosEnviados.map(r=>r.dataEnvio?.slice(0,10)).filter(Boolean))].sort().reverse();
          let enviadosFiltrados = !qE ? pedidosEnviados : pedidosEnviados.filter(r=>
            r.idPedido.toLowerCase().includes(qE) ||
            r.produto.toLowerCase().includes(qE) ||
            r.destinatario.toLowerCase().includes(qE) ||
            r.variacao.toLowerCase().includes(qE)
          );
          if (filterDataEnviados!=="all") enviadosFiltrados = enviadosFiltrados.filter(r=>r.dataEnvio?.slice(0,10)===filterDataEnviados);
          return (
          <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#059669"}}>📦 Pedidos Enviados / Entregues</span>
              <span style={{background:"#059669",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{enviadosFiltrados.length}</span>
              <input placeholder="🔍 Buscar ID, produto, destinatário..." value={searchEnviados} onChange={e=>setSearchEnviados(e.target.value)}
                style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"6px 12px",fontSize:12,outline:"none",flex:1,minWidth:180,maxWidth:320}} />
              <button onClick={()=>exportToExcel(enviadosFiltrados, "pedidos_enviados", {
                "ID Pedido":"idPedido","Status":"statusPedido","Destinatário":"destinatario","Loja":"loja",
                "Produto":"produto","Variação":"variacao","Qtd":"quantidade","Preço":r=>r.preco,"Data Envio":"dataEnvio",
              })} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"6px 13px",fontSize:11,cursor:"pointer",background:"#fff",fontWeight:600,color:"#374151"}}>
                📥 Exportar
              </button>
              <span style={{marginLeft:"auto",fontSize:12,fontWeight:700,color:"#059669"}}>{fmtBRL(enviadosFiltrados.reduce((s,r)=>s+(parseFloat(r.preco)||0),0))}</span>
            </div>
            {datasEnviados.length>0&&(
              <div style={{padding:"9px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                <span style={{fontSize:11,fontWeight:700,color:"#64748b",marginRight:2,whiteSpace:"nowrap"}}>📅 Data de envio:</span>
                <button onClick={()=>setFilterDataEnviados("all")} style={{padding:"4px 12px",border:`1.5px solid ${filterDataEnviados==="all"?"#059669":"#e2e8f0"}`,borderRadius:20,background:filterDataEnviados==="all"?"#059669":"#fff",color:filterDataEnviados==="all"?"#fff":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>Todas</button>
                {datasEnviados.slice(0,30).map(data=>{
                  const isActive = filterDataEnviados===data;
                  const [yyyy,mm,dd] = data.split("-");
                  const count = pedidosEnviados.filter(r=>r.dataEnvio?.slice(0,10)===data).length;
                  return (
                    <button key={data} onClick={()=>setFilterDataEnviados(isActive?"all":data)}
                      style={{padding:"4px 12px",border:`1.5px solid ${isActive?"#059669":"#e2e8f0"}`,borderRadius:20,background:isActive?"#059669":"#fff",color:isActive?"#fff":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                      {dd}/{mm} <span style={{opacity:0.8}}>({count})</span>
                    </button>
                  );
                })}
              </div>
            )}
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#f0fdf4"}}>{["Status","ID Pedido","Destinatário","Loja","Produto","Variação","Qtd","Preço","Data Envio"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                <tbody>
                  {enviadosFiltrados.map((row,i)=>{
                    const corCliente = row.clienteRecorrente ? corRecorrencia(row.loja, lojas) : null;
                    return (
                    <tr key={row.idPedido} style={{background:i%2===0?"#f8fffe":"#fff",borderBottom:"1px solid #f0fdf4",borderLeft:corCliente?`5px solid ${corCliente.color}`:"5px solid transparent"}}>
                      <td style={{...TD,textAlign:"center"}}><StatusBadge status={row.statusPedido} /></td>
                      <td style={{...TD,fontWeight:700,color:"#059669",fontFamily:"monospace",fontSize:11}}>{row.idPedido}</td>
                      <td style={TD}>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <ExpandCell value={row.destinatario||"—"} />
                          {row.clienteRecorrente&&<span title={`Cliente recorrente: ${row.totalComprasCliente}x no histórico`} style={{background:corCliente.color,color:"#fff",borderRadius:20,padding:"1px 7px",fontSize:9,fontWeight:800,flexShrink:0}}>⭐ {row.totalComprasCliente}x</span>}
                        </div>
                      </td>
                      <td style={TD}>{row.loja||"—"}</td>
                      <td style={{...TD,maxWidth:190}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:190}}>{row.produto||"—"}</div></td>
                      <td style={{...TD,whiteSpace:"normal",maxWidth:200}}>{row.variacao||"—"}</td>
                      <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,fontWeight:600,color:"#059669"}}>{fmtBRL(row.preco)}</td>
                      <td style={TD}>{row.dataEnvio?.slice(0,10)||"—"}</td>
                    </tr>
                    );
                  })}
                  {enviadosFiltrados.length===0&&<tr><td colSpan={9} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>{pedidosEnviados.length===0?"Nenhum pedido enviado ainda.":"Nenhum pedido encontrado."}</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
          );
        })()}

        {/* ── TAB: CANCELADOS ── */}
        {activeTab==="cancelados"&&(()=>{
          // Agrupa cancelados por motivo (texto já simplificado) para os chips de filtro
          const motivoContagem = {};
          for (const r of pedidosCancelados) {
            const info = formatMotivoCancelamento(r.motivoCancelamento);
            const chave = info ? info.texto : "Motivo não informado";
            if (!motivoContagem[chave]) motivoContagem[chave] = { texto: chave, total: 0, origem: info?.origem || "—", cor: info?.cor || { bg:"#f1f5f9", color:"#64748b" } };
            motivoContagem[chave].total++;
          }
          const motivosArr = Object.values(motivoContagem).sort((a,b)=>b.total-a.total);

          const canceladosFiltrados0 = filterMotivo==="all"
            ? pedidosCancelados
            : pedidosCancelados.filter(r => {
                const info = formatMotivoCancelamento(r.motivoCancelamento);
                const chave = info ? info.texto : "Motivo não informado";
                return chave === filterMotivo;
              });
          const datasCancelados = [...new Set(pedidosCancelados.map(r=>r.dataEnvio?.slice(0,10)).filter(Boolean))].sort().reverse();
          const canceladosFiltrados = filterDataCancelados==="all" ? canceladosFiltrados0 : canceladosFiltrados0.filter(r=>r.dataEnvio?.slice(0,10)===filterDataCancelados);

          return (
          <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#6b7280"}}>❌ Pedidos Cancelados</span>
              <span style={{background:"#6b7280",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{canceladosFiltrados.length}</span>
              <button onClick={()=>exportToExcel(canceladosFiltrados, "pedidos_cancelados", {
                "ID Pedido":"idPedido","Status":"statusPedido","Destinatário":"destinatario","Loja":"loja",
                "Produto":"produto","Variação":"variacao","Qtd":"quantidade","Preço":r=>r.preco,"Data":"dataEnvio",
                "Motivo do Cancelamento": r => formatMotivoCancelamento(r.motivoCancelamento)?.texto || "",
              })} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"6px 13px",fontSize:11,cursor:"pointer",background:"#fff",fontWeight:600,color:"#374151"}}>
                📥 Exportar
              </button>
            </div>

            {/* Filtro por motivo de cancelamento */}
            {motivosArr.length>0&&(
              <div style={{padding:"10px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                <span style={{fontSize:11,fontWeight:700,color:"#64748b",marginRight:2,whiteSpace:"nowrap"}}>🔍 Filtrar por motivo:</span>
                <button onClick={()=>setFilterMotivo("all")} style={{padding:"4px 12px",border:`1.5px solid ${filterMotivo==="all"?"#6b7280":"#e2e8f0"}`,borderRadius:20,background:filterMotivo==="all"?"#6b7280":"#fff",color:filterMotivo==="all"?"#fff":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                  Todos ({pedidosCancelados.length})
                </button>
                {motivosArr.map(m=>{
                  const isActive = filterMotivo===m.texto;
                  return (
                    <button key={m.texto} onClick={()=>setFilterMotivo(isActive?"all":m.texto)}
                      title={m.texto}
                      style={{padding:"4px 12px",border:`1.5px solid ${isActive?m.cor.color:"#e2e8f0"}`,borderRadius:20,background:isActive?m.cor.color:m.cor.bg,color:isActive?"#fff":m.cor.color,fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap",maxWidth:280,overflow:"hidden",textOverflow:"ellipsis"}}>
                      {m.texto.length>40?m.texto.slice(0,40)+"…":m.texto} <strong>({m.total})</strong>
                    </button>
                  );
                })}
              </div>
            )}

            {datasCancelados.length>0&&(
              <div style={{padding:"9px 18px",background:"#f8fafc",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                <span style={{fontSize:11,fontWeight:700,color:"#64748b",marginRight:2,whiteSpace:"nowrap"}}>📅 Data:</span>
                <button onClick={()=>setFilterDataCancelados("all")} style={{padding:"4px 12px",border:`1.5px solid ${filterDataCancelados==="all"?"#6b7280":"#e2e8f0"}`,borderRadius:20,background:filterDataCancelados==="all"?"#6b7280":"#fff",color:filterDataCancelados==="all"?"#fff":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>Todas</button>
                {datasCancelados.slice(0,30).map(data=>{
                  const isActive = filterDataCancelados===data;
                  const [yyyy,mm,dd] = data.split("-");
                  const count = pedidosCancelados.filter(r=>r.dataEnvio?.slice(0,10)===data).length;
                  return (
                    <button key={data} onClick={()=>setFilterDataCancelados(isActive?"all":data)}
                      style={{padding:"4px 12px",border:`1.5px solid ${isActive?"#6b7280":"#e2e8f0"}`,borderRadius:20,background:isActive?"#6b7280":"#fff",color:isActive?"#fff":"#374151",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                      {dd}/{mm} <span style={{opacity:0.8}}>({count})</span>
                    </button>
                  );
                })}
              </div>
            )}

            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#f9fafb"}}>{["Status","ID Pedido","Destinatário","Loja","Produto","Variação","Qtd","Preço","Data","Motivo do Cancelamento"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                <tbody>
                  {canceladosFiltrados.map((row,i)=>{
                    const corCliente = row.clienteRecorrente ? corRecorrencia(row.loja, lojas) : null;
                    return (
                    <tr key={row.idPedido} style={{background:i%2===0?"#f9fafb":"#fff",borderBottom:"1px solid #f3f4f6",opacity:0.85,borderLeft:corCliente?`5px solid ${corCliente.color}`:"5px solid transparent"}}>
                      <td style={{...TD,textAlign:"center"}}><StatusBadge status={row.statusPedido} /></td>
                      <td style={{...TD,fontWeight:700,color:"#6b7280",fontFamily:"monospace",fontSize:11}}>{row.idPedido}</td>
                      <td style={TD}>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <ExpandCell value={row.destinatario||"—"} />
                          {row.clienteRecorrente&&<span title={`Cliente recorrente: ${row.totalComprasCliente}x no histórico`} style={{background:corCliente.color,color:"#fff",borderRadius:20,padding:"1px 7px",fontSize:9,fontWeight:800,flexShrink:0}}>⭐ {row.totalComprasCliente}x</span>}
                        </div>
                      </td>
                      <td style={TD}>{row.loja||"—"}</td>
                      <td style={{...TD,maxWidth:190}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:190,textDecoration:"line-through",color:"#9ca3af"}}>{row.produto||"—"}</div></td>
                      <td style={{...TD,whiteSpace:"normal",maxWidth:200}}>{row.variacao||"—"}</td>
                      <td style={{...TD,textAlign:"center",color:"#9ca3af"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,fontWeight:600,color:"#9ca3af",textDecoration:"line-through"}}>{fmtBRL(row.preco)}</td>
                      <td style={{...TD,color:"#9ca3af"}}>{row.dataEnvio?.slice(0,10)||"—"}</td>
                      <td style={{...TD,maxWidth:220}}>
                        {(() => {
                          const info = formatMotivoCancelamento(row.motivoCancelamento);
                          if (!info) return <span style={{color:"#cbd5e1"}}>—</span>;
                          return (
                            <div style={{display:"flex",flexDirection:"column",gap:3}}>
                              <span style={{background:info.cor.bg,color:info.cor.color,borderRadius:6,padding:"1px 7px",fontSize:9,fontWeight:700,width:"fit-content"}}>{info.origem}</span>
                              <span style={{fontSize:11,color:"#475569",lineHeight:1.3}}>{info.texto}</span>
                            </div>
                          );
                        })()}
                      </td>
                    </tr>
                    );
                  })}
                  {canceladosFiltrados.length===0&&<tr><td colSpan={10} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>{pedidosCancelados.length===0?"Nenhum pedido cancelado.":"Nenhum pedido encontrado com esse motivo."}</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
          );
        })()}

        {/* ── TAB: DEVOLUCOES ── */}
        {activeTab==="devolucoes"&&(()=>{
          const motivoDevContagem = {};
          for (const r of devolucoes) {
            const chave = (r.motivo || "Motivo não informado").trim();
            motivoDevContagem[chave] = (motivoDevContagem[chave] || 0) + 1;
          }
          const motivosDevArr = Object.entries(motivoDevContagem).sort((a,b)=>b[1]-a[1]);
          const devolucoesFiltradas = filterMotivoDevolucao==="all"
            ? devolucoes
            : devolucoes.filter(r => (r.motivo || "Motivo não informado").trim() === filterMotivoDevolucao);
          const totalDevFiltrado = devolucoesFiltradas.reduce((s,r)=>s+(parseFloat(r.preco_unidade)||0)*(parseFloat(r.quantidade)||1),0);

          return (
          <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
            <div style={{padding:"14px 18px",borderBottom:"1px solid #fee2e2",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#ef4444"}}>🔄 Devoluções e Reembolsos</span>
              <span style={{background:"#ef4444",color:"#fff",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>{devolucoesFiltradas.length}</span>
              {totalDevFiltrado>0&&<span style={{background:"#fde8e8",color:"#991b1b",borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:800}}>-{fmtBRL(totalDevFiltrado)}</span>}
              <button onClick={()=>exportToExcel(devolucoesFiltradas, "devolucoes", {
                "ID Pedido":"id_pedido","Data":"data_criacao","Produto":"produto","Variação":"variacao",
                "Qtd":"quantidade","Preço Unit.":r=>r.preco_unidade,"Status Devolução":"status_devolucao",
                "Motivo":"motivo","Observações":"observacoes","Loja":"loja",
              })} style={{border:"1px solid #fee2e2",borderRadius:10,padding:"6px 13px",fontSize:11,cursor:"pointer",background:"#fff",fontWeight:600,color:"#374151"}}>
                📥 Exportar
              </button>
            </div>

            {motivosDevArr.length>0&&(
              <div style={{padding:"10px 18px",background:"#fef2f2",borderBottom:"1px solid #fee2e2",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                <span style={{fontSize:11,fontWeight:700,color:"#991b1b",marginRight:2,whiteSpace:"nowrap"}}>🔍 Filtrar por motivo:</span>
                <button onClick={()=>setFilterMotivoDevolucao("all")} style={{padding:"4px 12px",border:`1.5px solid ${filterMotivoDevolucao==="all"?"#ef4444":"#fecaca"}`,borderRadius:20,background:filterMotivoDevolucao==="all"?"#ef4444":"#fff",color:filterMotivoDevolucao==="all"?"#fff":"#991b1b",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
                  Todos ({devolucoes.length})
                </button>
                {motivosDevArr.map(([motivo,total])=>{
                  const isActive = filterMotivoDevolucao===motivo;
                  return (
                    <button key={motivo} onClick={()=>setFilterMotivoDevolucao(isActive?"all":motivo)}
                      title={motivo}
                      style={{padding:"4px 12px",border:`1.5px solid ${isActive?"#ef4444":"#fecaca"}`,borderRadius:20,background:isActive?"#ef4444":"#fff",color:isActive?"#fff":"#991b1b",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap",maxWidth:260,overflow:"hidden",textOverflow:"ellipsis"}}>
                      {motivo.length>35?motivo.slice(0,35)+"…":motivo} <strong>({total})</strong>
                    </button>
                  );
                })}
              </div>
            )}

            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#fef2f2"}}>{["ID Pedido","Data","Produto","Variação","Qtd","Preço Unit.","Status Dev.","Motivo","Observações","Loja"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                <tbody>
                  {devolucoesFiltradas.map((row,i)=>(
                    <tr key={row.id_pedido||i} style={{background:i%2===0?"#fff5f5":"#fff",borderBottom:"1px solid #fee2e2"}}>
                      <td style={{...TD,fontWeight:700,color:"#ef4444",fontFamily:"monospace",fontSize:11}}>{row.id_pedido}</td>
                      <td style={TD}>{row.data_criacao?.slice(0,10)||"—"}</td>
                      <td style={{...TD,maxWidth:180}}><div style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"normal",lineHeight:1.3,maxWidth:180}}>{row.produto||"—"}</div></td>
                      <td style={{...TD,whiteSpace:"normal",maxWidth:200}}>{row.variacao||"—"}</td>
                      <td style={{...TD,textAlign:"center"}}>{row.quantidade||"—"}</td>
                      <td style={{...TD,color:"#ef4444",fontWeight:600}}>-{fmtBRL(row.preco_unidade)}</td>
                      <td style={TD}><span style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:600}}>{row.status_devolucao||"—"}</span></td>
                      <td style={{...TD,maxWidth:160}}><ExpandCell value={row.motivo||"—"} maxLen={25} /></td>
                      <td style={{...TD,maxWidth:160}}><ExpandCell value={row.observacoes||"—"} maxLen={25} /></td>
                      <td style={TD}>{row.loja||"—"}</td>
                    </tr>
                  ))}
                  {devolucoesFiltradas.length===0&&<tr><td colSpan={10} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>{devolucoes.length===0?"Nenhuma devolução registrada.":"Nenhuma devolução encontrada com esse motivo."}</td></tr>}
                </tbody>
              </table>
            </div>
          </div>
          );
        })()}

        {/* ── TAB: CLIENTES ── */}
        {activeTab==="clientes"&&(()=>{
          // Meses únicos presentes nos pedidos (com base na data de criação)
          const mesesUnicos = [...new Set(allPedidos.map(r=>r.mesReferencia).filter(Boolean))].sort().reverse();

          // Total HISTÓRICO por cliente+loja (todas as planilhas, todos os meses)
          // — usado para decidir se o cliente é "recorrente" (2+ compras no total)
          const totalHistorico = {};
          for (const r of allPedidos) {
            if (!r.nomeUsuario) continue;
            const chave = `${r.nomeUsuario}__${r.loja}`;
            totalHistorico[chave] = (totalHistorico[chave] || 0) + 1;
          }

          // Lista exibida respeita o filtro de mês selecionado
          const pedidosDoMes = mesClientesFiltro==="all" ? allPedidos : allPedidos.filter(r=>r.mesReferencia===mesClientesFiltro);
          const ranking = {};
          for (const r of pedidosDoMes) {
            if (!r.nomeUsuario) continue;
            const chave = `${r.nomeUsuario}__${r.loja}`;
            if (!ranking[chave]) ranking[chave] = { nomeUsuario: r.nomeUsuario, loja: r.loja, totalMes: 0 };
            ranking[chave].totalMes++;
          }
          const rankingArr = Object.values(ranking)
            .map(c => ({ ...c, totalGeral: totalHistorico[`${c.nomeUsuario}__${c.loja}`] || c.totalMes }))
            .sort((a,b)=>b.totalGeral-a.totalGeral);
          const totalClientesUnicos = rankingArr.length;
          const totalCompras = rankingArr.reduce((s,r)=>s+r.totalMes,0);

          return (
          <div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18,flexWrap:"wrap",gap:10}}>
              <div>
                <div style={{fontSize:16,fontWeight:800,color:"#1e293b"}}>👤 Clientes Recorrentes</div>
                <div style={{fontSize:12,color:"#64748b"}}>Compras no mês selecionado · destaque considera o histórico completo do cliente</div>
              </div>
              <select value={mesClientesFiltro} onChange={e=>setMesClientesFiltro(e.target.value)}
                style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"8px 14px",fontSize:13,cursor:"pointer",background:"#fff",fontWeight:600}}>
                <option value="all">Todos os meses</option>
                {mesesUnicos.map(m=><option key={m} value={m}>{mesLabel(m)}</option>)}
              </select>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:14,marginBottom:18}}>
              <div style={{background:"#fff",borderRadius:16,padding:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:11,color:"#64748b",fontWeight:600}}>Clientes únicos</div>
                <div style={{fontSize:26,fontWeight:800,color:"#0d9488"}}>{totalClientesUnicos}</div>
              </div>
              <div style={{background:"#fff",borderRadius:16,padding:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:11,color:"#64748b",fontWeight:600}}>Total de compras</div>
                <div style={{fontSize:26,fontWeight:800,color:"#0d9488"}}>{totalCompras}</div>
              </div>
            </div>

            <div style={{display:"flex",gap:14,marginBottom:14,fontSize:12,fontWeight:600,color:"#374151"}}>
              <span style={{display:"flex",alignItems:"center",gap:6}}><span style={{width:14,height:14,borderRadius:4,background:corRecorrencia(lojas[0],lojas).bg,display:"inline-block"}}/> Recorrente · {lojas[0]||"Loja 1"}</span>
              <span style={{display:"flex",alignItems:"center",gap:6}}><span style={{width:14,height:14,borderRadius:4,background:corRecorrencia(lojas[1],lojas).bg,display:"inline-block"}}/> Recorrente · {lojas[1]||"Loja 2"}</span>
            </div>

            <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
              <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",fontSize:14,fontWeight:700,color:"#1e293b",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                <span>🏆 Ranking — {mesClientesFiltro==="all" ? "Todos os meses" : mesLabel(mesClientesFiltro)}</span>
                <button onClick={()=>exportToExcel(rankingArr, "ranking_clientes", {
                  "Nome de Usuário":"nomeUsuario","Loja":"loja","Compras no mês":"totalMes","Total geral":"totalGeral",
                })} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"6px 13px",fontSize:11,cursor:"pointer",background:"#fff",fontWeight:600,color:"#374151"}}>
                  📥 Exportar
                </button>
              </div>
              <div style={{overflowX:"auto",maxHeight:620,overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{background:"#f0fdfa",position:"sticky",top:0}}>
                    {["","#","Nome de Usuário","Loja","Compras no mês","Total geral"].map(h=><th key={h} style={TH}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {rankingArr.map((c,i)=>{
                      const recorrente = c.totalGeral > 1;
                      const cores = corRecorrencia(c.loja, lojas);
                      const chave = `${c.nomeUsuario}__${c.loja}`;
                      const expandido = clienteExpandido === chave;
                      // Histórico completo de compras desse cliente (todos os meses) para o detalhe
                      const comprasCliente = allPedidos.filter(r=>r.nomeUsuario===c.nomeUsuario && r.loja===c.loja);
                      // Agrupa por produto+variação para mostrar "o que ele mais compra"
                      const porProduto = {};
                      for (const p of comprasCliente) {
                        const k = `${p.produto}__${p.variacao}`;
                        if (!porProduto[k]) porProduto[k] = { produto:p.produto, variacao:p.variacao, qtd:0, ultimaData:p.dataCriacao };
                        porProduto[k].qtd++;
                        if (p.dataCriacao > porProduto[k].ultimaData) porProduto[k].ultimaData = p.dataCriacao;
                      }
                      const produtosArr = Object.values(porProduto).sort((a,b)=>b.qtd-a.qtd);
                      return (
                      <>
                      <tr key={chave} onClick={()=>setClienteExpandido(expandido?null:chave)}
                        style={{background:recorrente?cores.bg:(i%2===0?"#fff":"#f8fafc"),borderBottom:expandido?"none":"1px solid #f1f5f9",cursor:"pointer"}}>
                        <td style={{...TD,width:30,textAlign:"center"}}>
                          <span style={{display:"inline-block",transition:"transform 0.15s",transform:expandido?"rotate(90deg)":"none",color:recorrente?cores.color:"#94a3b8",fontWeight:700}}>▶</span>
                        </td>
                        <td style={{...TD,color:recorrente?cores.color:"#94a3b8",fontWeight:700}}>{i+1}</td>
                        <td style={{...TD,fontWeight:800,color:recorrente?cores.color:"#1e293b"}}>{c.nomeUsuario}</td>
                        <td style={{...TD,fontWeight:recorrente?700:400,color:recorrente?cores.color:"#374151"}}>{c.loja}</td>
                        <td style={{...TD,fontWeight:recorrente?700:400,color:recorrente?cores.color:"#374151"}}>{c.totalMes}x</td>
                        <td style={TD}>
                          <span style={{background:recorrente?cores.color:"#f1f5f9",color:recorrente?"#fff":"#64748b",borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:800}}>
                            {c.totalGeral}x {recorrente?"⭐":""}
                          </span>
                        </td>
                      </tr>
                      {expandido && (
                        <tr style={{background:recorrente?cores.bgLight:"#f8fafc",borderBottom:"1px solid #f1f5f9"}}>
                          <td colSpan={6} style={{padding:"12px 18px 16px 46px"}}>
                            <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:8}}>📦 Produtos que {c.nomeUsuario} já comprou ({comprasCliente.length} pedido{comprasCliente.length!==1?"s":""} no total):</div>
                            <div style={{display:"flex",flexDirection:"column",gap:5}}>
                              {produtosArr.map((p,j)=>(
                                <div key={j} style={{display:"flex",alignItems:"center",gap:8,background:"#fff",borderRadius:8,padding:"7px 11px",border:"1px solid #f1f5f9"}}>
                                  <span style={{fontSize:11,color:"#374151",flex:1}}>{p.produto||"—"}</span>
                                  {p.variacao&&<span style={{fontSize:10,color:"#7c3aed",fontWeight:600}}>🎨 {p.variacao}</span>}
                                  <span style={{background:p.qtd>1?cores.color:"#f1f5f9",color:p.qtd>1?"#fff":"#64748b",borderRadius:20,padding:"1px 9px",fontSize:10,fontWeight:700,flexShrink:0}}>{p.qtd}x</span>
                                </div>
                              ))}
                              {produtosArr.length===0&&<div style={{fontSize:11,color:"#94a3b8"}}>Nenhum produto encontrado.</div>}
                            </div>
                          </td>
                        </tr>
                      )}
                      </>
                      );
                    })}
                    {rankingArr.length===0&&<tr><td colSpan={6} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>Nenhum dado de cliente encontrado. Carregue uma planilha com a coluna "Nome de usuário (comprador)".</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          );
        })()}

        {/* ── TAB: FINANCEIRO ── */}
        {activeTab==="financeiro"&&financeUnlocked&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:18}}>
              {lojas.map(loja=>{
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

            {/* ── Gráfico: Pedidos por dia (últimos 30 dias com dados) ── */}
            <div style={{background:"#fff",borderRadius:18,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginTop:18}}>
              <div style={{fontSize:14,fontWeight:700,color:"#1e293b",marginBottom:16}}>📊 Pedidos por Dia</div>
              {(() => {
                const porDia = {};
                for (const r of allPedidos) {
                  const d = r.dataCriacao?.slice(0,10);
                  if (!d) continue;
                  if (!porDia[d]) porDia[d] = { dia: d, total: 0, cancelados: 0 };
                  porDia[d].total++;
                  if (isCancelado(r)) porDia[d].cancelados++;
                }
                const dadosDia = Object.values(porDia).sort((a,b)=>a.dia.localeCompare(b.dia)).slice(-30)
                  .map(d => ({ ...d, diaLabel: d.dia.slice(8,10)+"/"+d.dia.slice(5,7) }));
                if (dadosDia.length===0) return <div style={{textAlign:"center",padding:40,color:"#94a3b8",fontSize:13}}>Sem dados suficientes ainda.</div>;
                return (
                  <ResponsiveContainer width="100%" height={240}>
                    <LineChart data={dadosDia} margin={{top:5,right:10,left:10,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                      <XAxis dataKey="diaLabel" tick={{fontSize:10}} />
                      <YAxis tick={{fontSize:11}} allowDecimals={false} />
                      <Tooltip />
                      <Legend />
                      <Line type="monotone" dataKey="total" name="Pedidos" stroke="#1d4ed8" strokeWidth={2} dot={false} />
                      <Line type="monotone" dataKey="cancelados" name="Cancelados" stroke="#ef4444" strokeWidth={2} dot={false} />
                    </LineChart>
                  </ResponsiveContainer>
                );
              })()}
            </div>

            {/* ── Gráfico: Taxa de cancelamento ao longo do tempo ── */}
            <div style={{background:"#fff",borderRadius:18,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginTop:18}}>
              <div style={{fontSize:14,fontWeight:700,color:"#1e293b",marginBottom:16}}>📉 Taxa de Cancelamento por Mês</div>
              {(() => {
                const porMes = {};
                for (const r of allPedidos) {
                  const m = r.mesReferencia;
                  if (!m) continue;
                  if (!porMes[m]) porMes[m] = { mes: m, total: 0, cancelados: 0 };
                  porMes[m].total++;
                  if (isCancelado(r)) porMes[m].cancelados++;
                }
                const dadosMes = Object.values(porMes).sort((a,b)=>a.mes.localeCompare(b.mes))
                  .map(d => ({ ...d, mesLabel: mesLabel(d.mes), taxa: d.total ? Number(((d.cancelados/d.total)*100).toFixed(1)) : 0 }));
                if (dadosMes.length===0) return <div style={{textAlign:"center",padding:40,color:"#94a3b8",fontSize:13}}>Sem dados suficientes ainda.</div>;
                return (
                  <ResponsiveContainer width="100%" height={240}>
                    <LineChart data={dadosMes} margin={{top:5,right:10,left:10,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                      <XAxis dataKey="mesLabel" tick={{fontSize:11}} />
                      <YAxis tick={{fontSize:11}} unit="%" />
                      <Tooltip formatter={v=>`${v}%`} />
                      <Line type="monotone" dataKey="taxa" name="Taxa de cancelamento" stroke="#ef4444" strokeWidth={2} />
                    </LineChart>
                  </ResponsiveContainer>
                );
              })()}
            </div>
          </div>
        )}

        {/* ── TAB: EQUIPE ── */}
        {activeTab==="equipe"&&isAdminEquipe&&organizacao&&(()=>{
          const hoje = new Date(); hoje.setHours(0,0,0,0);
          const inicioSemana = new Date(hoje); inicioSemana.setDate(hoje.getDate()-hoje.getDay());
          const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);

          const statsPorMembro = membrosEquipe.map(m => {
            const feitosDoMembro = produtividade.filter(p => p.feito_por_user_id===m.user_id && p.status_interno==="feito");
            const revisaoDoMembro = produtividade.filter(p => p.feito_por_user_id===m.user_id && p.status_interno==="revisao");
            const contarPeriodo = (lista, desde) => lista.filter(p => p.feito_em && new Date(p.feito_em) >= desde).length;
            return {
              membro: m,
              feitoHoje:   contarPeriodo(feitosDoMembro, hoje),
              feitoSemana: contarPeriodo(feitosDoMembro, inicioSemana),
              feitoMes:    contarPeriodo(feitosDoMembro, inicioMes),
              feitoTotal:  feitosDoMembro.length,
              revisaoTotal: revisaoDoMembro.length,
            };
          });
          const maxFeitoMes = Math.max(1, ...statsPorMembro.map(s=>s.feitoMes));

          return (
          <div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18}}>
              <div>
                <div style={{fontSize:16,fontWeight:800,color:"#1e293b"}}>👥 Gerenciar Equipe</div>
                <div style={{fontSize:12,color:"#64748b"}}>{organizacao.nome} · {membrosEquipe.length} membro{membrosEquipe.length!==1?"s":""}</div>
              </div>
              <button onClick={()=>setShowNovoFuncionario(true)} style={{background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",color:"#fff",border:"none",borderRadius:12,padding:"10px 18px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                + Novo Funcionário
              </button>
            </div>

            {/* Cards de produtividade */}
            <div style={{background:"#fff",borderRadius:18,padding:20,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:18}}>
              <div style={{fontSize:14,fontWeight:700,color:"#1e293b",marginBottom:16}}>📊 Produtividade da Equipe</div>
              <div style={{display:"flex",flexDirection:"column",gap:14}}>
                {statsPorMembro.map(s=>(
                  <div key={s.membro.id} style={{display:"flex",alignItems:"center",gap:14}}>
                    <div style={{width:34,height:34,borderRadius:"50%",background:s.membro.cor,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:700,fontSize:13,flexShrink:0}}>
                      {s.membro.nome.slice(0,1).toUpperCase()}
                    </div>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                        <span style={{fontSize:13,fontWeight:600,color:"#1e293b"}}>{s.membro.nome}{s.membro.is_admin&&<span style={{fontSize:10,color:"#7c3aed",marginLeft:6}}>(admin)</span>}</span>
                        <span style={{fontSize:12,color:"#64748b"}}>
                          <strong style={{color:"#059669"}}>{s.feitoHoje}</strong> hoje · <strong>{s.feitoSemana}</strong> semana · <strong>{s.feitoMes}</strong> mês · <strong>{s.feitoTotal}</strong> total
                          {s.revisaoTotal>0&&<span style={{color:"#f59e0b"}}> · 📋{s.revisaoTotal}</span>}
                        </span>
                      </div>
                      <div style={{height:8,background:"#f1f5f9",borderRadius:6,overflow:"hidden"}}>
                        <div style={{height:"100%",width:`${(s.feitoMes/maxFeitoMes)*100}%`,background:s.membro.cor,borderRadius:6,transition:"width 0.3s"}}/>
                      </div>
                    </div>
                  </div>
                ))}
                {statsPorMembro.length===0&&<div style={{textAlign:"center",padding:20,color:"#94a3b8",fontSize:13}}>Nenhum membro cadastrado ainda.</div>}
              </div>
            </div>

            {/* Google Calendar — alarmes automáticos de prazo */}
            <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",padding:20,marginBottom:18}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:googleCalStatus?16:0}}>
                <div style={{display:"flex",alignItems:"center",gap:12}}>
                  <span style={{fontSize:28}}>📅</span>
                  <div>
                    <div style={{fontSize:14,fontWeight:700,color:"#1e293b"}}>Alarmes no Google Calendar</div>
                    <div style={{fontSize:12,color:"#64748b"}}>
                      {googleCalStatus
                        ? "Conectado — cria eventos automáticos nos horários configurados abaixo"
                        : "Não conectado. Conecte para receber alarmes de pedidos vencendo direto na sua agenda."}
                    </div>
                  </div>
                </div>
                {conectandoGoogle ? (
                  <span style={{fontSize:12,color:"#64748b",fontWeight:600}}>Conectando...</span>
                ) : googleCalStatus ? (
                  <button onClick={handleDesconectarGoogle} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:10,padding:"9px 18px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                    🔌 Desconectar
                  </button>
                ) : (
                  <button onClick={conectarGoogleCalendar} style={{background:"linear-gradient(135deg,#1d4ed8,#0891b2)",color:"#fff",border:"none",borderRadius:10,padding:"9px 18px",fontSize:12,fontWeight:700,cursor:"pointer"}}>
                    Conectar Google Calendar →
                  </button>
                )}
              </div>

              {/* Gerenciador de horários — só aparece quando conectado */}
              {googleCalStatus&&(
                <div style={{borderTop:"1px solid #f1f5f9",paddingTop:16}}>
                  <div style={{fontSize:12,fontWeight:700,color:"#374151",marginBottom:12}}>
                    ⏰ Horários dos alarmes (horário de Brasília)
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
                    {alarmeHorarios.map((h,i)=>(
                      <div key={i} style={{display:"flex",alignItems:"center",gap:6,background:"#f1f5f9",borderRadius:10,padding:"6px 10px"}}>
                        <input
                          type="time"
                          value={h}
                          onChange={e=>{
                            const novo = [...alarmeHorarios];
                            novo[i] = e.target.value;
                            setAlarmeHorarios(novo.sort());
                          }}
                          style={{border:"none",background:"transparent",fontSize:13,fontWeight:600,color:"#1e293b",cursor:"pointer",outline:"none"}}
                        />
                        <button onClick={()=>{
                          if (alarmeHorarios.length <= 1) { showToast("Precisa ter pelo menos 1 horário.","#d97706"); return; }
                          setAlarmeHorarios(prev=>prev.filter((_,idx)=>idx!==i));
                        }} style={{background:"none",border:"none",color:"#94a3b8",cursor:"pointer",fontSize:14,fontWeight:700,padding:"0 2px",lineHeight:1}}>✕</button>
                      </div>
                    ))}
                    <button onClick={()=>setAlarmeHorarios(prev=>[...prev,"08:00"].sort())}
                      style={{background:"none",border:"1.5px dashed #cbd5e1",borderRadius:10,padding:"6px 14px",fontSize:12,color:"#64748b",cursor:"pointer",fontWeight:600}}>
                      + Adicionar horário
                    </button>
                  </div>
                  <button
                    onClick={async()=>{
                      if (alarmeHorarios.length===0) { showToast("Adicione pelo menos 1 horário.","#d97706"); return; }
                      setSalvandoHorarios(true);
                      const res = await salvarAlarmeHorarios(alarmeHorarios);
                      setSalvandoHorarios(false);
                      if (res.error) { showToast(`❌ ${res.error}`,"#991b1b"); return; }
                      registrarAuditoria(orgId, user.id, meuNome, "alterou_horarios_alarme", `Configurou ${alarmeHorarios.length} horário(s): ${alarmeHorarios.join(", ")}`);
                      showToast(`✅ ${alarmeHorarios.length} horário(s) salvos com sucesso!`,"#059669");
                    }}
                    disabled={salvandoHorarios}
                    style={{background:salvandoHorarios?"#e2e8f0":"#059669",color:salvandoHorarios?"#94a3b8":"#fff",border:"none",borderRadius:10,padding:"9px 20px",fontSize:12,fontWeight:700,cursor:salvandoHorarios?"not-allowed":"pointer"}}>
                    {salvandoHorarios ? "Salvando..." : "💾 Salvar horários"}
                  </button>
                  <div style={{fontSize:11,color:"#94a3b8",marginTop:8}}>
                    Os alarmes só são criados quando há pedidos vencendo que não estão marcados como Feito.
                  </div>
                </div>
              )}
            </div>

            {/* Histórico de Atividades / Auditoria */}
            <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden",marginBottom:18}}>
              <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",fontSize:14,fontWeight:700,color:"#1e293b"}}>📜 Histórico de Atividades</div>
              <div style={{maxHeight:320,overflowY:"auto",padding:"6px 0"}}>
                {auditoria.length===0 && <div style={{textAlign:"center",padding:24,color:"#94a3b8",fontSize:12}}>Nenhuma atividade registrada ainda.</div>}
                {auditoria.map(a=>{
                  const icones = {
                    zerou_sistema: "🗑️", criou_funcionario: "👤➕", removeu_funcionario: "👤➖", alterou_permissao: "🔐",
                    conectou_google_calendar: "📅", desconectou_google_calendar: "🔌", alterou_horarios_alarme: "⏰",
                  };
                  const dt = new Date(a.criado_em);
                  const dataFmt = dt.toLocaleDateString("pt-BR") + " " + dt.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"});
                  return (
                    <div key={a.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"8px 18px",borderBottom:"1px solid #f8fafc"}}>
                      <span style={{fontSize:16,flexShrink:0}}>{icones[a.acao]||"📌"}</span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:12,color:"#1e293b"}}><strong>{a.nome_usuario||"Usuário"}</strong> {a.detalhes}</div>
                        <div style={{fontSize:10,color:"#94a3b8",marginTop:1}}>{dataFmt}</div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Lista de membros com permissões */}
            <div style={{background:"#fff",borderRadius:18,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden"}}>
              <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",fontSize:14,fontWeight:700,color:"#1e293b"}}>⚙️ Permissões</div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{background:"#f8fafc"}}>
                    {["Funcionário","Cor","Financeiro","Zerar Sistema","Carregar Planilha","Editar Status",""].map(h=><th key={h} style={TH}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {membrosEquipe.map((m,i)=>(
                      <tr key={m.id} style={{background:i%2===0?"#fff":"#f8fafc",borderBottom:"1px solid #f1f5f9"}}>
                        <td style={TD}><strong>{m.nome}</strong><div style={{fontSize:10,color:"#94a3b8"}}>{m.email}</div></td>
                        <td style={{...TD,textAlign:"center"}}><div style={{width:18,height:18,borderRadius:"50%",background:m.cor,margin:"0 auto"}}/></td>
                        {["pode_ver_financeiro","pode_zerar_sistema","pode_carregar_planilha","pode_editar_status"].map(perm=>(
                          <td key={perm} style={{...TD,textAlign:"center"}}>
                            {m.is_admin ? <span style={{color:"#94a3b8"}}>—</span> : (
                              <input type="checkbox" checked={!!m[perm]} disabled={m.is_admin}
                                onChange={async e=>{
                                  const novoValor = e.target.checked;
                                  setMembrosEquipe(prev=>prev.map(x=>x.id===m.id?{...x,[perm]:novoValor}:x));
                                  await atualizarPermissoesMembro(m.id, {[perm]: novoValor});
                                  registrarAuditoria(orgId, user.id, meuNome, "alterou_permissao", `${novoValor?"Concedeu":"Removeu"} "${perm}" para ${m.nome}.`);
                                  showToast("✅ Permissão atualizada!","#059669");
                                }}
                                style={{width:16,height:16,cursor:"pointer"}} />
                            )}
                          </td>
                        ))}
                        <td style={TD}>
                          {!m.is_admin && (
                            <div style={{display:"flex",gap:6}}>
                              <button onClick={async()=>{
                                const { error } = await resetPassword(m.email);
                                if (error) showToast("❌ Não foi possível enviar o e-mail.","#991b1b");
                                else showToast(`📧 E-mail de redefinição enviado para ${m.email}`,"#1d4ed8");
                              }} style={{background:"#eff6ff",color:"#1d4ed8",border:"none",borderRadius:8,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
                                🔑 Redefinir senha
                              </button>
                              <button onClick={async()=>{
                                if(!window.confirm(`Remover ${m.nome} permanentemente? O login (${m.email}) será apagado de vez e o e-mail poderá ser reutilizado depois.`)) return;
                                const res = await removerFuncionarioCompleto(m.user_id);
                                if (res.error) { showToast(`❌ ${res.error}`,"#991b1b"); return; }
                                registrarAuditoria(orgId, user.id, meuNome, "removeu_funcionario", `Removeu "${m.nome}" (${m.email}) permanentemente.`);
                                setMembrosEquipe(prev=>prev.filter(x=>x.id!==m.id));
                                showToast("🗑️ Funcionário removido permanentemente.","#991b1b");
                              }} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:8,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
                                Remover
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          );
        })()}
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
      {/* ── Floating copy button — follows scroll ── */}
      {selected.size>0&&(
        <div style={{
          position:"fixed", bottom:28, left:"50%", transform:"translateX(-50%)",
          zIndex:8888, display:"flex", gap:10, alignItems:"center",
          background:"#1e293b", borderRadius:20, padding:"10px 20px",
          boxShadow:"0 8px 32px rgba(0,0,0,0.3)",
          animation:"slideUp 0.2s ease",
        }}>
          <style>{`@keyframes slideUp{from{opacity:0;transform:translateX(-50%) translateY(10px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}`}</style>
          <span style={{color:"#94a3b8",fontSize:12}}>{selected.size} pedido{selected.size!==1?"s":""} selecionado{selected.size!==1?"s":""}</span>
          <div style={{width:1,height:20,background:"#334155"}}/>
          <button onClick={copyIds} style={{
            background:copied?"#059669":"#1d4ed8",color:"#fff",border:"none",
            borderRadius:12,padding:"7px 16px",fontSize:12,fontWeight:700,cursor:"pointer",
            display:"flex",alignItems:"center",gap:5,transition:"background 0.2s",
          }}>
            {copied?"✓ IDs Copiados!":"📋 Copiar IDs"}
          </button>
          <button onClick={copyFullData} style={{
            background:copiedFull?"#059669":"#7c3aed",color:"#fff",border:"none",
            borderRadius:12,padding:"7px 16px",fontSize:12,fontWeight:700,cursor:"pointer",
            display:"flex",alignItems:"center",gap:5,transition:"background 0.2s",
          }}>
            {copiedFull?"✓ Copiado!":"📄 Copiar Completo"}
          </button>
          {selected.size>1&&(
            <button onClick={markAllFeito} style={{
              background:"#059669",color:"#fff",border:"none",
              borderRadius:12,padding:"7px 16px",fontSize:12,fontWeight:700,cursor:"pointer",
            }}>
              ✅ Marcar {selected.size} como Feito
            </button>
          )}
          <button onClick={()=>setSelected(new Set())} style={{
            background:"none",color:"#94a3b8",border:"1px solid #334155",
            borderRadius:12,padding:"7px 12px",fontSize:12,cursor:"pointer",
          }}>✕</button>
        </div>
      )}

      {/* ── Novo Funcionário Modal ── */}
      {showNovoFuncionario&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.25)",maxHeight:"90vh",overflowY:"auto"}}>
            <div style={{fontSize:28,textAlign:"center",marginBottom:6}}>👤</div>
            <h3 style={{margin:"0 0 4px",fontSize:16,fontWeight:700,color:"#1e293b",textAlign:"center"}}>Novo Funcionário</h3>
            <p style={{margin:"0 0 18px",fontSize:12,color:"#64748b",textAlign:"center"}}>Crie o login e defina as permissões dele.</p>

            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              <div>
                <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Nome</label>
                <input value={novoFuncNome} onChange={e=>setNovoFuncNome(e.target.value)} placeholder="Ex: João Silva"
                  style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:9,fontSize:13,outline:"none",boxSizing:"border-box"}} />
              </div>
              <div>
                <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>E-mail (login)</label>
                <input type="email" value={novoFuncEmail} onChange={e=>setNovoFuncEmail(e.target.value)} placeholder="joao@email.com"
                  style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:9,fontSize:13,outline:"none",boxSizing:"border-box"}} />
              </div>
              <div>
                <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Senha (mín. 6 caracteres)</label>
                <input type="text" value={novoFuncSenha} onChange={e=>setNovoFuncSenha(e.target.value)} placeholder="Crie uma senha"
                  style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:9,fontSize:13,outline:"none",boxSizing:"border-box"}} />
              </div>
              <div>
                <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Cor de identificação</label>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  {["#ef4444","#f59e0b","#22c55e","#3b82f6","#7c3aed","#ec4899","#06b6d4","#84cc16"].map(c=>(
                    <div key={c} onClick={()=>setNovoFuncCor(c)} style={{
                      width:28,height:28,borderRadius:"50%",background:c,cursor:"pointer",
                      border: novoFuncCor===c ? "3px solid #1e293b" : "3px solid transparent",
                      transition:"border 0.15s",
                    }}/>
                  ))}
                </div>
              </div>

              <div style={{height:1,background:"#f1f5f9",margin:"4px 0"}}/>
              <div style={{fontSize:11,fontWeight:700,color:"#374151"}}>Permissões</div>
              {[
                {key:"pode_ver_financeiro",    label:"Pode ver o Financeiro"},
                {key:"pode_zerar_sistema",     label:"Pode zerar o sistema"},
                {key:"pode_carregar_planilha", label:"Pode carregar planilhas"},
                {key:"pode_editar_status",     label:"Pode marcar Feito/Revisão"},
              ].map(p=>(
                <label key={p.key} style={{display:"flex",alignItems:"center",gap:8,fontSize:13,color:"#374151",cursor:"pointer"}}>
                  <input type="checkbox" checked={novoFuncPerms[p.key]}
                    onChange={e=>setNovoFuncPerms(prev=>({...prev,[p.key]:e.target.checked}))}
                    style={{width:16,height:16,cursor:"pointer"}} />
                  {p.label}
                </label>
              ))}
            </div>

            <div style={{display:"flex",gap:10,marginTop:22}}>
              <button onClick={()=>setShowNovoFuncionario(false)} style={{flex:1,padding:"11px",border:"1px solid #e2e8f0",borderRadius:12,background:"#fff",fontSize:13,cursor:"pointer",color:"#374151"}}>Cancelar</button>
              <button onClick={handleCriarFuncionario} disabled={criandoFuncionario} style={{flex:2,padding:"11px",background:criandoFuncionario?"#e2e8f0":"linear-gradient(135deg,#1d4ed8,#7c3aed)",color:criandoFuncionario?"#94a3b8":"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:700,cursor:criandoFuncionario?"not-allowed":"pointer"}}>
                {criandoFuncionario ? "Criando..." : "Criar Funcionário →"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Config Lojas Modal ── */}
      {showConfigLojas&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
          <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
            <div style={{fontSize:32,textAlign:"center",marginBottom:8}}>🏪</div>
            <h3 style={{margin:"0 0 6px",fontSize:17,fontWeight:700,color:"#1e293b",textAlign:"center"}}>Configurar Suas Lojas</h3>
            <p style={{margin:"0 0 24px",fontSize:13,color:"#64748b",textAlign:"center"}}>Defina o nome das suas duas lojas. Isso ficará salvo na sua conta.</p>
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Nome da Loja 1</label>
                <input value={configLoja1} onChange={e=>setConfigLoja1(e.target.value)} placeholder="Ex: Gran Shop"
                  style={{width:"100%",padding:"11px 14px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:14,outline:"none",boxSizing:"border-box"}} />
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Nome da Loja 2</label>
                <input value={configLoja2} onChange={e=>setConfigLoja2(e.target.value)} placeholder="Ex: Aishael Mix"
                  onKeyDown={e=>e.key==="Enter"&&handleSaveConfig()}
                  style={{width:"100%",padding:"11px 14px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:14,outline:"none",boxSizing:"border-box"}} />
              </div>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              {lojas[0]!=="Loja 1"&&<button onClick={()=>setShowConfigLojas(false)} style={{flex:1,padding:"11px",border:"1px solid #e2e8f0",borderRadius:12,background:"#fff",fontSize:13,cursor:"pointer",color:"#374151"}}>Cancelar</button>}
              <button onClick={handleSaveConfig} style={{flex:2,padding:"11px",background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:700,cursor:"pointer"}}>Salvar →</button>
            </div>
          </div>
        </div>
      )}

      {showFinanceGate&&<FinanceGate onUnlock={()=>{setFinanceUnlocked(true);setActiveTab("financeiro");setShowFinanceGate(false);}} />}
      {revisaoModal&&<RevisaoModal order={revisaoModal} onConfirm={nota=>{handleStatusChange(revisaoModal,"revisao",nota);setRevisaoModal(null);}} onClose={()=>setRevisaoModal(null)} onDesmarcar={order=>{setRevisaoModal(null);tentarDesmarcar(order,"","Revisão");}} />}

      {/* ── Modal confirmação desmarcar ─────────────────────────────── */}
      {confirmarDesmarcar&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:9100,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
          onClick={e=>e.target===e.currentTarget&&setConfirmarDesmarcar(null)}>
          <div style={{background:"#fff",borderRadius:20,padding:"32px 28px",maxWidth:380,width:"100%",boxShadow:"0 20px 60px rgba(0,0,0,0.2)",textAlign:"center"}}>
            <div style={{fontSize:48,marginBottom:12}}>
              {confirmarDesmarcar.label==="Feito"?"✅":confirmarDesmarcar.label==="Vazio"?"📦":"📋"}
            </div>
            <div style={{fontSize:17,fontWeight:700,color:"#1e293b",marginBottom:8}}>
              Desmarcar como {confirmarDesmarcar.label}?
            </div>
            <div style={{fontSize:13,color:"#64748b",marginBottom:6}}>
              Pedido <strong style={{color:"#1d4ed8"}}>{confirmarDesmarcar.order.idPedido}</strong>
              {confirmarDesmarcar.order.destinatario&&<span> · {confirmarDesmarcar.order.destinatario}</span>}
            </div>
            <div style={{fontSize:12,color:"#94a3b8",marginBottom:28}}>
              O pedido voltará para o estado <strong>pendente</strong>.
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"center"}}>
              <button
                onClick={()=>setConfirmarDesmarcar(null)}
                style={{flex:1,padding:"10px 0",borderRadius:12,border:"1px solid #e2e8f0",background:"#f8fafc",color:"#64748b",fontSize:14,fontWeight:600,cursor:"pointer"}}>
                Cancelar
              </button>
              <button
                onClick={()=>{
                  handleStatusChange(confirmarDesmarcar.order, confirmarDesmarcar.novoSt);
                  setConfirmarDesmarcar(null);
                }}
                style={{flex:1,padding:"10px 0",borderRadius:12,border:"none",background:"#ef4444",color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"}}>
                Sim, desmarcar
              </button>
            </div>
          </div>
        </div>
      )}
      {toast&&<Toast msg={toast.msg} color={toast.color} />}
    </div>
  );
}
