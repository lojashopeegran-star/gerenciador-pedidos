import { useState, useRef, useEffect, lazy, Suspense } from "react";
// XLSX carregado sob demanda para não bloquear o primeiro carregamento
let _XLSX = null;
const getXLSX = async () => { if (!_XLSX) _XLSX = await import("xlsx"); return _XLSX; };
import { supabase, fetchPedidos, upsertPedidos, updatePedidoStatus, fetchDevolucoes, upsertDevolucoes, fetchFaturamento, upsertFaturamento, deleteAllPedidos, restaurarPedidos, contarPedidosArquivados, deleteAllDevolucoes, deleteAllFaturamento, fetchConfig, saveConfig, fetchMembro, fetchOrganizacao, fetchMembrosDaOrganizacao, criarFuncionario, atualizarPermissoesMembro, removerMembro, fetchProdutividade, resetPassword, removerFuncionarioCompleto, registrarAuditoria, fetchAuditoria, fetchGoogleCalendarStatus, desconectarGoogleCalendar, fetchAlarmeHorarios, salvarAlarmeHorarios, fetchEstoqueProdutos, salvarEstoqueProduto, excluirEstoqueProduto, registrarMovimentoEstoque, fetchMovimentosEstoque, fetchProdutos, salvarProduto, excluirProduto, salvarMateriaPrima, excluirMateriaPrima, fetchMateriasPrimas, salvarMateriaPrimaLib, excluirMateriaPrimaLib, salvarProdutoCompleto, fetchFornecedores, salvarFornecedor, excluirFornecedor, fetchEstoquePlusConfig, salvarEstoquePlusPin, atualizarEstoqueProdutoPlus, fetchMovimentosComPreco, fetchCustosFixos, salvarCustoFixo, excluirCustoFixo, fetchPagamentosCustos, marcarCustoPago, fetchSaidasMes, fetchConsumoRecente, fetchProjecao, salvarProjecao, fetchKits, salvarKit, excluirKit, fetchPedidosMes, fetchPedidosUpseller, fetchVariacoesUpseller, fetchVinculos, salvarVinculo, fetchBaixaResumo, ativarBaixa, processarBaixa, registrarOperacao, fetchOperacoes, desfazerOperacoes, pick, snapProduto, COLS_ESTOQUE_PROD, COLS_FORNECEDOR, COLS_CUSTO, COLS_MP, COLS_MOV, lerNotaFiscal, notaJaLancada, chamarClaudeFinanceiro, fetchChatFinanceiro, salvarMensagemChat, limparChatFinanceiro } from "./supabase.js";

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
  // XLSX já foi carregado antes de chamar parseSheet (via DropZone)
  const raw = _XLSX.utils.sheet_to_json(sheet, { defval: "" });
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
  const handle=f=>{if(!f)return;const r=new FileReader();r.onload=async e=>{const XL=await getXLSX();const wb=XL.read(e.target.result,{type:"binary"});onFile(wb.Sheets[wb.SheetNames[0]],f.name);};r.readAsBinaryString(f);};
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

function FinanceGate({onUnlock, pinHash, onDefinirPin}) {
  const [pin,setPin]=useState(""); const [err,setErr]=useState("");
  const check=async()=>{
    if (!pinHash) return;
    const hash = await crypto.subtle.digest("SHA-256",new TextEncoder().encode(pin))
      .then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
    if(hash===pinHash){onUnlock();}else{setErr("PIN incorreto.");setPin("");}
  };
  return (
    <div onClick={()=>{}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:380,boxShadow:"0 20px 60px rgba(0,0,0,0.25)",textAlign:"center",position:"relative"}}>
        <div style={{fontSize:40,marginBottom:10}}>🔐</div>
        <h3 style={{margin:"0 0 6px",fontSize:16,fontWeight:700,color:"#1e293b"}}>Área Restrita</h3>
        <p style={{margin:"0 0 20px",fontSize:13,color:"#64748b"}}>Digite o PIN do Estoque+ para acessar</p>
        {!pinHash ? (
          <div>
            <p style={{fontSize:12,color:"#f59e0b",marginBottom:12}}>Configure o PIN primeiro na aba <strong>Estoque+</strong>.</p>
            <button onClick={onDefinirPin} style={{width:"100%",padding:"11px",background:"#059669",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:700,cursor:"pointer"}}>Definir PIN agora</button>
          </div>
        ) : (
          <>
            <input type="password" inputMode="numeric" maxLength={8} value={pin} onChange={e=>{setPin(e.target.value);setErr("");}} onKeyDown={e=>e.key==="Enter"&&check()} placeholder="••••" autoFocus
              style={{width:"100%",padding:"12px 14px",border:`2px solid ${err?"#ef4444":"#e2e8f0"}`,borderRadius:12,fontSize:22,outline:"none",boxSizing:"border-box",textAlign:"center",letterSpacing:8,marginBottom:err?6:16}} />
            {err&&<p style={{margin:"0 0 12px",fontSize:12,color:"#ef4444",fontWeight:600}}>{err}</p>}
            <button onClick={check} style={{width:"100%",padding:"12px",background:"linear-gradient(135deg,#1d4ed8,#7c3aed)",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:700,cursor:"pointer"}}>Entrar →</button>
          </>
        )}
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
// ── Helpers de cálculo de produto (usados na projeção do Financeiro) ──
const _n = (s,d) => (s===""||s==null||isNaN(Number(s))) ? d : Number(s);
// ── Leitura de nota fiscal: utilitários ──
const _norm = (t) => String(t||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
const _tokens = (t) => new Set(_norm(t).split(" ").filter(w=>w.length>1));
const similaridade = (a,b) => {
  const A=_tokens(a), B=_tokens(b); if(!A.size||!B.size) return 0;
  let inter=0; A.forEach(w=>{ if(B.has(w)) inter++; });
  return inter/Math.max(A.size,B.size);
};
const melhorMatch = (texto, lista, minimo=0.5) => {
  let best=null, bs=0;
  for (const x of lista) { const sc=similaridade(texto,x.nome); if (sc>bs) { bs=sc; best=x; } }
  return bs>=minimo ? best : null;
};
const mapUnidade = (u) => {
  const x=_norm(u);
  if (["cx","caixa","cxa"].includes(x)) return "cx";
  if (["pct","pc","pac","pacote"].includes(x)) return "pct";
  if (["rl","rolo"].includes(x)) return "rolo";
  if (["m","mt","metro","metros"].includes(x)) return "m";
  if (["kg","quilo"].includes(x)) return "kg";
  if (["l","lt","litro","litros"].includes(x)) return "l";
  return "un";
};
const arquivoParaBase64 = (file) => new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(String(r.result).split(",")[1]); r.onerror=rej; r.readAsDataURL(file); });
// Reduz fotos grandes (celular) para acelerar o envio; se o navegador não abrir o formato, envia o original
const prepararImagemNota = async (file) => {
  try {
    const url = URL.createObjectURL(file);
    const img = await new Promise((res,rej)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=rej; i.src=url; });
    const max = 2000; const k = Math.min(1, max/Math.max(img.width,img.height));
    const c = document.createElement("canvas"); c.width=Math.round(img.width*k); c.height=Math.round(img.height*k);
    c.getContext("2d").drawImage(img,0,0,c.width,c.height);
    URL.revokeObjectURL(url);
    return { mime:"image/jpeg", base64:c.toDataURL("image/jpeg",0.85).split(",")[1] };
  } catch { return { mime:file.type||"image/jpeg", base64:await arquivoParaBase64(file) }; }
};
const fmtBRLn = (n) => "R$ "+Number(n||0).toFixed(2).replace(".",",");
const quandoFmt = (iso) => { const d=new Date(iso); return String(d.getDate()).padStart(2,"0")+"/"+String(d.getMonth()+1).padStart(2,"0")+" "+String(d.getHours()).padStart(2,"0")+":"+String(d.getMinutes()).padStart(2,"0"); };
const liquidoDe = (p,preco) => Number(preco||0)*(1-(_n(p.taxa_plataforma_pct,0)+_n(p.taxa_imposto_pct,0)+_n(p.taxa_adiantamento_pct,0))/100) - _n(p.taxa_fixa,0);
const custoMpDe = (v) => (v.produto_materias_primas||[]).reduce((a,l)=>{
  const m = l.materias_primas || l;
  const un = Number(m.qtd_por_unidade_compra)>0 ? Number(m.custo_unidade_compra)/Number(m.qtd_por_unidade_compra) : 0;
  return a + un*Number(l.qtd_usada_por_produto||0);
},0);

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
  const [activeTab,     setActiveTab]     = useState("estoque");
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
  const [filterProduto,     setFilterProduto]     = useState("all");
  const [filterFuncionario, setFilterFuncionario] = useState("all");
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
  // ── Estoque ────────────────────────────────────────────────────────────────
  const [estoqueProdutos,    setEstoqueProdutos]    = useState([]);
  const [estoqueMovimentos,  setEstoqueMovimentos]  = useState([]);
  const [showEstoqueForm,    setShowEstoqueForm]    = useState(false);
  const [editandoProduto,    setEditandoProduto]    = useState(null); // null = novo
  const [estoqueProdNome,    setEstoqueProdNome]    = useState("");
  const [estoqueProdUnidade, setEstoqueProdUnidade] = useState("un");
  const [estoqueProdQtd,     setEstoqueProdQtd]     = useState("");
  const [estoqueProdMin,     setEstoqueProdMin]     = useState("");
  const [salvandoProduto,    setSalvandoProduto]    = useState(false);
  const [abrirPlusAposSalvar, setAbrirPlusAposSalvar] = useState(false);
  const [dimOpcaoSelecionada, setDimOpcaoSelecionada] = useState(0);
  const [showMovForm,        setShowMovForm]        = useState(null); // produtoId ou null
  const [movTipo,            setMovTipo]            = useState("entrada");
  const [movQtd,             setMovQtd]             = useState("");
  const [movMotivo,          setMovMotivo]          = useState("");
  const [salvandoMov,        setSalvandoMov]        = useState(false);
  const [estoqueSearch,      setEstoqueSearch]      = useState("");

  // ── ESTOQUE+ (fornecedores, preços, PIN) ─────────────────────────────────
  const [estoquePlusDesbloqueado, setEstoquePlusDesbloqueado] = useState(false);
  const [estoquePlusConfig,       setEstoquePlusConfig]       = useState(null); // {pin_hash}
  const [showEstoquePlusPin,      setShowEstoquePlusPin]      = useState(false); // tela de PIN
  const [pinInput,                setPinInput]                = useState("");
  const [pinErro,                 setPinErro]                 = useState("");
  const [showDefinirPin,          setShowDefinirPin]          = useState(false); // modal configurar PIN
  const [novoPin,                 setNovoPin]                 = useState("");
  const [novoPinConfirm,          setNovoPinConfirm]          = useState("");
  const [salvandoPin,             setSalvandoPin]             = useState(false);
  const [fornecedores,            setFornecedores]            = useState([]);
  const [showFornecedorForm,      setShowFornecedorForm]      = useState(false);
  const [editandoFornecedor,      setEditandoFornecedor]      = useState(null);
  const [fornNome,                setFornNome]                = useState("");
  const [fornContato,             setFornContato]             = useState("");
  const [fornSite,                setFornSite]                = useState("");
  const [fornObs,                 setFornObs]                 = useState("");
  const [salvandoForn,            setSalvandoForn]            = useState(false);
  const [movimentosMes,           setMovimentosMes]           = useState([]);
  const [mesSelecionado,          setMesSelecionado]          = useState(() => new Date().toISOString().slice(0,7));
  // Edição de campos plus de um produto do estoque
  const [showEstoqueProdPlus,     setShowEstoqueProdPlus]     = useState(false);
  const [estoqueProdPlusId,       setEstoqueProdPlusId]       = useState(null);
  const [estoqueProdFornId,       setEstoqueProdFornId]       = useState("");
  const [estoqueProdPreco,        setEstoqueProdPreco]        = useState("");
  const [estoqueProdLink,         setEstoqueProdLink]         = useState("");
  const [estoqueProdObs,          setEstoqueProdObs]          = useState("");
  const [estoqueProdCat,          setEstoqueProdCat]          = useState("");
  const [salvandoProdPlus,        setSalvandoProdPlus]        = useState(false);
  // Fornecedor autocomplete no modal entrada estoque
  const [fornSugestoes,           setFornSugestoes]           = useState([]);

  // ── FINANCEIRO ────────────────────────────────────────────────────────────
  const [custosFixos,         setCustosFixos]         = useState([]);
  const [showCustoForm,       setShowCustoForm]       = useState(false);
  const [editandoCusto,       setEditandoCusto]       = useState(null);
  const [custoNome,           setCustoNome]           = useState("");
  const [custoCategoria,      setCustoCategoria]      = useState("outros");
  const [custoValor,          setCustoValor]          = useState("");
  const [custoRecorrente,     setCustoRecorrente]     = useState(true);
  const [custoObs,            setCustoObs]            = useState("");
  const [custoVenc,           setCustoVenc]           = useState("");
  const [custosPagos,         setCustosPagos]         = useState([]); // ids pagos no mês finMes
  const [projBase,            setProjBase]            = useState("compras"); // compras | saidas | produto
  const [finSaidas,           setFinSaidas]           = useState([]);
  const [projEstoque,         setProjEstoque]         = useState("");   // vazio = usa o gasto real do mês
  const [projOutros,          setProjOutros]          = useState("");
  const [projLucro,           setProjLucro]           = useState("");
  const [projJa,              setProjJa]              = useState("");
  const [projMix,             setProjMix]             = useState({});   // {variacaoId: "%"}
  const [salvandoCusto,       setSalvandoCusto]       = useState(false);
  const [finMes,              setFinMes]              = useState(() => { const d=new Date(); return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0"); });
  const [finMovimentos,       setFinMovimentos]       = useState([]);
  // Chat com Claude
  const [chatMsgs,            setChatMsgs]            = useState([]);
  const [chatInput,           setChatInput]           = useState("");
  const [chatCarregando,      setChatCarregando]      = useState(false);
  const [chatAnalise,         setChatAnalise]         = useState(""); // análise automática
  const [chatAnaliseLoading,  setChatAnaliseLoading]  = useState(false);
  const chatEndRef = useRef(null);
  // Vendas ERP
  const [erpVendas,           setErpVendas]           = useState(null);   // { total, quantidade, imagemUrl }
  const [erpImagemPreview,    setErpImagemPreview]    = useState(null);
  const [erpAnalisando,       setErpAnalisando]       = useState(false);
  const [erpResultado,        setErpResultado]        = useState(null);   // { totalVendas, lucroLiquido, margemMedia, ... }
  const erpFileRef = useRef(null);

  // ── PRODUTOS / CALCULADORA ────────────────────────────────────────────────
  const [produtosCalc,          setProdutosCalc]              = useState([]);
  const [materiasPrimas, setMateriasPrimas] = useState([]);
  const [showMpLib, setShowMpLib] = useState(false);
  const [prodVars, setProdVars] = useState([]);
  const [produtoAtivo,          setProdutoAtivo]          = useState(null); // produto selecionado para ver detalhes
  const [showProdutoForm,       setShowProdutoForm]       = useState(false);
  const [editandoProdutoCalc,   setEditandoProdutoCalc]   = useState(null);
  const [prodNome,              setProdNome]              = useState("");
  const [prodPreco,             setProdPreco]             = useState("");
  const [prodTaxaPlat,          setProdTaxaPlat]          = useState("20");
  const [prodTaxaImp,           setProdTaxaImp]           = useState("4");
  const [prodTaxaAdiant,        setProdTaxaAdiant]        = useState("3.5");
  const [prodTaxaFixa,          setProdTaxaFixa]          = useState("4");
  const [salvandoProdutoCalc,   setSalvandoProdutoCalc]   = useState(false);
  // Matéria-prima form
  const [showMpForm,            setShowMpForm]            = useState(false);
  const [editandoMp,            setEditandoMp]            = useState(null);
  const [mpNome,                setMpNome]                = useState("");
  const [mpCustoCompra,         setMpCustoCompra]         = useState("");
  const [mpQtdPorCompra,        setMpQtdPorCompra]        = useState("");
  const [mpQtdUsada,            setMpQtdUsada]            = useState("");
  const [mpUnidade,             setMpUnidade]             = useState("folha");
  const [mpEstoqueId,           setMpEstoqueId]           = useState("");
  const [mpRendimento,          setMpRendimento]          = useState("");
  const [salvandoMp,            setSalvandoMp]            = useState(false);
  const [mpModoDimensao,        setMpModoDimensao]        = useState(false);
  // campos modo dimensão (rolo/folha)
  const [mpRoloLargura,         setMpRoloLargura]         = useState(""); // cm
  const [mpRoloComprimento,     setMpRoloComprimento]     = useState(""); // m
  const [mpCorteL,              setMpCorteL]              = useState(""); // cm largura do corte
  const [mpCorteC,              setMpCorteC]              = useState(""); // cm comprimento do corte
  const [mpUnidadeCorte,        setMpUnidadeCorte]        = useState("folha");
  const tableRef  = useRef(null);
  const notifRef  = useRef(null);

  // ── Notificações in-app (declarados ANTES do useEffect que os usa) ──────────
  const [showNotifPanel,  setShowNotifPanel]  = useState(false);
  const [showTodosCards,  setShowTodosCards]  = useState(false);
  const [lidas,           setLidas]           = useState(() => {
    try { return new Set(JSON.parse(localStorage.getItem("notif_lidas")||"[]")); } catch { return new Set(); }
  });

  const showToast=(msg,color,ms=3500)=>{setToast({msg,color});setTimeout(()=>setToast(null),ms);};

  // ── Voz do browser ──────────────────────────────────────────────────────────
  const falarFrase = (frase) => {
    if (!frase || !frase.trim()) return;
    if (!window.speechSynthesis) return;
    window.speechSynthesis.cancel();
    const utterance = new SpeechSynthesisUtterance(frase.trim());
    utterance.lang = "pt-BR";
    utterance.rate = 1;
    utterance.pitch = 1.1;
    // Preferir voz em português se disponível
    const vozes = window.speechSynthesis.getVoices();
    const vozPT = vozes.find(v => v.lang.startsWith("pt"));
    if (vozPT) utterance.voice = vozPT;
    window.speechSynthesis.speak(utterance);
  };

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
        fetchOrganizacao(orgId).then(setOrganizacao);
        // Todos os membros precisam da lista da equipe para resolver as cores
        // de quem marcou cada pedido como feito — não só o admin.
        fetchMembrosDaOrganizacao(orgId).then(setMembrosEquipe);
        fetchEstoqueProdutos(user.id, orgId).then(setEstoqueProdutos);
        fetchProdutos(user.id, orgId).then(setProdutosCalc);
        fetchMateriasPrimas(user.id, orgId).then(setMateriasPrimas);
        fetchKits(orgId).then(setKits);
        fetchFornecedores(user.id, orgId).then(setFornecedores);
        fetchEstoquePlusConfig(user.id, orgId).then(setEstoquePlusConfig);
        fetchCustosFixos(user.id, orgId).then(setCustosFixos);
        fetchChatFinanceiro(user.id, orgId).then(setChatMsgs);
        if (m?.is_admin) {
          fetchProdutividade(orgId).then(setProdutividade);
          fetchAuditoria(orgId, true).then(setAuditoria);
          fetchGoogleCalendarStatus(orgId).then(setGoogleCalStatus);
          fetchAlarmeHorarios(orgId).then(setAlarmeHorarios);
          fetchMovimentosEstoque(user.id, orgId, null, 100).then(setEstoqueMovimentos);
        }
      }

      Promise.all([
        fetchPedidos(user.id, orgId),
        fetchDevolucoes(user.id, orgId),
        fetchFaturamento(user.id, orgId),
        fetchConfig(user.id, orgId),
        contarPedidosArquivados(user.id, orgId),
      ]).then(([p,d,f,cfg,qtdArq])=>{
          setAllPedidos(p); setDevolucoes(d); setFaturamento(f);
          setQtdArquivados(qtdArq||0);
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
  const [pedidosMes, setPedidosMes] = useState(null);
  const [kits, setKits] = useState([]);
  const [showKitForm, setShowKitForm] = useState(false);
  const [editandoKit, setEditandoKit] = useState(null);
  const [kitNome, setKitNome] = useState("");
  const [kitItens, setKitItens] = useState([{mpId:"",fator:"1"}]);
  const [salvandoKit, setSalvandoKit] = useState(false);
  const [ups, setUps] = useState({total:0,linhas:[]});
  const [upsBusca, setUpsBusca] = useState("");
  const [upsBuscaAtiva, setUpsBuscaAtiva] = useState("");
  const [upsStatus, setUpsStatus] = useState("");
  const [upsLoja, setUpsLoja] = useState("");
  const [upsVinculo, setUpsVinculo] = useState("");
  const [upsPag, setUpsPag] = useState(0);
  const [upsCarregando, setUpsCarregando] = useState(false);
  const [upsVincular, setUpsVincular] = useState(null);
  const [upsEscolha, setUpsEscolha] = useState("");
  const [upsSalvando, setUpsSalvando] = useState(false);
  const UPS_POR_PAG = 50;
  const [projSalvaEm, setProjSalvaEm] = useState(null);
  const [salvandoProj, setSalvandoProj] = useState(false);
  const [alertaContas, setAlertaContas] = useState({custos:[],pagos:[],mes:""});
  const [consumo30, setConsumo30] = useState({});
  const [showVinc, setShowVinc] = useState(false);
  const [vincLista, setVincLista] = useState([]);
  const [vincCarregando, setVincCarregando] = useState(false);
  const [vincBusca, setVincBusca] = useState("");
  const [vincSo, setVincSo] = useState("pendentes");
  const [baixaInfo, setBaixaInfo] = useState(null);
  const [baixaBusy, setBaixaBusy] = useState(false);
  const [showHistorico, setShowHistorico] = useState(false);
  const [operacoes, setOperacoes] = useState([]);
  const [carregandoOps, setCarregandoOps] = useState(false);
  const [desfazendoKey, setDesfazendoKey] = useState(null);
  const [filtroHist, setFiltroHist] = useState("todos");
  useEffect(() => {
    if (!user || !orgId) return;
    fetchPagamentosCustos(user.id, orgId, finMes).then(setCustosPagos);
    fetchSaidasMes(user.id, orgId, finMes).then(setFinSaidas);
    fetchPedidosMes(finMes).then(r=>setPedidosMes(r&&!r.error?r:null));
  }, [finMes, orgId, user?.id]);
  useEffect(() => {
    if (!user || !orgId || !isAdminEquipe) return;
    const d = new Date(); const mes = d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0");
    Promise.all([fetchCustosFixos(user.id,orgId), fetchPagamentosCustos(user.id,orgId,mes), fetchConsumoRecente(user.id,orgId,30)])
      .then(([cf,pg,cons])=>{ setAlertaContas({custos:cf||[],pagos:pg||[],mes}); setConsumo30(cons||{}); })
      .catch(()=>{});
  }, [orgId, user?.id, isAdminEquipe, custosFixos.length, custosPagos.length, estoqueProdutos.length]);
  useEffect(() => { const t = setTimeout(()=>{ setUpsBuscaAtiva(upsBusca.trim()); setUpsPag(0); }, 400); return ()=>clearTimeout(t); }, [upsBusca]);
  const carregarUps = async () => {
    if (!orgId) return;
    setUpsCarregando(true);
    const r = await fetchPedidosUpseller(orgId, {busca:upsBuscaAtiva, status:upsStatus, loja:upsLoja, vinculo:upsVinculo, limite:UPS_POR_PAG, offset:upsPag*UPS_POR_PAG});
    setUpsCarregando(false);
    if (r.error) { showToast("❌ "+r.error,"#991b1b"); return; }
    setUps(r);
  };
  useEffect(() => {
    if (activeTab !== "upseller" || !orgId || !isAdminEquipe) return;
    carregarUps();
  }, [activeTab, orgId, upsBuscaAtiva, upsStatus, upsLoja, upsVinculo, upsPag]);
  const opcoesVarTop = [];
  produtosCalc.forEach(pr=>(pr.produto_variacoes||[]).forEach(v=>{
    const temEstoque = (v.produto_materias_primas||[]).some(l=>l.materias_primas?.estoque_produto_id);
    opcoesVarTop.push({id:v.id,nome:`${pr.nome}${v.nome?" — "+v.nome:""}`,temEstoque,qtdMp:(v.produto_materias_primas||[]).length});
  }));
  const confirmarVinculoUps = async () => {
    if (!upsVincular) return;
    setUpsSalvando(true);
    const ign = upsEscolha === "__ignorar";
    const r = await salvarVinculo(orgId, upsVincular.chave, upsVincular.variacao, ign ? null : (upsEscolha || null), ign);
    setUpsSalvando(false);
    if (r.error) { showToast("❌ "+r.error,"#991b1b"); return; }
    showToast(upsEscolha?"✅ Vinculado! Todos os pedidos com essa parte foram vinculados.":"🔓 Vínculo removido.","#059669");
    setUpsVincular(null); setUpsEscolha("");
    carregarUps();
  };
  useEffect(() => {
    if (!orgId || !isAdminEquipe) return;
    let vivo = true;
    fetchProjecao(orgId, finMes).then(r=>{
      if (!vivo) return;
      const d = r?.dados || {};
      setProjBase(d.base||"compras"); setProjEstoque(d.estoque??""); setProjOutros(d.outros??""); setProjLucro(d.lucro??"");
      setProjJa(d.ja??""); setProjMix(d.mix||{});
      setProjSalvaEm(r ? {quando:r.atualizado_em, quem:r.atualizado_por} : null);
    }).catch(()=>{});
    return ()=>{ vivo = false; };
  }, [finMes, orgId, isAdminEquipe]);
  const salvarProjecaoMes = async () => {
    setSalvandoProj(true);
    const r = await salvarProjecao(orgId, finMes, {base:projBase, estoque:projEstoque, outros:projOutros, lucro:projLucro, ja:projJa, mix:projMix}, meuNome);
    setSalvandoProj(false);
    if (r.error) { showToast("❌ "+r.error,"#991b1b"); return; }
    setProjSalvaEm({quando:new Date().toISOString(), quem:meuNome});
    showToast("💾 Projeção de "+finMes+" salva!","#059669");
  };
  const exportarContador = async () => {
    const R2 = (n)=>Math.round((Number(n)||0)*100)/100;
    const [fy,fm] = finMes.split("-").map(Number);
    const venc = (c)=>{ if(!c.data_vencimento) return ""; const b=String(c.data_vencimento).slice(0,10).split("-").map(Number); const dia=c.recorrente!==false?Math.min(b[2],new Date(fy,fm,0).getDate()):b[2]; const d=c.recorrente!==false?new Date(fy,fm-1,dia):new Date(b[0],b[1]-1,b[2]); return d.toLocaleDateString("pt-BR"); };
    const compras = finMovimentos.map(m=>({Data:new Date(m.criado_em).toLocaleDateString("pt-BR"),Fornecedor:m.fornecedores?.nome||"",Material:m.estoque_produtos?.nome||"",Quantidade:Number(m.quantidade)||0,Unidade:m.estoque_produtos?.unidade||"","Preço unitário (R$)":R2(m.preco_unitario),"Total (R$)":R2((m.preco_unitario||0)*m.quantidade),"Nota fiscal":m.nota_fiscal||""}));
    const saidas = finSaidas.map(m=>({Data:new Date(m.criado_em).toLocaleDateString("pt-BR"),Material:m.estoque_produtos?.nome||"",Quantidade:Number(m.quantidade)||0,Unidade:m.estoque_produtos?.unidade||"","Custo (R$)":R2((Number(m.quantidade)||0)*(Number(m.estoque_produtos?.preco_unitario)||0)),Motivo:m.motivo||""}));
    const custos = custosFixos.map(c=>({Custo:c.nome,Categoria:c.categoria||"","Valor (R$)":R2(c.valor),Vencimento:venc(c),Recorrente:c.recorrente!==false?"Sim":"Não",Situação:custosPagos.includes(c.id)?"Pago":"A pagar"}));
    const totCompras = compras.reduce((a,x)=>a+x["Total (R$)"],0);
    const totSaidas = saidas.reduce((a,x)=>a+x["Custo (R$)"],0);
    const totCustos = custos.reduce((a,x)=>a+x["Valor (R$)"],0);
    const totPagos = custos.filter(x=>x.Situação==="Pago").reduce((a,x)=>a+x["Valor (R$)"],0);
    const resumo = [
      {Item:"Mês",Valor:finMes},
      {Item:"Pedidos no mês (UpSeller, sem cancelados)",Valor:pedidosMes?pedidosMes.pedidos:"—"},
      {Item:"Pedidos cancelados",Valor:pedidosMes?pedidosMes.cancelados:"—"},
      {Item:"Compras de material (R$)",Valor:R2(totCompras)},
      {Item:"Material usado no mês (R$)",Valor:R2(totSaidas)},
      {Item:"Custos fixos do mês (R$)",Valor:R2(totCustos)},
      {Item:"Custos fixos pagos (R$)",Valor:R2(totPagos)},
      {Item:"Custos fixos a pagar (R$)",Valor:R2(totCustos-totPagos)},
    ];
    const XL = await getXLSX();
    const wb = XL.utils.book_new();
    [["Resumo",resumo],["Compras",compras],["Material usado",saidas],["Custos fixos",custos]].forEach(([nome,rows])=>{
      XL.utils.book_append_sheet(wb, XL.utils.json_to_sheet(rows.length?rows:[{"(vazio)":"Nenhum registro"}]), nome);
    });
    XL.writeFile(wb, `contador_${finMes}.xlsx`);
    showToast("📊 Planilha do contador gerada!","#059669");
  };
  const minhaCor              = membro?.cor || "#059669";
  const meuNome                = membro?.nome || user?.email?.split("@")[0] || "Você";
  const logOp = (op) => registrarOperacao(orgId, user?.id, meuNome, op);

  // ── Histórico de operações (desfazer) ───────────────────────────────────────
  const recarregarTudo = async () => {
    if (!user) return;
    const [ep, em, fo, cf, mp, pr, fm, fs, pg, mm] = await Promise.all([
      fetchEstoqueProdutos(user.id, orgId), fetchMovimentosEstoque(user.id, orgId, null, 100), fetchFornecedores(user.id, orgId),
      fetchCustosFixos(user.id, orgId), fetchMateriasPrimas(user.id, orgId), fetchProdutos(user.id, orgId),
      fetchMovimentosComPreco(user.id, orgId, finMes), fetchSaidasMes(user.id, orgId, finMes), fetchPagamentosCustos(user.id, orgId, finMes),
      fetchMovimentosComPreco(user.id, orgId, mesSelecionado),
    ]);
    setEstoqueProdutos(ep); setEstoqueMovimentos(em); setFornecedores(fo); setCustosFixos(cf); setMateriasPrimas(mp); setProdutosCalc(pr);
    setFinMovimentos(fm); setFinSaidas(fs); setCustosPagos(pg); setMovimentosMes(mm);
    setProdutoAtivo(a=>a ? (pr.find(p=>p.id===a.id)||null) : a);
  };
  const carregarBaixaInfo = async () => {
    if (!orgId) return;
    const r = await fetchBaixaResumo(orgId);
    setBaixaInfo(r && !r.error ? r : {ativo:false});
  };
  const abrirVinc = async () => {
    setShowVinc(true); setVincCarregando(true);
    const [lista] = await Promise.all([fetchVariacoesUpseller(orgId), carregarBaixaInfo()]);
    if (lista?.error) { showToast("❌ "+lista.error,"#991b1b"); setVincLista([]); }
    else setVincLista(lista||[]);
    setVincCarregando(false);
  };
  const escolherVinculo = async (it, valor) => {
    const ign = valor === "__ignorar";
    const variacaoId = ign ? null : valor;
    const r = await salvarVinculo(orgId, it.chave, it.variacao, variacaoId, ign);
    if (r.error) { showToast("❌ "+r.error,"#991b1b"); return; }
    setVincLista(l=>l.map(x=>x.chave===it.chave?{...x,variacao_id:variacaoId||null,ignorar:ign}:x));
    carregarBaixaInfo();
  };
  const alternarBaixa = async () => {
    const novo = !baixaInfo?.ativo;
    if (novo && !window.confirm("Ativar a baixa automática? A partir de agora, cada pedido novo do UpSeller em “Para imprimir” ou “Para retirada”, com variação vinculada, vai descontar os materiais do estoque. Pedidos anteriores NÃO serão baixados. Você pode desfazer pelo 🕘 Histórico.")) return;
    setBaixaBusy(true);
    const r = await ativarBaixa(orgId, novo);
    setBaixaBusy(false);
    if (r.error) { showToast("❌ "+r.error,"#991b1b"); return; }
    showToast(novo?"✅ Baixa automática ativada!":"⏸️ Baixa automática pausada.","#059669");
    carregarBaixaInfo();
  };
  const rodarBaixa = async () => {
    setBaixaBusy(true);
    const r = await processarBaixa(orgId);
    setBaixaBusy(false);
    if (r?.error) { showToast("❌ "+r.error,"#991b1b"); return; }
    if (!r?.ativo) { showToast("A baixa automática está pausada.","#64748b"); return; }
    if (!r.pedidos) showToast("Nenhum pedido novo para baixar.","#64748b");
    else { showToast(`✅ ${r.pedidos} pedido(s) baixado(s) · ${r.materiais} material(is) atualizados.`,"#059669"); await recarregarTudo(); }
    carregarBaixaInfo();
  };
  const abrirHistorico = async () => {
    setShowHistorico(true); setCarregandoOps(true);
    setOperacoes(await fetchOperacoes(user.id, orgId));
    setCarregandoOps(false);
  };
  const desfazerGrupo = async (g) => {
    const pendentes = g.ops.filter(o=>!o.desfeito_em);
    if (!pendentes.length) return;
    const msg = g.ops.length>1 ? `Desfazer TUDO desta operação (${pendentes.length} ações)?\n\n${g.titulo}` : `Desfazer esta operação?\n\n${g.titulo}`;
    if (!window.confirm(msg)) return;
    setDesfazendoKey(g.key);
    const r = await desfazerOperacoes(pendentes, { userId: user.id, orgId }, meuNome);
    setDesfazendoKey(null);
    await recarregarTudo();
    setOperacoes(await fetchOperacoes(user.id, orgId));
    if (r.error) showToast("❌ "+r.error+(r.feitas?` (${r.feitas} ação(ões) já desfeita(s))`:""),"#991b1b",6000);
    else showToast("↩️ Operação desfeita!","#059669");
  };

  // ── Leitura de nota fiscal ──────────────────────────────────────────────────
  const [showNota, setShowNota] = useState(false);
  const [notaFase, setNotaFase] = useState("upload"); // upload | lendo | revisar
  const [notaPreview, setNotaPreview] = useState(null);
  const [notaErro, setNotaErro] = useState("");
  const [notaCab, setNotaCab] = useState({fornecedor:"",cnpj:"",numero:"",data:"",total:""});
  const [notaFornId, setNotaFornId] = useState("");
  const [notaItens, setNotaItens] = useState([]);
  const [notaDuplicada, setNotaDuplicada] = useState(false);
  const [salvandoNota, setSalvandoNota] = useState(false);
  const abrirNota = () => { setNotaFase("upload"); setNotaPreview(null); setNotaErro(""); setNotaItens([]); setNotaCab({fornecedor:"",cnpj:"",numero:"",data:"",total:""}); setNotaFornId(""); setNotaDuplicada(false); setShowNota(true); };
  const lerArquivoNota = async (file) => {
    if (!file) return;
    setNotaErro("");
    const ehPdf = file.type==="application/pdf" || /\.pdf$/i.test(file.name);
    const ehImg = file.type.startsWith("image/");
    if (!ehPdf && !ehImg) { setNotaErro("Envie uma foto (JPG/PNG) ou um PDF."); return; }
    if (file.size > 15*1024*1024) { setNotaErro("Arquivo muito grande (máx. 15 MB)."); return; }
    setNotaFase("lendo");
    setNotaPreview({ url: URL.createObjectURL(file), tipo: ehPdf?"pdf":"img", nome: file.name });
    let mime, base64;
    try {
      if (ehPdf) { mime="application/pdf"; base64=await arquivoParaBase64(file); }
      else { const r=await prepararImagemNota(file); mime=r.mime; base64=r.base64; }
    } catch { setNotaErro("Não consegui abrir o arquivo."); setNotaFase("upload"); return; }
    const res = await lerNotaFiscal(mime, base64);
    if (res.error) { setNotaErro(res.error); setNotaFase("upload"); return; }
    const d = res.dados || {};
    const itens = (d.itens||[]).map((it,i)=>{
      const qtd = Number(it.quantidade)||0;
      let preco = Number(it.valor_unitario)||0;
      if (!preco && qtd) preco = (Number(it.valor_total)||0)/qtd;
      const m = melhorMatch(it.descricao, estoqueProdutos);
      return { key:i, descricao:String(it.descricao||""), nome:String(it.descricao||"").trim(), unNota:String(it.unidade||""), un:mapUnidade(it.unidade), qtd:qtd?String(qtd):"", preco:preco?preco.toFixed(2):"", destino:m?m.id:"novo" };
    });
    const nomeF = String(d.fornecedor||"").trim();
    const fm = nomeF ? melhorMatch(nomeF, fornecedores, 0.5) : null;
    const numero = String(d.numero_nota||"").trim();
    setNotaItens(itens);
    setNotaCab({ fornecedor:nomeF, cnpj:String(d.cnpj||""), numero, data:/^\d{4}-\d{2}-\d{2}$/.test(d.data_emissao||"")?d.data_emissao:"", total:d.valor_total?String(d.valor_total):"" });
    setNotaFornId(fm?.id||"");
    setNotaDuplicada(await notaJaLancada(user.id, orgId, numero));
    setNotaFase("revisar");
  };
  const setNotaItem = (key, patch) => setNotaItens(l=>l.map(i=>i.key===key?{...i,...patch}:i));
  const confirmarNota = async () => {
    const validos = notaItens.filter(i=>i.destino!=="ignorar" && Number(i.qtd)>0);
    if (!validos.length) { showToast("❌ Nenhum item para lançar.","#991b1b"); return; }
    if (validos.some(i=>i.destino==="novo" && !i.nome.trim())) { showToast("❌ Dê um nome aos materiais novos.","#991b1b"); return; }
    if (notaDuplicada && !window.confirm("Uma nota com esse número já foi lançada antes. Lançar de novo mesmo assim?")) return;
    setSalvandoNota(true);
    const lote = crypto.randomUUID();
    const numero = notaCab.numero.trim();
    const totalNota = validos.reduce((a,i)=>a+Number(i.qtd)*(Number(i.preco)||0),0);
    await logOp({tipo:"nota_fiscal",acao:"nota",descricao:`📄 Nota fiscal${numero?" nº "+numero:""} — ${notaCab.fornecedor.trim()||"sem fornecedor"} (${validos.length} item${validos.length>1?"ns":""}, ${fmtBRLn(totalNota)})`,tabela:"nota_fiscal",lote_id:lote});
    const parar = async (msg) => {
      setSalvandoNota(false);
      await recarregarTudo();
      showToast("❌ "+msg+" O que já foi lançado pode ser desfeito no 🕘 Histórico.","#991b1b",8000);
    };
    let fornId = notaFornId || null;
    if (!fornId && notaCab.fornecedor.trim()) {
      const r = await salvarFornecedor(user.id, orgId, { nome:notaCab.fornecedor.trim(), observacoes: notaCab.cnpj?("CNPJ "+notaCab.cnpj):"" });
      if (r.error) { await parar("Erro ao criar fornecedor: "+r.error); return; }
      fornId = r.data.id;
      await logOp({tipo:"fornecedores",acao:"criar",descricao:"Fornecedor criado: "+notaCab.fornecedor.trim(),tabela:"fornecedores",registro_id:fornId,lote_id:lote});
    }
    const quando = notaCab.data ? `${notaCab.data}T12:00:00-03:00` : undefined;
    let atual = estoqueProdutos;
    for (const it of validos) {
      let prodId = it.destino;
      let prodAntes = atual.find(x=>x.id===prodId);
      if (prodId==="novo") {
        const r = await salvarEstoqueProduto(user.id, orgId, { nome:it.nome.trim(), unidade:it.un, quantidade_atual:0, quantidade_minima:0 });
        if (r.error || !r.data?.[0]) { await parar("Erro ao criar material "+it.nome+": "+(r.error||"")); return; }
        prodId = r.data[0].id; prodAntes = null;
        await logOp({tipo:"estoque",acao:"criar",descricao:"Material criado: "+it.nome.trim(),tabela:"estoque_produtos",registro_id:prodId,lote_id:lote});
      }
      const preco = Number(it.preco)||0;
      const mv = await registrarMovimentoEstoque(user.id, orgId, meuNome, { produto_id:prodId, tipo:"entrada", quantidade:Number(it.qtd), motivo:"Nota fiscal"+(numero?" nº "+numero:""), preco_unitario:preco||null, fornecedor_id:fornId, nota_fiscal:numero||null, criado_em:quando });
      if (mv.error) { await parar("Erro ao lançar entrada de "+(it.nome||it.descricao)+": "+mv.error); return; }
      await logOp({tipo:"estoque",acao:"criar",descricao:`Entrada: ${Number(it.qtd)} ${prodAntes?.unidade||it.un} de ${prodAntes?.nome||it.nome} (nota)`,tabela:"estoque_movimentos",registro_id:mv.movimento.id,depois:{...pick(mv.movimento,COLS_MOV),quantidade:Number(mv.movimento.quantidade)},lote_id:lote});
      if (preco>0 || fornId) {
        const antesP = { preco_unitario: prodAntes?.preco_unitario ?? null, fornecedor_id: prodAntes?.fornecedor_id ?? null };
        const campos = { preco_unitario: preco>0?preco:antesP.preco_unitario, fornecedor_id: fornId||antesP.fornecedor_id };
        const up = await atualizarEstoqueProdutoPlus(prodId, campos);
        if (!up.error) await logOp({tipo:"estoque",acao:"editar",descricao:"Preço/fornecedor atualizado: "+(prodAntes?.nome||it.nome),tabela:"estoque_produtos",registro_id:prodId,antes:antesP,depois:campos,lote_id:lote});
      }
    }
    setSalvandoNota(false);
    setShowNota(false);
    await recarregarTudo();
    showToast("✅ Nota lançada! Se algo estiver errado, desfaça no 🕘 Histórico.","#059669",6000);
  };

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
  const [showConfirmClear,   setShowConfirmClear]   = useState(false);
  const [qtdArquivados,      setQtdArquivados]      = useState(0);
  const [showConfirmRestore, setShowConfirmRestore] = useState(false);


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
      registrarAuditoria(orgId, user.id, meuNome, "zerou_sistema", `Arquivou ${totalAntes} pedido(s) (soft delete - expira em 7 dias).`);
      const qtd = await contarPedidosArquivados(user.id, orgId);
      setQtdArquivados(qtd);
    }
    setAllPedidos([]);
    setDevolucoes([]);
    setFaturamento([]);
    setUploadNames({});
    setSelected(new Set());
    setActiveTab("abertos");
    setSaving(false);
    showToast("🗑️ Sistema zerado! Você tem 7 dias para restaurar.", "#f59e0b");
  };

  const handleRestaurar = async () => {
    setShowConfirmRestore(false);
    setSaving(true);
    if (supabase) {
      await restaurarPedidos(user.id, orgId);
      registrarAuditoria(orgId, user.id, meuNome, "restaurou_sistema", `Restaurou pedidos arquivados.`);
      const pedidos = await fetchPedidos(user.id, orgId);
      setAllPedidos(pedidos);
      setQtdArquivados(0);
    }
    setSaving(false);
    showToast("✅ Pedidos restaurados com sucesso!", "#059669");
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
    if (supabase) {
      await updatePedidoStatus(order.idPedido,user.id,newSt,nota,orgId,meuNome);
      if (orgId) fetchProdutividade(orgId).then(setProdutividade);
    }
    showToast(newSt==="feito"?"✅ Marcado como Feito!":newSt==="vazio"?"📦 Marcado como Vazio!":newSt==="revisao"?"📋 Enviado para Revisão":"↩️ Desmarcado",newSt==="feito"?"#059669":newSt==="vazio"?"#6366f1":newSt==="revisao"?"#f59e0b":"#64748b");
    // Falar frase personalizada ao MARCAR (não ao desmarcar)
    if (newSt && membro) {
      const frase = newSt==="feito" ? membro.frase_feito : newSt==="revisao" ? membro.frase_revisao : newSt==="vazio" ? membro.frase_vazio : null;
      if (membro.som_ativado !== false) falarFrase(frase);
    }
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
  const exportToExcel = async (rows, filename, columnsMap) => {
    if (!rows.length) { showToast("Nada para exportar.","#d97706"); return; }
    const data = rows.map(r => {
      const obj = {};
      for (const [label, key] of Object.entries(columnsMap)) {
        obj[label] = typeof key === "function" ? key(r) : (r[key] ?? "");
      }
      return obj;
    });
    const XL = await getXLSX();
    const ws = XL.utils.json_to_sheet(data);
    const wb = XL.utils.book_new();
    XL.utils.book_append_sheet(wb, ws, "Dados");
    XL.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
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
    navigator.clipboard.writeText(lines.map(l=>"* "+l).join("\n\n")).then(()=>{
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
      (filterFuncionario==="all"||r.feitoPorNome===filterFuncionario)&&
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
    var extras = [];
    if (isAdminEquipe) {
      // Contas a vencer / vencidas (mês atual, ainda não pagas)
      var agora = new Date(); var hoje0 = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
      var ac = alertaContas;
      for (var ci = 0; ci < ac.custos.length; ci++) {
        var c = ac.custos[ci];
        if (!c.data_vencimento || ac.pagos.indexOf(c.id) >= 0) continue;
        var pp = String(c.data_vencimento).slice(0,10).split("-").map(Number);
        var alvo = c.recorrente !== false
          ? new Date(agora.getFullYear(), agora.getMonth(), Math.min(pp[2], new Date(agora.getFullYear(), agora.getMonth()+1, 0).getDate()))
          : new Date(pp[0], pp[1]-1, pp[2]);
        var dias = Math.round((alvo - hoje0) / 86400000);
        if (dias > 3) continue;
        var vf = c.valor ? " · R$ " + Number(c.valor).toFixed(2).replace(".", ",") : "";
        var tit = dias < 0 ? "Conta vencida: " + c.nome : dias === 0 ? "Conta vence hoje: " + c.nome : "Conta vence em " + dias + " dia" + (dias > 1 ? "s" : "") + ": " + c.nome;
        extras.push({ id: "conta__" + c.id + "__" + ac.mes, tipo: "conta", icon: "💸", titulo: tit, desc: "Venc. " + String(alvo.getDate()).padStart(2,"0") + "/" + String(alvo.getMonth()+1).padStart(2,"0") + vf, cor: dias < 0 ? "#ef4444" : "#f59e0b" });
      }
      // Reposição de estoque (consumo dos últimos 30 dias)
      var sem = Math.floor(Date.now() / (7 * 86400000));
      for (var ei = 0; ei < estoqueProdutos.length; ei++) {
        var ep = estoqueProdutos[ei];
        if (ep.ativo === false) continue;
        var cons = Number(consumo30[ep.id]) || 0;
        var qa = Number(ep.quantidade_atual) || 0;
        var porDia = cons / 30;
        var diasRest = porDia > 0 ? qa / porDia : null;
        var baixo = Number(ep.quantidade_minima) > 0 && qa <= Number(ep.quantidade_minima);
        if (!baixo && !(diasRest !== null && diasRest <= 7)) continue;
        var sug = cons > 0 ? Math.max(0, Math.ceil((cons - qa) * 10) / 10) : 0;
        var fmt = function(n) { return String(Number(Number(n).toFixed(2))).replace(".", ","); };
        extras.push({ id: "repor__" + ep.id + "__" + sem, tipo: "estoque", icon: "📦", titulo: "Repor: " + ep.nome,
          desc: "Restam " + fmt(qa) + " " + ep.unidade + (diasRest !== null ? " (~" + Math.max(0, Math.floor(diasRest)) + " dias)" : "") + (sug > 0 ? " · sugestão: comprar " + fmt(sug) + " " + ep.unidade + " (cobre 30 dias)" : ""),
          cor: (diasRest !== null && diasRest <= 3) || qa <= 0 ? "#ef4444" : "#f59e0b" });
      }
    }
    return revisoes.concat(prazosFiltrados, extras);
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
    if (n.tipo === "conta") { setActiveTab("financeiro"); }
    else if (n.tipo === "estoque") { setActiveTab("estoque"); }
    else if (n.tipo === "revisao") { setActiveTab("abertos"); setFilterSt("revisao"); setFilterPrazo("red"); }
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
    ...(isAdminEquipe ? [{id:"estoque",  icon:"🏪", label:"Estoque",  badge: estoqueProdutos.filter(p=>p.quantidade_atual<=p.quantidade_minima).length||null, color:"#0d9488"}] : []),
    ...(isAdminEquipe ? [{id:"produtos", icon: estoquePlusDesbloqueado?"📦":"🔒", label:"Produtos", badge: null, color:"#7c3aed"}] : []),
    ...(isAdminEquipe ? [{id:"financeiro", icon: estoquePlusDesbloqueado?"💰":"🔒", label:"Financeiro", badge: null, color:"#059669"}] : []),
    ...(isAdminEquipe ? [{id:"upseller", icon:"🧾", label:"Pedidos UpSeller", badge: null, color:"#1d4ed8"}] : []),
    ...(isAdminEquipe&&organizacao ? [{id:"equipe",  icon:"👥", label:"Equipe",   badge: membrosEquipe.length||null, color:"#7c3aed"}] : []),
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
          {isAdminEquipe&&<button onClick={abrirHistorico} style={{background:"rgba(255,255,255,0.15)",color:"#fff",border:"1px solid rgba(255,255,255,0.3)",borderRadius:20,padding:"3px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>🕘 Histórico</button>}
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
                      <div style={{fontSize:12,marginTop:4}}>Nenhum prazo, conta ou reposição urgente.</div>
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

        {/* Ações do sistema */}
        {qtdArquivados>0&&(
          <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginBottom:12}}>
            <button onClick={()=>setShowConfirmRestore(true)} style={{background:"#ecfdf5",color:"#065f46",border:"1px solid #6ee7b7",borderRadius:10,padding:"6px 14px",fontSize:11,fontWeight:700,cursor:"pointer"}}>
              ♻️ Restaurar ({qtdArquivados})
            </button>
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
                {organizacao&&membrosEquipe.length>0&&(
                  <select value={filterFuncionario} onChange={e=>setFilterFuncionario(e.target.value)} style={{border:"1px solid #e2e8f0",borderRadius:10,padding:"7px 11px",fontSize:12,cursor:"pointer",background:"#fff"}}>
                    <option value="all">👤 Todos os funcionários</option>
                    <option value="">⏳ Sem responsável</option>
                    {membrosEquipe.map(m=><option key={m.id} value={m.nome}>{m.nome}</option>)}
                  </select>
                )}
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
                    {filtered.length===0&&<tr><td colSpan={12} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>{pedidosAbertos.length===0?"📭 Nenhum pedido ainda.":"Nenhum pedido com esses filtros."}</td></tr>}
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
                    {rankingArr.length===0&&<tr><td colSpan={6} style={{textAlign:"center",padding:36,color:"#94a3b8",fontSize:13}}>Nenhum dado de cliente encontrado.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          );
        })()}

        {/* ── TAB: EQUIPE ── */}
        {activeTab==="equipe"&&isAdminEquipe&&organizacao&&(()=>{
          // Usar UTC para bater com os timestamps do Supabase (que são UTC)
          const agora = new Date();
          const hoje = new Date(Date.UTC(agora.getFullYear(), agora.getMonth(), agora.getDate()));
          const diaAtual = agora.getDay();
          const inicioSemana = new Date(Date.UTC(agora.getFullYear(), agora.getMonth(), agora.getDate() - diaAtual));
          const inicioMes = new Date(Date.UTC(agora.getFullYear(), agora.getMonth(), 1));

          const statsPorMembro = membrosEquipe.map(m => {
            const feitosDoMembro  = produtividade.filter(p => p.feito_por_user_id===m.user_id && p.status_interno==="feito");
            const revisaoDoMembro = produtividade.filter(p => p.feito_por_user_id===m.user_id && p.status_interno==="revisao");
            const vazioDoMembro   = produtividade.filter(p => p.feito_por_user_id===m.user_id && p.status_interno==="vazio");
            const contarPeriodo = (lista, desde) => lista.filter(p => p.feito_em && new Date(p.feito_em) >= desde).length;
            return {
              membro: m,
              feitoHoje:    contarPeriodo(feitosDoMembro, hoje),
              feitoSemana:  contarPeriodo(feitosDoMembro, inicioSemana),
              feitoMes:     contarPeriodo(feitosDoMembro, inicioMes),
              feitoTotal:   feitosDoMembro.length,
              revisaoHoje:  contarPeriodo(revisaoDoMembro, hoje),
              revisaoTotal: revisaoDoMembro.length,
              vazioHoje:    contarPeriodo(vazioDoMembro, hoje),
              vazioTotal:   vazioDoMembro.length,
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
                        <span style={{fontSize:12,color:"#64748b",display:"flex",gap:6,flexWrap:"wrap"}}>
                          <span style={{cursor:"pointer",padding:"2px 8px",borderRadius:20,background:"#ecfdf5"}}
                            onClick={()=>{setActiveTab("abertos");setFilterSt("feito");setFilterFuncionario(s.membro.nome);setTimeout(()=>tableRef.current?.scrollIntoView({behavior:"smooth"}),100);}}>
                            ✅ <strong style={{color:"#059669"}}>{s.feitoHoje}</strong> hoje · <strong>{s.feitoTotal}</strong> total
                          </span>
                          {(s.revisaoHoje>0||s.revisaoTotal>0)&&(
                            <span style={{cursor:"pointer",padding:"2px 8px",borderRadius:20,background:"#fffbeb"}}
                              onClick={()=>{setActiveTab("abertos");setFilterSt("revisao");setFilterFuncionario(s.membro.nome);setTimeout(()=>tableRef.current?.scrollIntoView({behavior:"smooth"}),100);}}>
                              📋 <strong style={{color:"#f59e0b"}}>{s.revisaoHoje}</strong> hoje · <strong>{s.revisaoTotal}</strong> total
                            </span>
                          )}
                          {(s.vazioHoje>0||s.vazioTotal>0)&&(
                            <span style={{cursor:"pointer",padding:"2px 8px",borderRadius:20,background:"#f5f3ff"}}
                              onClick={()=>{setActiveTab("abertos");setFilterSt("vazio");setFilterFuncionario(s.membro.nome);setTimeout(()=>tableRef.current?.scrollIntoView({behavior:"smooth"}),100);}}>
                              📦 <strong style={{color:"#6366f1"}}>{s.vazioHoje}</strong> hoje · <strong>{s.vazioTotal}</strong> total
                            </span>
                          )}
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
                    {["Funcionário","Cor","Financeiro","Zerar Sistema","Editar Status","🔊 Frases de Voz",""].map(h=><th key={h} style={TH}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {membrosEquipe.map((m,i)=>(
                      <tr key={m.id} style={{background:i%2===0?"#fff":"#f8fafc",borderBottom:"1px solid #f1f5f9"}}>
                        <td style={TD}><strong>{m.nome}</strong><div style={{fontSize:10,color:"#94a3b8"}}>{m.email}</div></td>
                        <td style={{...TD,textAlign:"center"}}><div style={{width:18,height:18,borderRadius:"50%",background:m.cor,margin:"0 auto"}}/></td>
                        {["pode_ver_financeiro","pode_zerar_sistema","pode_editar_status"].map(perm=>(
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
                        <td style={{...TD,minWidth:220}}>
                          <div style={{display:"flex",flexDirection:"column",gap:4}}>
                            {[
                              {key:"frase_feito",   icon:"✅", placeholder:"Ex: Parabéns Mylena!"},
                              {key:"frase_revisao", icon:"📋", placeholder:"Ex: Revise no final do dia"},
                              {key:"frase_vazio",   icon:"📦", placeholder:"Ex: Produto sem estoque"},
                            ].map(f=>(
                              <div key={f.key} style={{display:"flex",alignItems:"center",gap:4}}>
                                <span style={{fontSize:12}}>{f.icon}</span>
                                <input
                                  defaultValue={m[f.key]||""}
                                  placeholder={f.placeholder}
                                  onBlur={async e=>{
                                    const val = e.target.value.trim()||null;
                                    if(val===(m[f.key]||null)) return;
                                    setMembrosEquipe(prev=>prev.map(x=>x.id===m.id?{...x,[f.key]:val}:x));
                                    await atualizarPermissoesMembro(m.id,{[f.key]:val});
                                    showToast("🔊 Frase salva!","#059669");
                                  }}
                                  style={{flex:1,padding:"3px 7px",border:"1px solid #e2e8f0",borderRadius:7,fontSize:11,outline:"none"}}
                                />
                              </div>
                            ))}
                            <label style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:"#64748b",cursor:"pointer",marginTop:2}}>
                              <input type="checkbox" checked={m.som_ativado!==false}
                                onChange={async e=>{
                                  const val=e.target.checked;
                                  setMembrosEquipe(prev=>prev.map(x=>x.id===m.id?{...x,som_ativado:val}:x));
                                  await atualizarPermissoesMembro(m.id,{som_ativado:val});
                                }} style={{width:13,height:13}} />
                              Som ativado
                            </label>
                          </div>
                        </td>
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
              Os pedidos serão <strong>arquivados</strong> por <strong>7 dias</strong> e você poderá restaurá-los nesse período.<br/>
              Devoluções e faturamento também serão apagados.<br/>
              <strong style={{color:"#f59e0b"}}>Após 7 dias a exclusão é permanente.</strong>
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
      {/* ── Modal confirmação restaurar ── */}
      {showConfirmRestore&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}
          onClick={e=>e.target===e.currentTarget&&setShowConfirmRestore(false)}>
          <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:400,boxShadow:"0 20px 60px rgba(0,0,0,0.25)",textAlign:"center"}}>
            <div style={{fontSize:44,marginBottom:12}}>♻️</div>
            <h3 style={{margin:"0 0 8px",fontSize:17,fontWeight:700,color:"#1e293b"}}>Restaurar Pedidos?</h3>
            <p style={{margin:"0 0 24px",fontSize:13,color:"#64748b",lineHeight:1.6}}>
              Isso vai restaurar <strong>{qtdArquivados} pedido(s)</strong> arquivados de volta ao sistema.<br/>
              <strong style={{color:"#059669"}}>Todos os status e informações serão preservados.</strong>
            </p>
            <div style={{display:"flex",gap:10,justifyContent:"center"}}>
              <button onClick={()=>setShowConfirmRestore(false)} style={{padding:"10px 24px",border:"1px solid #e2e8f0",borderRadius:12,background:"#fff",fontSize:13,fontWeight:600,cursor:"pointer",color:"#374151"}}>
                Cancelar
              </button>
              <button onClick={handleRestaurar} style={{padding:"10px 24px",border:"none",borderRadius:12,background:"#059669",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                ♻️ Sim, Restaurar
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
      {/* ── Histórico de operações / desfazer ── */}
      {showHistorico&&isAdminEquipe&&(()=>{
        const grupos = [];
        const porLote = {};
        operacoes.forEach(o=>{
          if (o.lote_id) {
            if (!porLote[o.lote_id]) { porLote[o.lote_id] = {key:o.lote_id, ops:[], tipo:o.tipo}; grupos.push(porLote[o.lote_id]); }
            porLote[o.lote_id].ops.push(o);
          } else grupos.push({key:o.id, ops:[o], tipo:o.tipo});
        });
        grupos.forEach(g=>{
          const ordenadas = [...g.ops].sort((a,b)=>a.seq-b.seq);
          g.ops = ordenadas;
          g.titulo = ordenadas[0].descricao;
          g.quando = ordenadas[0].criado_em;
          g.quem = ordenadas[0].usuario_nome||"";
          g.desfeito = ordenadas.every(o=>o.desfeito_em);
          g.parcial = !g.desfeito && ordenadas.some(o=>o.desfeito_em);
          g.por = ordenadas.find(o=>o.desfeito_por)?.desfeito_por||"";
          g.quandoDesfeito = ordenadas.find(o=>o.desfeito_em)?.desfeito_em;
        });
        const filtros = [["todos","Tudo"],["estoque","📦 Estoque"],["financeiro","💰 Financeiro"],["produtos","🧱 Produtos"],["fornecedores","🏭 Fornecedores"],["nota_fiscal","📄 Notas"]];
        const visiveis = grupos.filter(g=>filtroHist==="todos"||g.tipo===filtroHist);
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:22,width:"100%",maxWidth:760,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6,gap:10}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:800,color:"#1e293b"}}>🕘 Histórico de operações</h3>
                <button onClick={()=>setShowHistorico(false)} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Fechar</button>
              </div>
              <p style={{margin:"0 0 12px",fontSize:12,color:"#64748b"}}>Tudo que foi digitado ou lançado fica aqui. Clique em <strong>↩ Desfazer</strong> para voltar atrás.</p>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
                {filtros.map(([k,l])=>(
                  <button key={k} onClick={()=>setFiltroHist(k)} style={{padding:"5px 12px",borderRadius:16,border:`1.5px solid ${filtroHist===k?"#7c3aed":"#e2e8f0"}`,background:filtroHist===k?"#f5f3ff":"#fff",color:filtroHist===k?"#7c3aed":"#64748b",fontSize:11,fontWeight:700,cursor:"pointer"}}>{l}</button>
                ))}
              </div>
              {carregandoOps ? (
                <div style={{textAlign:"center",padding:30,color:"#94a3b8",fontSize:13}}>Carregando...</div>
              ) : visiveis.length===0 ? (
                <div style={{textAlign:"center",padding:30,color:"#94a3b8",fontSize:13}}>Nenhuma operação registrada ainda.</div>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {visiveis.map(g=>(
                    <div key={g.key} style={{border:"1px solid #e2e8f0",borderRadius:12,padding:"10px 12px",background:g.desfeito?"#f8fafc":"#fff",opacity:g.desfeito?0.65:1}}>
                      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:10}}>
                        <div style={{minWidth:0}}>
                          <div style={{fontSize:13,fontWeight:700,color:"#1e293b",textDecoration:g.desfeito?"line-through":"none",wordBreak:"break-word"}}>{g.titulo}</div>
                          <div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>{quandoFmt(g.quando)}{g.quem?` · ${g.quem}`:""}</div>
                          {g.desfeito&&<div style={{fontSize:11,color:"#64748b",marginTop:2}}>↩ Desfeito{g.por?` por ${g.por}`:""}{g.quandoDesfeito?` em ${quandoFmt(g.quandoDesfeito)}`:""}</div>}
                          {g.parcial&&<div style={{fontSize:11,color:"#f59e0b",marginTop:2}}>⚠️ Parcialmente desfeito</div>}
                        </div>
                        {!g.desfeito&&(
                          <button onClick={()=>desfazerGrupo(g)} disabled={desfazendoKey===g.key}
                            style={{flexShrink:0,background:"#fef3c7",color:"#92400e",border:"1.5px solid #fde68a",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:800,cursor:"pointer"}}>
                            {desfazendoKey===g.key?"Desfazendo...":"↩ Desfazer"}
                          </button>
                        )}
                      </div>
                      {g.ops.length>1&&(
                        <div style={{marginTop:6,paddingLeft:10,borderLeft:"3px solid #e2e8f0"}}>
                          {g.ops.slice(1).map(o=>(
                            <div key={o.id} style={{fontSize:11,color:"#64748b",textDecoration:o.desfeito_em?"line-through":"none"}}>• {o.descricao}</div>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        );
      })()}

      {/* ── Ler nota fiscal ── */}
      {showNota&&isAdminEquipe&&(()=>{
        const inp = {width:"100%",padding:"8px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:13,outline:"none",boxSizing:"border-box"};
        const lb = {fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2};
        const validos = notaItens.filter(i=>i.destino!=="ignorar"&&Number(i.qtd)>0);
        const somaItens = validos.reduce((a,i)=>a+Number(i.qtd)*(Number(i.preco)||0),0);
        const totalNotaLido = Number(notaCab.total)||0;
        const diverge = totalNotaLido>0 && Math.abs(totalNotaLido-somaItens)>1;
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:22,width:"100%",maxWidth:820,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:10}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:800,color:"#1e293b"}}>📷 Ler nota fiscal</h3>
                <button onClick={()=>setShowNota(false)} disabled={salvandoNota} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Fechar</button>
              </div>

              {(notaFase==="upload"||notaFase==="lendo")&&(
                <div>
                  <p style={{margin:"0 0 12px",fontSize:12,color:"#64748b"}}>Envie uma <strong>foto</strong> ou o <strong>PDF</strong> da nota. O sistema lê os dados e você confere tudo antes de lançar no estoque.</p>
                  {notaFase==="lendo" ? (
                    <div style={{textAlign:"center",padding:"36px 0",color:"#0d9488",fontSize:14,fontWeight:700}}>🔎 Lendo a nota... pode levar alguns segundos.</div>
                  ) : (
                    <label style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:6,padding:"32px 16px",border:"2px dashed #99f6e4",borderRadius:14,background:"#f0fdfa",cursor:"pointer",textAlign:"center"}}>
                      <span style={{fontSize:34}}>🧾</span>
                      <span style={{fontSize:13,fontWeight:700,color:"#0f766e"}}>Tirar foto ou escolher arquivo</span>
                      <span style={{fontSize:11,color:"#64748b"}}>JPG, PNG ou PDF · até 15 MB</span>
                      <input type="file" accept="image/*,application/pdf" onChange={e=>{const f=e.target.files?.[0]; e.target.value=""; lerArquivoNota(f);}} style={{display:"none"}} />
                    </label>
                  )}
                  {notaErro&&<div style={{marginTop:10,padding:"9px 12px",background:"#fde8e8",color:"#991b1b",borderRadius:10,fontSize:12}}>⚠️ {notaErro}</div>}
                </div>
              )}

              {notaFase==="revisar"&&(
                <div>
                  <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:10,padding:"8px 12px",fontSize:12,color:"#92400e",marginBottom:12}}>
                    ⚠️ Confira os dados abaixo com a nota — a leitura automática pode errar números. Nada é lançado até você clicar em <strong>Lançar no estoque</strong>.
                  </div>
                  {notaDuplicada&&<div style={{background:"#fde8e8",border:"1px solid #fecaca",borderRadius:10,padding:"8px 12px",fontSize:12,color:"#991b1b",marginBottom:12}}>🚫 Já existe um lançamento com o nº {notaCab.numero}. Provavelmente essa nota já foi lançada.</div>}
                  <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:12}}>
                    {notaPreview&&(
                      <div style={{width:130,flexShrink:0}}>
                        {notaPreview.tipo==="img"
                          ? <a href={notaPreview.url} target="_blank" rel="noreferrer"><img src={notaPreview.url} alt="nota" style={{width:"100%",maxHeight:170,objectFit:"cover",borderRadius:10,border:"1px solid #e2e8f0"}} /></a>
                          : <a href={notaPreview.url} target="_blank" rel="noreferrer" style={{display:"block",padding:"24px 8px",textAlign:"center",border:"1px solid #e2e8f0",borderRadius:10,fontSize:12,color:"#0f766e",textDecoration:"none"}}>📄 {notaPreview.nome}<br/><span style={{fontSize:10}}>abrir PDF</span></a>}
                      </div>
                    )}
                    <div style={{flex:1,minWidth:240,display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                      <div style={{gridColumn:"1 / -1"}}>
                        <label style={lb}>Fornecedor</label>
                        <select value={notaFornId} onChange={e=>setNotaFornId(e.target.value)} style={{...inp,background:"#fff",marginBottom:4}}>
                          <option value="">{notaCab.fornecedor?`➕ Criar novo: ${notaCab.fornecedor}`:"(sem fornecedor)"}</option>
                          {fornecedores.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}
                        </select>
                        {!notaFornId&&<input value={notaCab.fornecedor} onChange={e=>setNotaCab(c=>({...c,fornecedor:e.target.value}))} placeholder="Nome do fornecedor" style={inp} />}
                      </div>
                      <div><label style={lb}>Nº da nota</label><input value={notaCab.numero} onChange={e=>{const v=e.target.value;setNotaCab(c=>({...c,numero:v}));}} onBlur={async()=>setNotaDuplicada(await notaJaLancada(user.id,orgId,notaCab.numero.trim()))} style={inp} /></div>
                      <div><label style={lb}>Data da nota</label><input type="date" value={notaCab.data} onChange={e=>setNotaCab(c=>({...c,data:e.target.value}))} style={inp} /></div>
                    </div>
                  </div>

                  <div style={{fontSize:12,fontWeight:700,color:"#1e293b",marginBottom:6}}>Itens da nota</div>
                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    {notaItens.map(it=>{
                      const prod = estoqueProdutos.find(x=>x.id===it.destino);
                      const ign = it.destino==="ignorar";
                      const un = prod?.unidade||it.un;
                      return (
                        <div key={it.key} style={{border:"1.5px solid #e2e8f0",borderRadius:12,padding:"10px 12px",background:ign?"#f8fafc":"#fff",opacity:ign?0.55:1}}>
                          <div style={{fontSize:11,color:"#64748b",marginBottom:6}}>Na nota: <strong style={{color:"#1e293b"}}>{it.descricao||"(sem descrição)"}</strong>{it.unNota?` · ${it.unNota}`:""}</div>
                          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:8,alignItems:"end"}}>
                            <div style={{gridColumn:"span 2"}}>
                              <label style={lb}>Vai para o material</label>
                              <select value={it.destino} onChange={e=>setNotaItem(it.key,{destino:e.target.value})} style={{...inp,background:"#fff"}}>
                                {estoqueProdutos.map(x=><option key={x.id} value={x.id}>{x.nome} ({x.unidade})</option>)}
                                <option value="novo">➕ Criar novo material</option>
                                <option value="ignorar">⏭ Ignorar este item</option>
                              </select>
                            </div>
                            {!ign&&it.destino==="novo"&&(
                              <>
                                <div style={{gridColumn:"span 2"}}><label style={lb}>Nome do novo material</label><input value={it.nome} onChange={e=>setNotaItem(it.key,{nome:e.target.value})} style={inp} /></div>
                                <div><label style={lb}>Unidade</label>
                                  <select value={it.un} onChange={e=>setNotaItem(it.key,{un:e.target.value})} style={{...inp,background:"#fff"}}>{["un","cx","pct","rolo","m","kg","l"].map(u=><option key={u} value={u}>{u}</option>)}</select>
                                </div>
                              </>
                            )}
                            {!ign&&(
                              <>
                                <div><label style={lb}>Quantidade ({un})</label><input type="number" min="0" step="any" value={it.qtd} onChange={e=>setNotaItem(it.key,{qtd:e.target.value})} style={inp} /></div>
                                <div><label style={lb}>Preço por {un} (R$)</label><input type="number" min="0" step="0.01" value={it.preco} onChange={e=>setNotaItem(it.key,{preco:e.target.value})} style={inp} /></div>
                                <div style={{fontSize:12,color:"#374151",paddingBottom:8}}>Subtotal <strong style={{color:"#059669"}}>{fmtBRLn(Number(it.qtd)*(Number(it.preco)||0))}</strong></div>
                              </>
                            )}
                          </div>
                          {!ign&&prod&&Number(prod.preco_unitario)>0&&Number(it.preco)>0&&(()=>{
                            const ant=Number(prod.preco_unitario), novo=Number(it.preco), dif=((novo-ant)/ant)*100;
                            if (Math.abs(dif)<5) return null;
                            const sobe=dif>0;
                            return <div style={{fontSize:11,fontWeight:600,marginTop:4,color:sobe?"#991b1b":"#065f46",background:sobe?"#fde8e8":"#d1fae5",borderRadius:6,padding:"4px 8px"}}>{sobe?"📈 Preço subiu":"📉 Preço caiu"} {Math.abs(dif).toFixed(0)}% vs. a última compra ({fmtBRLn(ant)} → {fmtBRLn(novo)})</div>;
                          })()}
                          {!ign&&prod&&it.unNota&&mapUnidade(it.unNota)!==prod.unidade&&(
                            <div style={{fontSize:10,color:"#f59e0b",marginTop:4}}>⚠️ A nota vem em "{it.unNota}" e o estoque controla em "{prod.unidade}". Ajuste a quantidade e o preço para a unidade do estoque.</div>
                          )}
                        </div>
                      );
                    })}
                  </div>

                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,marginTop:12,padding:"10px 12px",background:"#f0fdfa",borderRadius:10}}>
                    <span style={{fontSize:13}}>Total a lançar: <strong style={{color:"#0f766e"}}>{fmtBRLn(somaItens)}</strong></span>
                    {totalNotaLido>0&&<span style={{fontSize:11,color:diverge?"#991b1b":"#64748b"}}>{diverge?"⚠️ ":""}Total lido na nota: {fmtBRLn(totalNotaLido)}{diverge?" — diferente da soma dos itens (frete/desconto/impostos ou erro de leitura)":""}</span>}
                  </div>
                  <div style={{display:"flex",gap:10,marginTop:14}}>
                    <button onClick={()=>setNotaFase("upload")} disabled={salvandoNota} style={{flex:1,padding:"11px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Enviar outra</button>
                    <button onClick={confirmarNota} disabled={salvandoNota||!validos.length}
                      style={{flex:2,padding:"11px 0",border:"none",borderRadius:10,background:validos.length?"#0d9488":"#e2e8f0",color:validos.length?"#fff":"#94a3b8",fontSize:13,fontWeight:800,cursor:"pointer"}}>
                      {salvandoNota?"Lançando...":"✅ Lançar no estoque"}
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        );
      })()}

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

      {/* ── TAB: PRODUTOS / CALCULADORA ── */}
      {activeTab==="produtos"&&isAdminEquipe&&!estoquePlusDesbloqueado&&(
        <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:320,padding:32}}>
          <div style={{background:"#fff",borderRadius:20,padding:32,maxWidth:340,width:"100%",boxShadow:"0 4px 20px rgba(0,0,0,0.08)",textAlign:"center"}}>
            <div style={{fontSize:44,marginBottom:10}}>🔒</div>
            <h3 style={{margin:"0 0 6px",fontSize:17,fontWeight:800,color:"#1e293b"}}>Área Protegida</h3>
            <p style={{margin:"0 0 20px",fontSize:12,color:"#64748b"}}>A aba Produtos contém cálculos e valores sigilosos. Digite o PIN para continuar.</p>
            {!estoquePlusConfig ? (
              <div>
                <p style={{fontSize:12,color:"#f59e0b",marginBottom:12}}>Você ainda não definiu um PIN. Configure pelo botão <strong>🔒 Estoque+</strong> no Estoque.</p>
                <button onClick={()=>{setActiveTab("estoque");setShowDefinirPin(true);}} style={{width:"100%",padding:"10px 0",background:"#7c3aed",color:"#fff",border:"none",borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer"}}>Definir PIN agora</button>
              </div>
            ) : (
              <div>
                <input type="password" inputMode="numeric" maxLength={8} value={pinInput}
                  onChange={e=>{setPinInput(e.target.value);setPinErro("");}}
                  onKeyDown={async e=>{
                    if (e.key==="Enter") {
                      const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pinInput))
                        .then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
                      if (hash===estoquePlusConfig?.pin_hash) {
                        setEstoquePlusDesbloqueado(true); setPinInput(""); setPinErro("");
                      } else { setPinErro("PIN incorreto."); setPinInput(""); }
                    }
                  }}
                  autoFocus placeholder="••••"
                  style={{width:"100%",padding:"12px 16px",border:`2px solid ${pinErro?"#ef4444":"#e2e8f0"}`,borderRadius:12,fontSize:22,outline:"none",textAlign:"center",letterSpacing:8,boxSizing:"border-box",marginBottom:8}} />
                {pinErro&&<div style={{fontSize:12,color:"#ef4444",marginBottom:8}}>{pinErro}</div>}
                <button onClick={async()=>{
                  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pinInput))
                    .then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
                  if (hash===estoquePlusConfig?.pin_hash) {
                    setEstoquePlusDesbloqueado(true); setPinInput(""); setPinErro("");
                  } else { setPinErro("PIN incorreto."); setPinInput(""); }
                }} style={{width:"100%",padding:"12px 0",background:"#7c3aed",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:800,cursor:"pointer"}}>
                  Desbloquear
                </button>
              </div>
            )}
          </div>
        </div>
      )}
      {activeTab==="produtos"&&isAdminEquipe&&estoquePlusDesbloqueado&&(()=>{
        const num = (s,def) => (s===""||s==null||isNaN(Number(s))) ? def : Number(s);
        const fmtR = (n) => "R$ "+Number(n||0).toFixed(2);
        const reloadProdutos = async () => { const todos = await fetchProdutos(user.id, orgId); setProdutosCalc(todos); return todos; };
        const reloadMps = async () => { const l = await fetchMateriasPrimas(user.id, orgId); setMateriasPrimas(l); return l; };

        const inp = {width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"};
        const lbl = {fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3};

        // ── Cálculos ────────────────────────────────────────────────────────
        const dadosMp = (l) => l.materias_primas || {nome:l.nome, custo_unidade_compra:l.custo_unidade_compra, qtd_por_unidade_compra:l.qtd_por_unidade_compra, unidade:l.unidade};
        const custoUnitMp = (m) => Number(m.qtd_por_unidade_compra)>0 ? Number(m.custo_unidade_compra)/Number(m.qtd_por_unidade_compra) : 0;
        const custoLinha = (l) => custoUnitMp(dadosMp(l)) * Number(l.qtd_usada_por_produto||0);
        const custoVar = (v) => (v.produto_materias_primas||[]).reduce((a,l)=>a+custoLinha(l),0);
        const pctTaxas = (p) => num(p.taxa_plataforma_pct,0)+num(p.taxa_imposto_pct,0)+num(p.taxa_adiantamento_pct,0);
        const liquidoPreco = (p,preco) => Number(preco||0)*(1-pctTaxas(p)/100) - num(p.taxa_fixa,0);
        const liquidoVar = (p,v) => liquidoPreco(p, v.preco_venda);
        const margemDe = (liquido,custo) => custo>0 ? (liquido/custo)*100 : null;
        const margemVar = (p,v) => margemDe(liquidoVar(p,v), custoVar(v));
        const margemMedia = (p) => {
          const ms = (p.produto_variacoes||[]).map(v=>margemVar(p,v)).filter(m=>m!==null);
          return ms.length ? ms.reduce((a,b)=>a+b,0)/ms.length : null;
        };
        const corMargem = (m) => m===null?"#94a3b8":m>=100?"#059669":m>=50?"#f59e0b":"#ef4444";
        const bgMargem = (m) => m===null?"#f8fafc":m>=100?"#d1fae5":m>=50?"#fffbeb":"#fde8e8";
        const fmtM = (m) => m===null?"—":Math.round(m)+"%";
        // Margem sobre o preço de venda: quanto do preço vira lucro (depois de taxas e custo)
        const margemPreco = (p,v) => { const pr=Number(v.preco_venda)||0; return pr>0 ? ((liquidoVar(p,v)-custoVar(v))/pr)*100 : null; };
        const corMP = (m) => m===null?"#94a3b8":m>=30?"#059669":m>=15?"#f59e0b":"#ef4444";
        const usosDaMp = (mpId) => produtosCalc.reduce((n,p)=>n+(p.produto_variacoes||[]).reduce((k,v)=>k+(v.produto_materias_primas||[]).filter(l=>l.materia_prima_id===mpId).length,0),0);

        // ── Editor de produto (novo / editar) ───────────────────────────────
        const novaVar = (nome="") => ({key:Math.random().toString(36).slice(2), nome, preco:"", mps:[{mpId:"",qtd:""}]});
        const abrirFormNovoProd = () => {
          setEditandoProdutoCalc(null);
          setProdNome(""); setProdTaxaPlat("20"); setProdTaxaImp("4"); setProdTaxaAdiant("3.5"); setProdTaxaFixa("4");
          setProdVars([novaVar("")]);
          setShowProdutoForm(true);
        };
        const abrirFormEditarProd = (p) => {
          setEditandoProdutoCalc(p);
          setProdNome(p.nome); setProdTaxaPlat(String(p.taxa_plataforma_pct)); setProdTaxaImp(String(p.taxa_imposto_pct)); setProdTaxaAdiant(String(p.taxa_adiantamento_pct)); setProdTaxaFixa(String(p.taxa_fixa));
          const vs = (p.produto_variacoes||[]).map(v=>({
            key:v.id, id:v.id, nome:v.nome, preco:String(v.preco_venda),
            mps:(v.produto_materias_primas||[]).map(l=>({mpId:l.materia_prima_id||"", qtd:String(l.qtd_usada_por_produto)})),
          }));
          setProdVars(vs.length?vs:[novaVar("")]);
          setShowProdutoForm(true);
        };
        const setVar = (key,patch) => setProdVars(vs=>vs.map(v=>v.key===key?{...v,...patch}:v));
        const setVarMp = (key,idx,patch) => setProdVars(vs=>vs.map(v=>v.key===key?{...v,mps:v.mps.map((m,i)=>i===idx?{...m,...patch}:m)}:v));
        const addVarMp = (key) => setProdVars(vs=>vs.map(v=>v.key===key?{...v,mps:[...v.mps,{mpId:"",qtd:""}]}:v));
        const delVarMp = (key,idx) => setProdVars(vs=>vs.map(v=>v.key===key?{...v,mps:v.mps.length>1?v.mps.filter((_,i)=>i!==idx):[{mpId:"",qtd:""}]}:v));
        const addVar = () => setProdVars(vs=>[...vs, novaVar("")]);
        const dupVar = (key) => setProdVars(vs=>{
          const i = vs.findIndex(v=>v.key===key); if (i<0) return vs;
          const o = vs[i];
          const copia = {key:Math.random().toString(36).slice(2), nome:(o.nome?o.nome+" (cópia)":""), preco:o.preco, mps:o.mps.map(m=>({...m}))};
          return [...vs.slice(0,i+1), copia, ...vs.slice(i+1)];
        });
        const delVar = (key) => setProdVars(vs=>vs.length>1?vs.filter(v=>v.key!==key):vs);

        const pEdit = {taxa_plataforma_pct:num(prodTaxaPlat,20), taxa_imposto_pct:num(prodTaxaImp,4), taxa_adiantamento_pct:num(prodTaxaAdiant,3.5), taxa_fixa:num(prodTaxaFixa,4)};
        // Kit: expande em linhas normais de matéria-prima (qtd do kit × fator de cada item)
        const expandirMps = (mps) => mps.flatMap(m=>{
          if (m.kitId===undefined) return [m];
          const kit = kits.find(k=>k.id===m.kitId);
          return kit ? (kit.itens||[]).map(i=>({mpId:i.mpId,qtd:(Number(m.qtd)||0)*(Number(i.fator)||0)})) : [];
        });
        const custoEditor = (v) => expandirMps(v.mps).reduce((a,m)=>{ const mp=materiasPrimas.find(x=>x.id===m.mpId); return mp ? a+custoUnitMp(mp)*(Number(m.qtd)||0) : a; },0);
        const addVarKit = (key) => setProdVars(vs=>vs.map(v=>v.key===key?{...v,mps:[...v.mps.filter(m=>m.kitId!==undefined||m.mpId),{kitId:"",mpId:"",qtd:""}]}:v));
        const abrirKitForm = (k=null) => {
          setEditandoKit(k); setKitNome(k?.nome||"");
          setKitItens(k&&(k.itens||[]).length?k.itens.map(i=>({mpId:i.mpId,fator:String(i.fator)})):[{mpId:"",fator:"1"}]);
          setShowKitForm(true);
        };
        const salvarKitForm = async () => {
          const itens = kitItens.filter(i=>i.mpId&&Number(i.fator)>0).map(i=>({mpId:i.mpId,fator:Number(i.fator)}));
          if (!kitNome.trim()||itens.length===0) { showToast("❌ Dê um nome e escolha ao menos uma matéria-prima.","#991b1b"); return; }
          setSalvandoKit(true);
          const r = await salvarKit(orgId, {id:editandoKit?.id,nome:kitNome.trim(),itens});
          setSalvandoKit(false);
          if (r.error) { showToast("❌ "+r.error,"#991b1b"); return; }
          setShowKitForm(false);
          setKits(await fetchKits(orgId));
          showToast("✅ Kit salvo!","#059669");
        };
        const excluirKitLista = async (k) => {
          if (!window.confirm(`Remover o kit "${k.nome}"? Produtos que já usam as matérias-primas dele não mudam.`)) return;
          await excluirKit(k.id);
          setKits(await fetchKits(orgId));
        };

        const salvarProdutoCalc = async () => {
          if (!prodNome.trim()) return;
          const unica = prodVars.length===1;
          if (prodVars.some(v=>!v.nome.trim()&&!unica)) { showToast("❌ Dê um nome a cada variação.","#991b1b"); return; }
          setSalvandoProdutoCalc(true);
          const antesSnap = editandoProdutoCalc ? snapProduto(editandoProdutoCalc) : null;
          const variacoes = prodVars.map(v=>({
            id: v.id,
            nome: v.nome.trim()||"Padrão",
            preco_venda: Number(v.preco)||0,
            mps: expandirMps(v.mps).filter(m=>m.mpId&&Number(m.qtd)>0).map(m=>({
              materia_prima_id: m.mpId,
              nome: materiasPrimas.find(x=>x.id===m.mpId)?.nome||"",
              qtd_usada: Number(m.qtd),
            })),
          }));
          const res = await salvarProdutoCompleto(user.id, orgId, {
            id: editandoProdutoCalc?.id,
            nome: prodNome.trim(),
            taxa_plataforma_pct: pEdit.taxa_plataforma_pct,
            taxa_imposto_pct: pEdit.taxa_imposto_pct,
            taxa_adiantamento_pct: pEdit.taxa_adiantamento_pct,
            taxa_fixa: pEdit.taxa_fixa,
          }, variacoes);
          setSalvandoProdutoCalc(false);
          if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
          setShowProdutoForm(false);
          logOp({tipo:"produtos",acao:editandoProdutoCalc?"editar":"criar",descricao:(editandoProdutoCalc?"Produto editado: ":"Produto criado: ")+prodNome.trim(),tabela:"produtos",registro_id:res.id,antes:antesSnap});
          const todos = await reloadProdutos();
          if (editandoProdutoCalc) {
            if (produtoAtivo) setProdutoAtivo(todos.find(p=>p.id===produtoAtivo.id)||null);
          } else if (res.id) {
            const novo = todos.find(p=>p.id===res.id);
            if (novo) setProdutoAtivo(novo);
          }
          showToast(editandoProdutoCalc?"✅ Produto atualizado!":"✅ Produto criado!","#059669");
        };

        // ── Biblioteca de matérias-primas ───────────────────────────────────
        const abrirMpForm = (mp = null) => {
          setEditandoMp(mp);
          setMpNome(mp?.nome||""); setMpCustoCompra(mp?String(mp.custo_unidade_compra):""); setMpQtdPorCompra(mp?String(mp.qtd_por_unidade_compra):""); setMpUnidade(mp?.unidade||"folha");
          setMpEstoqueId(mp?.estoque_produto_id||""); setMpRendimento(mp?.rendimento_estoque?String(mp.rendimento_estoque):"");
          setMpModoDimensao(false); setMpRoloLargura(""); setMpRoloComprimento(""); setMpCorteL(""); setMpCorteC(""); setMpUnidadeCorte("folha");
          setShowMpForm(true);
        };
        const puxarDoEstoque = (id) => {
          setMpEstoqueId(id);
          const ep = estoqueProdutos.find(x=>x.id===id);
          if (!ep) return;
          const unidades = ["folha","pedaço","un","cx","pct","rolo","m","cm","kg","g","l","ml"];
          setMpNome(ep.nome);
          setMpCustoCompra(ep.preco_unitario?Number(ep.preco_unitario).toFixed(2):"");
          setMpQtdPorCompra("1");
          setMpUnidade(unidades.includes(ep.unidade)?ep.unidade:"un");
          setMpRendimento("");
        };
        const calcDimensao = () => {
          const rl = Number(mpRoloLargura); const rc = Number(mpRoloComprimento)*100;
          const cl = Number(mpCorteL); const cc = Number(mpCorteC);
          if (!rl||!rc||!cl||!cc) return null;
          const custo_compra = Number(mpCustoCompra)||0;
          const opcoes = [];
          const variantes = [
            {label:"Corte padrão", larg:cl, comp:cc},
            {label:"Corte girado 90°", larg:cc, comp:cl},
          ];
          if (cl !== cc) {
            variantes.push({label:"Metade da largura", larg:cl/2, comp:cc});
            variantes.push({label:"Metade do comprimento", larg:cl, comp:cc/2});
          }
          variantes.forEach(v => {
            if (!v.larg||!v.comp) return;
            const cols = Math.floor(rl/v.larg);
            const lins = Math.floor(rc/v.comp);
            const tot  = cols*lins;
            if (tot<=0) return;
            const sobra_larg = rl - cols*v.larg;
            const sobra_comp = rc - lins*v.comp;
            const aproveitamento = ((rl*rc - sobra_larg*rc - sobra_comp*rl + sobra_larg*sobra_comp)/(rl*rc)*100);
            opcoes.push({
              label: v.label,
              colunas: cols, linhas: lins, total: tot,
              custo_cada: custo_compra>0&&tot>0 ? custo_compra/tot : 0,
              aproveitamento: Math.round(aproveitamento),
              sobra_larg: sobra_larg.toFixed(1), sobra_comp: sobra_comp.toFixed(1),
              larg_usada: v.larg, comp_usada: v.comp,
            });
          });
          opcoes.sort((a,b)=>b.total-a.total);
          const vistos = new Set();
          const unicas = opcoes.filter(o=>{ if(vistos.has(o.total)) return false; vistos.add(o.total); return true; });
          return unicas.length>0 ? unicas : null;
        };
        const aplicarDimensao = (idx=0) => {
          const opcoes = calcDimensao();
          if (!opcoes||!opcoes[idx]) return;
          const r = opcoes[idx];
          setMpQtdPorCompra(String(r.total));
          setMpUnidade(mpUnidadeCorte);
          setMpModoDimensao(false);
        };
        const salvarMp = async () => {
          if (!mpNome.trim()) return;
          setSalvandoMp(true);
          const res = await salvarMateriaPrimaLib(user.id, orgId, {
            id: editandoMp?.id,
            nome: mpNome.trim(),
            custo_unidade_compra: Number(mpCustoCompra)||0,
            qtd_por_unidade_compra: Number(mpQtdPorCompra)||1,
            unidade: mpUnidade,
            estoque_produto_id: mpEstoqueId||null,
            rendimento_estoque: mpEstoqueId&&Number(mpRendimento)>0 ? Number(mpRendimento) : null,
          });
          setSalvandoMp(false);
          if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
          setShowMpForm(false);
          logOp({tipo:"produtos",acao:editandoMp?"editar":"criar",descricao:(editandoMp?"Matéria-prima editada: ":"Matéria-prima criada: ")+mpNome.trim(),tabela:"materias_primas",registro_id:res.id,antes:editandoMp?pick(editandoMp,COLS_MP):null});
          await reloadMps();
          const todos = await reloadProdutos();
          if (produtoAtivo) setProdutoAtivo(todos.find(p=>p.id===produtoAtivo.id)||null);
          showToast("✅ Matéria-prima salva!","#059669");
        };
        const excluirMp = async (mp) => {
          const n = usosDaMp(mp.id);
          const aviso = n>0 ? `"${mp.nome}" é usada em ${n} vínculo(s) de produtos. Eles continuam calculando o custo, mas ela não aparecerá mais para novos vínculos. Remover?` : `Remover "${mp.nome}" da lista?`;
          if (!window.confirm(aviso)) return;
          await excluirMateriaPrimaLib(mp.id);
          logOp({tipo:"produtos",acao:"excluir",descricao:"Matéria-prima removida: "+mp.nome,tabela:"materias_primas",registro_id:mp.id,antes:pick(mp,COLS_MP)});
          await reloadMps();
          showToast("🗑️ Removida.","#64748b");
        };

        // ── Modais compartilhados ───────────────────────────────────────────
        const mpLibModal = showMpLib&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:680,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,gap:10,flexWrap:"wrap"}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:700,color:"#1e293b"}}>🧱 Matérias-primas</h3>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>abrirMpForm()} style={{background:"#0d9488",color:"#fff",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>+ Nova matéria-prima</button>
                  <button onClick={()=>abrirKitForm()} style={{background:"#f0fdfa",color:"#0d9488",border:"1.5px solid #99f6e4",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🧰 Novo kit</button>
                  <button onClick={()=>setShowMpLib(false)} style={{background:"#f1f5f9",color:"#374151",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Fechar</button>
                </div>
              </div>
              <p style={{margin:"0 0 12px",fontSize:12,color:"#64748b"}}>Cadastre aqui uma vez e vincule a quantos produtos e variações quiser.</p>
              {kits.length>0&&(
                <div style={{marginBottom:14}}>
                  <div style={{fontSize:12,fontWeight:700,color:"#0d9488",marginBottom:6}}>🧰 Kits</div>
                  <div style={{display:"flex",flexDirection:"column",gap:6}}>
                    {kits.map(k=>(
                      <div key={k.id} style={{display:"flex",alignItems:"center",gap:8,background:"#f0fdfa",border:"1px solid #99f6e4",borderRadius:10,padding:"8px 12px",flexWrap:"wrap"}}>
                        <div style={{flex:1,minWidth:160}}>
                          <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>{k.nome}</div>
                          <div style={{fontSize:10,color:"#64748b"}}>{(k.itens||[]).map(i=>`${materiasPrimas.find(x=>x.id===i.mpId)?.nome||"?"}${Number(i.fator)!==1?" ×"+i.fator:""}`).join(" + ")}</div>
                        </div>
                        <button onClick={()=>abrirKitForm(k)} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:6,padding:"4px 8px",fontSize:11,cursor:"pointer"}}>✏️</button>
                        <button onClick={()=>excluirKitLista(k)} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:6,padding:"4px 8px",fontSize:11,cursor:"pointer"}}>🗑️</button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {materiasPrimas.length===0 ? (
                <div style={{textAlign:"center",padding:"28px 0",color:"#94a3b8",fontSize:13}}>Nenhuma matéria-prima cadastrada ainda.</div>
              ) : (
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr>{["Material","Custo da compra","Qtd na compra","Custo unitário","Usada em",""].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                    <tbody>
                      {materiasPrimas.map(mp=>(
                        <tr key={mp.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                          <td style={TD}><strong>{mp.nome}</strong></td>
                          <td style={TD}>{fmtR(mp.custo_unidade_compra)}</td>
                          <td style={TD}>{mp.qtd_por_unidade_compra} {mp.unidade}</td>
                          <td style={{...TD,fontWeight:700,color:"#f59e0b"}}>R$ {custoUnitMp(mp).toFixed(2)}/{mp.unidade}</td>
                          <td style={TD}>{usosDaMp(mp.id)} vínculo(s)</td>
                          <td style={TD}>
                            <div style={{display:"flex",gap:6}}>
                              <button onClick={()=>abrirMpForm(mp)} style={{background:"#f1f5f9",border:"none",borderRadius:6,padding:"4px 8px",fontSize:11,cursor:"pointer"}}>✏️</button>
                              <button onClick={()=>excluirMp(mp)} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:6,padding:"4px 8px",fontSize:11,cursor:"pointer"}}>🗑️</button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        );

        const produtoFormModal = showProdutoForm&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:700,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <h3 style={{margin:"0 0 14px",fontSize:16,fontWeight:700,color:"#1e293b"}}>{editandoProdutoCalc?"✏️ Editar produto":"+ Novo Produto"}</h3>
              <div style={{display:"flex",flexDirection:"column",gap:12}}>
                <div>
                  <label style={lbl}>Nome do produto *</label>
                  <input value={prodNome} onChange={e=>setProdNome(e.target.value)} placeholder="Ex: Mini Fotos" style={inp} />
                </div>
                <div style={{background:"#f8fafc",borderRadius:10,padding:"12px 14px"}}>
                  <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:8}}>Taxas da plataforma (valem para todas as variações)</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:8}}>
                    {[["Comissão (%)",prodTaxaPlat,setProdTaxaPlat],["Imposto (%)",prodTaxaImp,setProdTaxaImp],["Adiantamento (%)",prodTaxaAdiant,setProdTaxaAdiant],["Taxa fixa (R$)",prodTaxaFixa,setProdTaxaFixa]].map(([l,val,setter])=>(
                      <div key={l}>
                        <label style={{fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2}}>{l}</label>
                        <input type="number" min="0" step="0.01" value={val} onChange={e=>setter(e.target.value)} style={{width:"100%",padding:"7px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                    ))}
                  </div>
                </div>

                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:4}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>🎛️ Variações</div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>abrirMpForm()} style={{background:"#f0fdfa",color:"#0d9488",border:"1px solid #99f6e4",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>🧱 + Nova matéria-prima</button>
                    <button onClick={addVar} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:8,padding:"5px 12px",fontSize:11,fontWeight:700,cursor:"pointer"}}>+ Variação</button>
                  </div>
                </div>
                <div style={{fontSize:11,color:"#94a3b8",marginTop:-6}}>Ex: "9 unidades", "18 unidades com imã". Cada variação tem seu preço e suas matérias-primas.</div>

                {prodVars.map((v,vi)=>{
                  const custo = custoEditor(v);
                  const liq = liquidoPreco(pEdit, num(v.preco,0));
                  const marg = margemDe(liq, custo);
                  return (
                    <div key={v.key} style={{border:"1.5px solid #e9d5ff",background:"#faf5ff",borderRadius:14,padding:"12px 14px"}}>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 130px",gap:8,marginBottom:10}}>
                        <div>
                          <label style={{fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2}}>Nome da variação {prodVars.length>1?"*":"(opcional)"}</label>
                          <input value={v.nome} onChange={e=>setVar(v.key,{nome:e.target.value})} placeholder={`Ex: ${vi===0?"9 unidades":"18 unidades com imã"}`} style={{...inp,padding:"8px 10px"}} />
                        </div>
                        <div>
                          <label style={{fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2}}>Preço de venda (R$)</label>
                          <input type="number" min="0" step="0.01" value={v.preco} onChange={e=>setVar(v.key,{preco:e.target.value})} placeholder="15.00" style={{...inp,padding:"8px 10px"}} />
                        </div>
                      </div>

                      <div style={{fontSize:10,fontWeight:700,color:"#7c3aed",marginBottom:4}}>MATÉRIAS-PRIMAS DESTA VARIAÇÃO</div>
                      {materiasPrimas.length===0&&<div style={{fontSize:11,color:"#f59e0b",marginBottom:6}}>Nenhuma matéria-prima cadastrada. Use o botão "🧱 + Nova matéria-prima" acima.</div>}
                      <div style={{display:"flex",flexDirection:"column",gap:6}}>
                        {v.mps.map((m,mi)=>{
                          if (m.kitId!==undefined) {
                            const kit = kits.find(k=>k.id===m.kitId);
                            const custoKit = kit ? expandirMps([m]).reduce((a,x)=>{const mp2=materiasPrimas.find(y=>y.id===x.mpId);return mp2?a+custoUnitMp(mp2)*x.qtd:a;},0) : 0;
                            return (
                              <div key={mi} style={{background:"#f0fdfa",border:"1.5px solid #99f6e4",borderRadius:10,padding:"8px 8px"}}>
                                <div style={{display:"grid",gridTemplateColumns:"1fr 90px 30px",gap:6,alignItems:"center"}}>
                                  <select value={m.kitId} onChange={e=>setVarMp(v.key,mi,{kitId:e.target.value})}
                                    style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                                    <option value="">🧰 Escolher kit…</option>
                                    {kits.map(k=><option key={k.id} value={k.id}>{k.nome}</option>)}
                                  </select>
                                  <input type="number" min="0" step="any" value={m.qtd} onChange={e=>setVarMp(v.key,mi,{qtd:e.target.value})} placeholder="qtd" style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",boxSizing:"border-box"}} />
                                  <button onClick={()=>delVarMp(v.key,mi)} title="Remover" style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:6,height:30,cursor:"pointer",fontSize:12}}>✕</button>
                                </div>
                                {kit&&<div style={{fontSize:10,color:"#0d9488",marginTop:4}}>Entra: {(kit.itens||[]).map(i=>materiasPrimas.find(x=>x.id===i.mpId)?.nome||"?").join(" + ")}{Number(m.qtd)>0?` · → R$ ${custoKit.toFixed(2)}`:""}</div>}
                              </div>
                            );
                          }
                          const mp = materiasPrimas.find(x=>x.id===m.mpId);
                          const custoLin = mp ? custoUnitMp(mp)*(Number(m.qtd)||0) : 0;
                          return (
                            <div key={mi}>
                              <div style={{display:"grid",gridTemplateColumns:"1fr 90px 30px",gap:6,alignItems:"center"}}>
                                <select value={m.mpId} onChange={e=>setVarMp(v.key,mi,{mpId:e.target.value})}
                                  style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                                  <option value="">Escolher matéria-prima…</option>
                                  {materiasPrimas.map(x=><option key={x.id} value={x.id}>{x.nome} (R$ {custoUnitMp(x).toFixed(2)}/{x.unidade})</option>)}
                                </select>
                                <input type="number" min="0" step="any" value={m.qtd} onChange={e=>setVarMp(v.key,mi,{qtd:e.target.value})} placeholder={mp?`qtd (${mp.unidade})`:"qtd"}
                                  style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",boxSizing:"border-box"}} />
                                <button onClick={()=>delVarMp(v.key,mi)} title="Remover" style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:6,height:30,cursor:"pointer",fontSize:12}}>✕</button>
                              </div>
                              {mp&&Number(m.qtd)>0&&<div style={{fontSize:10,color:"#f59e0b",fontWeight:600,margin:"2px 0 0 4px"}}>→ R$ {custoLin.toFixed(2)}</div>}
                            </div>
                          );
                        })}
                      </div>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:10,flexWrap:"wrap",gap:8}}>
                        <div style={{display:"flex",gap:6}}>
                          <button onClick={()=>addVarMp(v.key)} style={{background:"#fff",color:"#7c3aed",border:"1.5px solid #ddd6fe",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>+ Matéria-prima</button>
                          <button onClick={()=>addVarKit(v.key)} style={{background:"#f0fdfa",color:"#0d9488",border:"1.5px solid #99f6e4",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>🧰 + Kit</button>
                          <button onClick={()=>dupVar(v.key)} style={{background:"#fff",color:"#64748b",border:"1.5px solid #e2e8f0",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>⧉ Duplicar</button>
                          {prodVars.length>1&&<button onClick={()=>delVar(v.key)} style={{background:"#fff",color:"#991b1b",border:"1.5px solid #fecaca",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>🗑️ Remover variação</button>}
                        </div>
                        <div style={{fontSize:11,color:"#374151"}}>
                          Custo <strong style={{color:"#f59e0b"}}>{fmtR(custo)}</strong> · Recebe <strong style={{color:"#059669"}}>{fmtR(liq)}</strong> · Margem <strong style={{color:corMargem(marg)}}>{fmtM(marg)}</strong>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
              <div style={{display:"flex",gap:10,marginTop:18}}>
                <button onClick={()=>setShowProdutoForm(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                <button onClick={salvarProdutoCalc} disabled={!prodNome.trim()||salvandoProdutoCalc}
                  style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:prodNome.trim()?"#7c3aed":"#e2e8f0",color:prodNome.trim()?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  {salvandoProdutoCalc?"Salvando...":editandoProdutoCalc?"Salvar alterações":"Criar produto"}
                </button>
              </div>
            </div>
          </div>
        );

        const mpFormModal = showMpForm&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10001,display:"flex",alignItems:"center",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:500,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <h3 style={{margin:"0 0 14px",fontSize:15,fontWeight:700}}>{editandoMp?"✏️ Editar matéria-prima":"+ Nova matéria-prima"}</h3>
              <div style={{display:"flex",flexDirection:"column",gap:11}}>
                {!editandoMp&&estoqueProdutos.filter(x=>x.ativo!==false).length>0&&(
                  <div style={{background:"#f0fdfa",border:"1.5px solid #99f6e4",borderRadius:12,padding:"10px 12px"}}>
                    <label style={{fontSize:11,fontWeight:700,color:"#0d9488",display:"block",marginBottom:4}}>📦 Puxar do Estoque (preenche nome, custo e unidade)</label>
                    <select value={mpEstoqueId} onChange={e=>puxarDoEstoque(e.target.value)}
                      style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                      <option value="">Escolher item do estoque…</option>
                      {estoqueProdutos.filter(x=>x.ativo!==false).map(x=>{
                        const ja = materiasPrimas.some(m=>m.estoque_produto_id===x.id);
                        return <option key={x.id} value={x.id}>{x.nome} ({x.unidade}){x.preco_unitario?` · R$ ${Number(x.preco_unitario).toFixed(2)}`:""}{ja?" ✓ já cadastrada":""}</option>;
                      })}
                    </select>
                    <div style={{fontSize:10,color:"#64748b",marginTop:4}}>Depois é só ajustar quanto 1 unidade do estoque rende e as demais características.</div>
                  </div>
                )}
                <div>
                  <label style={lbl}>Nome do material *</label>
                  <input value={mpNome} onChange={e=>setMpNome(e.target.value)} placeholder="Ex: Manta Imã, Papel micro poroso, Envelope..." style={{...inp,padding:"9px 12px"}} />
                </div>
                <div>
                  <label style={lbl}>Custo da unidade de compra (R$) *</label>
                  <input type="number" min="0" step="0.01" value={mpCustoCompra} onChange={e=>setMpCustoCompra(e.target.value)} placeholder="Ex: 700,00" style={inp} />
                  <span style={{fontSize:10,color:"#94a3b8"}}>Quanto você paga no rolo/caixa/pacote inteiro</span>
                </div>
                <button onClick={()=>setMpModoDimensao(v=>!v)}
                  style={{background:mpModoDimensao?"#ede9fe":"#f1f5f9",border:`1.5px solid ${mpModoDimensao?"#7c3aed":"#e2e8f0"}`,borderRadius:10,padding:"8px 14px",fontSize:12,fontWeight:600,cursor:"pointer",color:mpModoDimensao?"#7c3aed":"#64748b",display:"flex",alignItems:"center",gap:8,width:"100%",justifyContent:"center"}}>
                  📐 {mpModoDimensao?"Ocultar calculadora de rolo/folha":"Calcular quantas peças saem de um rolo ou folha grande"}
                </button>
                {mpModoDimensao&&(
                  <div style={{background:"#faf5ff",border:"1.5px solid #e9d5ff",borderRadius:12,padding:"14px 16px"}}>
                    <div style={{fontSize:12,fontWeight:700,color:"#7c3aed",marginBottom:10}}>📐 Calculadora de corte</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                      {[["Largura do rolo/folha (cm)",mpRoloLargura,setMpRoloLargura,"Ex: 64"],["Comprimento do rolo (metros)",mpRoloComprimento,setMpRoloComprimento,"Ex: 40"],["Largura do corte (cm)",mpCorteL,setMpCorteL,"Ex: 21 (A4)"],["Altura do corte (cm)",mpCorteC,setMpCorteC,"Ex: 29.7 (A4)"]].map(([l,val,setter,ph])=>(
                        <div key={l}>
                          <label style={{fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2}}>{l}</label>
                          <input type="number" min="0" step="0.1" value={val} onChange={e=>setter(e.target.value)} placeholder={ph}
                            style={{width:"100%",padding:"7px 10px",border:"1.5px solid #ddd6fe",borderRadius:8,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                        </div>
                      ))}
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                      <label style={{fontSize:10,fontWeight:600,color:"#64748b"}}>Nome da peça cortada:</label>
                      <select value={mpUnidadeCorte} onChange={e=>setMpUnidadeCorte(e.target.value)}
                        style={{padding:"5px 8px",border:"1.5px solid #ddd6fe",borderRadius:8,fontSize:12,outline:"none",background:"#fff"}}>
                        {["folha","pedaço","un","tira","quadrado","retângulo"].map(u=><option key={u} value={u}>{u}</option>)}
                      </select>
                    </div>
                    {(()=>{
                      const opcoes = calcDimensao();
                      if (!opcoes) return <div style={{fontSize:11,color:"#94a3b8",textAlign:"center",padding:"6px 0"}}>Preencha as dimensões acima para ver o resultado</div>;
                      return (
                        <div style={{display:"flex",flexDirection:"column",gap:6}}>
                          <div style={{fontSize:10,fontWeight:700,color:"#7c3aed",marginBottom:2}}>💡 Sugestões de aproveitamento:</div>
                          {opcoes.map((o,i)=>(
                            <div key={i} style={{background: i===0?"#f5f3ff":"#fff", border:`1.5px solid ${i===0?"#7c3aed":"#e9d5ff"}`,borderRadius:9,padding:"9px 11px"}}>
                              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:3}}>
                                <span style={{fontSize:11,fontWeight:700,color: i===0?"#6d28d9":"#374151"}}>{i===0?"🥇 ":i===1?"🥈 ":"🥉 "}{o.label}</span>
                                <span style={{fontSize:10,background: o.aproveitamento>=90?"#d1fae5":o.aproveitamento>=70?"#fef3c7":"#fee2e2",color: o.aproveitamento>=90?"#065f46":o.aproveitamento>=70?"#92400e":"#991b1b",borderRadius:6,padding:"1px 7px",fontWeight:700}}>{o.aproveitamento}% aproveit.</span>
                              </div>
                              <div style={{fontSize:11,color:"#374151"}}>
                                {o.colunas} col × {o.linhas} lin = <strong style={{color:"#7c3aed"}}>{o.total} {mpUnidadeCorte}s</strong>
                                {mpCustoCompra&&<span style={{color:"#059669",fontWeight:700}}> · R$ {o.custo_cada.toFixed(2)}/{mpUnidadeCorte}</span>}
                              </div>
                              <div style={{fontSize:10,color:"#94a3b8",marginTop:1}}>Sobra: {o.sobra_larg}cm na largura · {o.sobra_comp}cm no comprimento</div>
                              <button onClick={()=>aplicarDimensao(i)}
                                style={{marginTop:6,padding:"5px 12px",border:"none",borderRadius:7,background:i===0?"#7c3aed":"#6d28d9",color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer",width:"100%"}}>
                                ✅ Usar este ({o.total} {mpUnidadeCorte}s)
                              </button>
                            </div>
                          ))}
                        </div>
                      );
                    })()}
                  </div>
                )}
                <div>
                  <label style={lbl}>Qtd que vem nessa compra *</label>
                  <div style={{display:"flex",gap:6,alignItems:"center"}}>
                    <input type="number" min="0.001" step="any" value={mpQtdPorCompra} onChange={e=>setMpQtdPorCompra(e.target.value)} placeholder="Ex: 402 folhas A4"
                      style={{flex:1,padding:"9px 10px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    <select value={mpUnidade} onChange={e=>setMpUnidade(e.target.value)}
                      style={{padding:"9px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:12,outline:"none",background:"#fff"}}>
                      {["folha","pedaço","un","cx","pct","rolo","m","cm","kg","g","l","ml"].map(u=><option key={u} value={u}>{u}</option>)}
                    </select>
                  </div>
                  {mpCustoCompra&&mpQtdPorCompra&&<span style={{fontSize:10,color:"#059669",fontWeight:600}}>→ R$ {(Number(mpCustoCompra)/Number(mpQtdPorCompra)).toFixed(2)} por {mpUnidade}</span>}
                </div>
                <div style={{background:"#f0fdfa",border:"1.5px solid #99f6e4",borderRadius:12,padding:"12px 14px"}}>
                  <div style={{fontSize:12,fontWeight:700,color:"#0d9488",marginBottom:6}}>📦 Baixa automática no estoque (opcional)</div>
                  <label style={{fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2}}>Item do Estoque que será descontado</label>
                  <select value={mpEstoqueId} onChange={e=>setMpEstoqueId(e.target.value)}
                    style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                    <option value="">Não descontar do estoque</option>
                    {estoqueProdutos.filter(x=>x.ativo!==false).map(x=><option key={x.id} value={x.id}>{x.nome} ({x.unidade})</option>)}
                  </select>
                  {mpEstoqueId&&(()=>{
                    const ep = estoqueProdutos.find(x=>x.id===mpEstoqueId);
                    return (
                      <div style={{marginTop:8}}>
                        <label style={{fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2}}>Rendimento: 1 {ep?.unidade||"un"} do estoque rende quantos(as) {mpUnidade}?</label>
                        <input type="number" min="0" step="any" value={mpRendimento} onChange={e=>setMpRendimento(e.target.value)} placeholder={mpQtdPorCompra?`Vazio = ${mpQtdPorCompra} (a qtd da compra)`:"Ex: 5000"}
                          style={{width:"100%",padding:"8px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                        <div style={{fontSize:10,color:"#64748b",marginTop:4}}>Ex: tinta que rende 5000 impressões A4 → 1 frasco = 5000 folha. Cada pedido desconta (folhas usadas ÷ rendimento) do estoque.</div>
                      </div>
                    );
                  })()}
                </div>
              </div>
              <div style={{display:"flex",gap:10,marginTop:18}}>
                <button onClick={()=>setShowMpForm(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                <button onClick={salvarMp} disabled={!mpNome.trim()||!mpCustoCompra||!mpQtdPorCompra||salvandoMp}
                  style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:(mpNome.trim()&&mpCustoCompra&&mpQtdPorCompra)?"#0d9488":"#e2e8f0",color:(mpNome.trim()&&mpCustoCompra&&mpQtdPorCompra)?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  {salvandoMp?"Salvando...":"Salvar matéria-prima"}
                </button>
              </div>
            </div>
          </div>
        );

        const opcoesVar = [];
        produtosCalc.forEach(pr=>(pr.produto_variacoes||[]).forEach(v=>opcoesVar.push({id:v.id,nome:`${pr.nome}${v.nome?" — "+v.nome:""}`})));
        const vincModal = showVinc&&(()=>{
          const q = _norm(vincBusca);
          const lista = vincLista.filter(x=>(vincSo==="todas"||(!x.variacao_id&&!x.ignorar))&&(!q||_norm(x.variacao).includes(q)));
          const pend = vincLista.filter(x=>!x.variacao_id&&!x.ignorar).length;
          return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:820,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap",marginBottom:10}}>
                <h3 style={{margin:0,fontSize:16,fontWeight:700,color:"#1e293b"}}>🔗 Vincular variações do UpSeller</h3>
                <button onClick={()=>setShowVinc(false)} style={{background:"#f1f5f9",color:"#374151",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:600,cursor:"pointer"}}>Fechar</button>
              </div>
              <p style={{margin:"0 0 12px",fontSize:12,color:"#64748b"}}>Vincule cada parte que aparece nos pedidos a uma variação de produto cadastrada (pedidos com “A + B” viram duas partes). Depois disso, todo pedido novo com essa variação que chegar em “Para imprimir” ou “Para retirada” desconta do estoque as matérias-primas do produto (usando o rendimento de cada material).</p>
              <div style={{background:baixaInfo?.ativo?"#ecfdf5":"#fffbeb",border:`1px solid ${baixaInfo?.ativo?"#a7f3d0":"#fde68a"}`,borderRadius:12,padding:"12px 14px",marginBottom:12}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
                  <div style={{fontSize:13,fontWeight:700,color:baixaInfo?.ativo?"#065f46":"#92400e"}}>{baixaInfo?.ativo?"🟢 Baixa automática ATIVA":"⏸️ Baixa automática desativada"}</div>
                  <div style={{display:"flex",gap:8}}>
                    {baixaInfo?.ativo&&<button disabled={baixaBusy} onClick={rodarBaixa} style={{background:"#0d9488",color:"#fff",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>{baixaBusy?"...":"▶ Processar agora"}</button>}
                    <button disabled={baixaBusy} onClick={alternarBaixa} style={{background:"#fff",color:"#374151",border:"1.5px solid #e2e8f0",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>{baixaInfo?.ativo?"Pausar":"Ativar"}</button>
                  </div>
                </div>
                {baixaInfo?.desde&&<div style={{fontSize:11,color:"#475569",marginTop:6}}>Vale para pedidos feitos a partir de {new Date(baixaInfo.desde).toLocaleString("pt-BR")} · {baixaInfo.processados||0} já baixado(s){baixaInfo.estornados>0?` · ${baixaInfo.estornados} estornado(s) por cancelamento`:""}{baixaInfo.sem_vinculo>0?` · ${baixaInfo.sem_vinculo} pedido(s) com variação sem vínculo`:""}{baixaInfo.sem_material>0?` · ${baixaInfo.sem_material} aguardando processar`:""}</div>}
                <div style={{fontSize:10,color:"#94a3b8",marginTop:4}}>A baixa só acontece para produtos cuja matéria-prima está ligada a um item do Estoque (em 🧱 Matérias-primas). Se um pedido já baixado for cancelado, o material volta sozinho para o estoque. Cada baixa pode ser desfeita no 🕘 Histórico.</div>
              </div>
              <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
                <input value={vincBusca} onChange={e=>setVincBusca(e.target.value)} placeholder="Buscar variação…" style={{...inp,flex:1,minWidth:160}} />
                <select value={vincSo} onChange={e=>setVincSo(e.target.value)} style={{padding:"8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:12,background:"#fff"}}>
                  <option value="pendentes">Sem vínculo ({pend})</option>
                  <option value="todas">Todas ({vincLista.length})</option>
                </select>
              </div>
              {vincCarregando ? <div style={{textAlign:"center",padding:24,color:"#94a3b8",fontSize:13}}>Carregando…</div>
              : lista.length===0 ? <div style={{textAlign:"center",padding:24,color:"#94a3b8",fontSize:13}}>{vincLista.length===0?"Nenhuma variação encontrada nos pedidos.":"Nada por aqui 🎉"}</div>
              : (
                <div style={{display:"flex",flexDirection:"column",gap:8,maxHeight:"55vh",overflowY:"auto"}}>
                  {lista.map(it=>{
                    const sug = !it.variacao_id&&!it.ignorar ? melhorMatch(it.variacao, opcoesVar, 0.45) : null;
                    return (
                      <div key={it.chave} style={{border:"1px solid #e2e8f0",borderRadius:12,padding:"10px 12px",background:(it.variacao_id||it.ignorar)?"#f0fdf4":"#fff"}}>
                        <div style={{fontSize:12,fontWeight:600,color:"#1e293b",wordBreak:"break-word"}}>{it.variacao}</div>
                        <div style={{fontSize:10,color:"#94a3b8",margin:"2px 0 6px"}}>{it.pedidos} pedido(s) · {it.unidades} un · último {new Date(it.ultimo).toLocaleDateString("pt-BR")}</div>
                        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                          <select value={it.ignorar?"__ignorar":(it.variacao_id||"")} onChange={e=>escolherVinculo(it,e.target.value)}
                            style={{flex:1,minWidth:200,padding:"7px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",background:"#fff"}}>
                            <option value="">— Sem vínculo —</option>
                            <option value="__ignorar">🚫 Sem material (não dá baixa)</option>
                            {opcoesVar.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
                          </select>
                          {sug&&<button onClick={()=>escolherVinculo(it,sug.id)} style={{background:"#f5f3ff",color:"#6d28d9",border:"1px solid #ddd6fe",borderRadius:8,padding:"6px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>💡 {sug.nome}</button>}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
          );
        })();

        const kitFormModal = showKitForm&&(
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10002,display:"flex",alignItems:"center",justifyContent:"center",padding:16,overflowY:"auto"}}>
            <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:520,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",margin:"auto"}}>
              <h3 style={{margin:"0 0 6px",fontSize:15,fontWeight:700}}>{editandoKit?"✏️ Editar kit":"🧰 Novo kit"}</h3>
              <p style={{margin:"0 0 12px",fontSize:11,color:"#64748b"}}>Junte matérias-primas que sempre são usadas juntas (ex: as 4 cores de tinta). No produto você liga o kit uma vez, com a quantidade, e cada material entra com seu custo e sua baixa de estoque.</p>
              <label style={lbl}>Nome do kit *</label>
              <input value={kitNome} onChange={e=>setKitNome(e.target.value)} placeholder="Ex: Tinta 4 cores" style={{...inp,marginBottom:12}} />
              <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:6}}>Matérias-primas do kit (e quanto de cada uma entra em 1 unidade do kit)</div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {kitItens.map((it,i)=>(
                  <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 80px 30px",gap:6,alignItems:"center"}}>
                    <select value={it.mpId} onChange={e=>setKitItens(l=>l.map((x,j)=>j===i?{...x,mpId:e.target.value}:x))}
                      style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                      <option value="">Escolher matéria-prima…</option>
                      {materiasPrimas.map(x=><option key={x.id} value={x.id}>{x.nome}</option>)}
                    </select>
                    <input type="number" min="0" step="any" value={it.fator} onChange={e=>setKitItens(l=>l.map((x,j)=>j===i?{...x,fator:e.target.value}:x))} title="Quanto entra por unidade do kit" style={{width:"100%",padding:"8px 8px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none",boxSizing:"border-box"}} />
                    <button onClick={()=>setKitItens(l=>l.length>1?l.filter((_,j)=>j!==i):[{mpId:"",fator:"1"}])} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:6,height:30,cursor:"pointer",fontSize:12}}>✕</button>
                  </div>
                ))}
              </div>
              <button onClick={()=>setKitItens(l=>[...l,{mpId:"",fator:"1"}])} style={{marginTop:8,background:"#fff",color:"#0d9488",border:"1.5px solid #99f6e4",borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>+ Adicionar matéria-prima</button>
              <div style={{display:"flex",gap:10,marginTop:18}}>
                <button onClick={()=>setShowKitForm(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                <button onClick={salvarKitForm} disabled={salvandoKit} style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:"#0d9488",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>{salvandoKit?"Salvando…":"Salvar kit"}</button>
              </div>
            </div>
          </div>
        );
        const modais = <>{mpLibModal}{produtoFormModal}{mpFormModal}{kitFormModal}{vincModal}</>;

        // ── Detalhe do produto ──────────────────────────────────────────────
        if (produtoAtivo) {
          const p = produtoAtivo;
          const vars = p.produto_variacoes||[];
          const mm = margemMedia(p);
          return (
            <div>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,flexWrap:"wrap"}}>
                <button onClick={()=>setProdutoAtivo(null)} style={{background:"#f1f5f9",border:"none",borderRadius:10,padding:"8px 14px",fontSize:13,cursor:"pointer",color:"#374151",fontWeight:600}}>← Voltar</button>
                <h2 style={{margin:0,fontSize:17,fontWeight:800,color:"#1e293b",flex:1}}>{p.nome}</h2>
                <button onClick={()=>setShowMpLib(true)} style={{background:"#f0fdfa",border:"1px solid #99f6e4",borderRadius:10,padding:"8px 14px",fontSize:12,cursor:"pointer",color:"#0d9488",fontWeight:700}}>🧱 Matérias-primas</button>
                <button onClick={abrirVinc} style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:10,padding:"8px 14px",fontSize:12,cursor:"pointer",color:"#1d4ed8",fontWeight:700}}>🔗 Vincular UpSeller</button>
                <button onClick={()=>abrirFormEditarProd(p)} style={{background:"#f1f5f9",border:"none",borderRadius:10,padding:"8px 14px",fontSize:12,cursor:"pointer",color:"#374151"}}>✏️ Editar produto</button>
              </div>

              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:12,marginBottom:16}}>
                <div style={{background:bgMargem(mm),borderRadius:12,padding:"12px 14px",border:"1px solid #e2e8f0"}}>
                  <div style={{fontSize:10,fontWeight:600,color:"#64748b",marginBottom:4}}>Margem de lucro média</div>
                  <div style={{fontSize:20,fontWeight:800,color:corMargem(mm)}}>{fmtM(mm)}</div>
                </div>
                <div style={{background:"#f8fafc",borderRadius:12,padding:"12px 14px",border:"1px solid #e2e8f0"}}>
                  <div style={{fontSize:10,fontWeight:600,color:"#64748b",marginBottom:4}}>Variações</div>
                  <div style={{fontSize:20,fontWeight:800,color:"#1e293b"}}>{vars.length}</div>
                </div>
                <div style={{background:"#fff7ed",borderRadius:12,padding:"12px 14px",border:"1px solid #e2e8f0"}}>
                  <div style={{fontSize:10,fontWeight:600,color:"#64748b",marginBottom:4}}>Taxas da plataforma</div>
                  <div style={{fontSize:13,fontWeight:700,color:"#991b1b"}}>{pctTaxas(p)}% + {fmtR(p.taxa_fixa)}</div>
                  <div style={{fontSize:9,color:"#94a3b8",marginTop:2}}>comissão {num(p.taxa_plataforma_pct,0)}% · imposto {num(p.taxa_imposto_pct,0)}% · adiant. {num(p.taxa_adiantamento_pct,0)}%</div>
                </div>
              </div>

              {vars.map(v=>{
                const custo = custoVar(v);
                const liq = liquidoVar(p,v);
                const marg = margemVar(p,v);
                const linhas = v.produto_materias_primas||[];
                const taxasR = Number(v.preco_venda) - liq;
                return (
                  <div key={v.id} style={{background:"#fff",borderRadius:14,padding:"16px 18px",marginBottom:14,boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
                      <div style={{fontSize:14,fontWeight:800,color:"#1e293b"}}>🎛️ {v.nome}</div>
                      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                        <span title="Quanto do preço de venda vira lucro" style={{fontSize:12,fontWeight:800,color:corMP(margemPreco(p,v)),background:"#f1f5f9",borderRadius:8,padding:"3px 10px"}}>Margem s/ preço {fmtM(margemPreco(p,v))}</span>
                        <span title="Valor recebido ÷ custo" style={{fontSize:12,fontWeight:800,color:corMargem(marg),background:bgMargem(marg),borderRadius:8,padding:"3px 10px"}}>Retorno s/ custo {fmtM(marg)}</span>
                      </div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:10,marginBottom:12}}>
                      {[
                        {label:"Preço de venda",value:fmtR(v.preco_venda),color:"#1e293b",bg:"#f8fafc"},
                        {label:"Taxas",value:"- "+fmtR(taxasR),color:"#991b1b",bg:"#fde8e8"},
                        {label:"Você recebe",value:fmtR(liq),color:"#059669",bg:"#d1fae5"},
                        {label:"Custo de produção",value:fmtR(custo),color:"#f59e0b",bg:"#fffbeb"},
                        {label:"Lucro por unidade",value:fmtR(liq-custo),color:liq-custo>=0?"#0d9488":"#ef4444",bg:liq-custo>=0?"#f0fdfa":"#fde8e8"},
                      ].map(c=>(
                        <div key={c.label} style={{background:c.bg,borderRadius:10,padding:"10px 12px",border:"1px solid #e2e8f0"}}>
                          <div style={{fontSize:10,fontWeight:600,color:"#64748b",marginBottom:3}}>{c.label}</div>
                          <div style={{fontSize:15,fontWeight:800,color:c.color}}>{c.value}</div>
                        </div>
                      ))}
                    </div>
                    {linhas.length===0 ? (
                      <div style={{textAlign:"center",padding:"12px 0",color:"#94a3b8",fontSize:12}}>Nenhuma matéria-prima vinculada. Clique em ✏️ Editar produto.</div>
                    ) : (
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead><tr>{["Material","Custo unitário","Qtd usada","Custo"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                          <tbody>
                            {linhas.map(l=>{
                              const m = dadosMp(l);
                              return (
                                <tr key={l.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                                  <td style={TD}><strong>{m.nome}</strong></td>
                                  <td style={TD}>R$ {custoUnitMp(m).toFixed(2)}/{m.unidade}</td>
                                  <td style={TD}>{l.qtd_usada_por_produto} {m.unidade}</td>
                                  <td style={{...TD,fontWeight:700,color:"#f59e0b"}}>R$ {custoLinha(l).toFixed(2)}</td>
                                </tr>
                              );
                            })}
                            <tr style={{background:"#f8fafc",fontWeight:700}}>
                              <td colSpan={3} style={{...TD,color:"#374151"}}>Total custo de produção</td>
                              <td style={{...TD,fontSize:14,color:"#f59e0b"}}>R$ {custo.toFixed(2)}</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                );
              })}
              {modais}
            </div>
          );
        }

        // ── Lista de produtos ───────────────────────────────────────────────
        return (
          <div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
              <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#1e293b"}}>📦 Produtos</h2>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={()=>setShowMpLib(true)} style={{background:"#f0fdfa",color:"#0d9488",border:"1.5px solid #99f6e4",borderRadius:10,padding:"9px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  🧱 Matérias-primas
                </button>
                <button onClick={abrirVinc} style={{background:"#eff6ff",color:"#1d4ed8",border:"1.5px solid #bfdbfe",borderRadius:10,padding:"9px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  🔗 Vincular UpSeller
                </button>
                <button onClick={abrirFormNovoProd} style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",color:"#fff",border:"none",borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  + Novo Produto
                </button>
              </div>
            </div>

            {produtosCalc.length===0 ? (
              <div style={{background:"#fff",borderRadius:16,padding:40,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:40,marginBottom:8}}>📦</div>
                <div style={{fontSize:16,fontWeight:700,color:"#1e293b",marginBottom:6}}>Nenhum produto cadastrado</div>
                <div style={{fontSize:13,color:"#64748b",marginBottom:20}}>Cadastre seus produtos para calcular custo, margem e lucro automaticamente.</div>
                <button onClick={abrirFormNovoProd} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:10,padding:"10px 22px",fontSize:13,fontWeight:700,cursor:"pointer"}}>+ Cadastrar Primeiro Produto</button>
              </div>
            ) : (
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(340px,1fr))",gap:14}}>
                {produtosCalc.map(p=>{
                  const vars = p.produto_variacoes||[];
                  const mm = margemMedia(p);
                  return (
                    <div key={p.id} onClick={()=>setProdutoAtivo(p)} style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",cursor:"pointer",border:"2px solid #f1f5f9",transition:"border-color 0.15s"}}
                      onMouseEnter={e=>e.currentTarget.style.borderColor="#7c3aed"}
                      onMouseLeave={e=>e.currentTarget.style.borderColor="#f1f5f9"}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:8}}>
                        <div style={{fontSize:15,fontWeight:700,color:"#1e293b"}}>{p.nome}</div>
                        <span style={{fontSize:10,fontWeight:700,color:"#7c3aed",background:"#f5f3ff",borderRadius:6,padding:"2px 8px",whiteSpace:"nowrap"}}>{vars.length} variaç{vars.length===1?"ão":"ões"}</span>
                      </div>
                      <div style={{overflowX:"auto",marginBottom:10}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                          <thead><tr>{["Variação","Preço","Custo","Recebe","Retorno","Margem s/ preço"].map(h=><th key={h} style={{...TH,padding:"5px 6px"}}>{h}</th>)}</tr></thead>
                          <tbody>
                            {vars.map(v=>{
                              const marg = margemVar(p,v);
                              return (
                                <tr key={v.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                                  <td style={{padding:"6px",fontWeight:600,color:"#1e293b"}}>{v.nome}</td>
                                  <td style={{padding:"6px"}}>{fmtR(v.preco_venda)}</td>
                                  <td style={{padding:"6px",color:"#f59e0b",fontWeight:700}}>{fmtR(custoVar(v))}</td>
                                  <td style={{padding:"6px",color:"#059669",fontWeight:700}}>{fmtR(liquidoVar(p,v))}</td>
                                  <td style={{padding:"6px"}}><span style={{fontWeight:800,color:corMargem(marg),background:bgMargem(marg),borderRadius:6,padding:"1px 6px"}}>{fmtM(marg)}</span></td>
                                  <td style={{padding:"6px"}}><span style={{fontWeight:800,color:corMP(margemPreco(p,v))}}>{fmtM(margemPreco(p,v))}</span></td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:bgMargem(mm),borderRadius:8,padding:"8px 10px",marginBottom:8}}>
                        <span style={{fontSize:10,fontWeight:700,color:"#64748b",letterSpacing:0.5}}>MARGEM MÉDIA DO PRODUTO</span>
                        <span style={{fontSize:15,fontWeight:800,color:corMargem(mm)}}>{fmtM(mm)}</span>
                      </div>
                      <div style={{fontSize:11,color:"#64748b",textAlign:"right"}}>
                        <span style={{color:"#7c3aed",fontWeight:600}}>Ver detalhes →</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
            {modais}
          </div>
        );
      })()}

      {/* ── TAB: ESTOQUE ── */}
      {activeTab==="estoque"&&isAdminEquipe&&(()=>{
        const baixo = (p) => p.quantidade_atual <= p.quantidade_minima;
        const produtosFiltrados = estoqueProdutos.filter(p =>
          p.nome.toLowerCase().includes(estoqueSearch.toLowerCase())
        );
        const reloadEstoque = () => {
          fetchEstoqueProdutos(user.id, orgId).then(setEstoqueProdutos);
          fetchMovimentosEstoque(user.id, orgId, null, 100).then(setEstoqueMovimentos);
        };
        const abrirFormNovo = () => {
          setEditandoProduto(null);
          setEstoqueProdNome(""); setEstoqueProdUnidade("un"); setEstoqueProdQtd(""); setEstoqueProdMin("");
          setShowEstoqueForm(true);
        };
        const abrirFormEditar = (p) => {
          setEditandoProduto(p);
          setEstoqueProdNome(p.nome); setEstoqueProdUnidade(p.unidade);
          setEstoqueProdQtd(String(p.quantidade_atual)); setEstoqueProdMin(String(p.quantidade_minima));
          setShowEstoqueForm(true);
        };
        const salvarProduto = async () => {
          if (!estoqueProdNome.trim()) return;
          setSalvandoProduto(true);
          const antesEst = editandoProduto ? pick(editandoProduto,COLS_ESTOQUE_PROD) : null;
          const res = await salvarEstoqueProduto(user.id, orgId, {
            id: editandoProduto?.id,
            nome: estoqueProdNome.trim(),
            unidade: estoqueProdUnidade,
            quantidade_atual: editandoProduto ? Number(estoqueProdQtd)||0 : Number(estoqueProdQtd)||0,
            quantidade_minima: Number(estoqueProdMin)||0,
          });
          setSalvandoProduto(false);
          if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
          setShowEstoqueForm(false);
          logOp(editandoProduto
            ? {tipo:"estoque",acao:"editar",descricao:"Material editado: "+estoqueProdNome.trim(),tabela:"estoque_produtos",registro_id:editandoProduto.id,antes:antesEst,depois:{...antesEst,nome:estoqueProdNome.trim(),unidade:estoqueProdUnidade,quantidade_atual:Number(estoqueProdQtd)||0,quantidade_minima:Number(estoqueProdMin)||0}}
            : {tipo:"estoque",acao:"criar",descricao:"Material criado: "+estoqueProdNome.trim(),tabela:"estoque_produtos",registro_id:res.data?.[0]?.id});
          const produtos = await fetchEstoqueProdutos(user.id, orgId);
          setEstoqueProdutos(produtos);
          fetchMovimentosEstoque(user.id, orgId, null, 100).then(setEstoqueMovimentos);
          if (abrirPlusAposSalvar && res.data?.[0]) {
            const novo = res.data[0];
            setAbrirPlusAposSalvar(false);
            setEstoqueProdPlusId(novo.id);
            setEstoqueProdFornId(""); setEstoqueProdPreco(""); setEstoqueProdLink(""); setEstoqueProdObs(""); setEstoqueProdCat("");
            setShowEstoqueProdPlus(true);
          } else {
            setAbrirPlusAposSalvar(false);
          }
          showToast(editandoProduto?"✅ Produto atualizado!":"✅ Produto cadastrado!","#059669");
        };
        const excluirProduto = async (id) => {
          if (!window.confirm("Remover este produto do estoque?")) return;
          const prodAntes = estoqueProdutos.find(x=>x.id===id);
          await excluirEstoqueProduto(id);
          logOp({tipo:"estoque",acao:"excluir",descricao:"Material removido: "+(prodAntes?.nome||""),tabela:"estoque_produtos",registro_id:id,antes:pick(prodAntes,COLS_ESTOQUE_PROD)});
          reloadEstoque();
          showToast("🗑️ Produto removido.","#64748b");
        };
        const abrirMov = (produtoId, tipo) => {
          setShowMovForm(produtoId); setMovTipo(tipo); setMovQtd(""); setMovMotivo("");
        };
        const salvarMov = async () => {
          if (!movQtd || Number(movQtd)<=0) { showToast("❌ Informe a quantidade.","#991b1b"); return; }
          setSalvandoMov(true);
          const res = await registrarMovimentoEstoque(user.id, orgId, meuNome, {
            produto_id: showMovForm,
            tipo: movTipo,
            quantidade: Number(movQtd),
            motivo: movMotivo,
          });
          setSalvandoMov(false);
          if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
          setShowMovForm(null);
          if (res.movimento) {
            const pm = estoqueProdutos.find(x=>x.id===showMovForm);
            logOp({tipo:"estoque",acao:"criar",descricao:`${movTipo==="entrada"?"Entrada":"Saída"}: ${Number(movQtd)} ${pm?.unidade||""} de ${pm?.nome||"material"}`,tabela:"estoque_movimentos",registro_id:res.movimento.id,depois:{...pick(res.movimento,COLS_MOV),quantidade:Number(res.movimento.quantidade)}});
          }
          reloadEstoque();
          showToast(movTipo==="entrada"?"📥 Entrada registrada!":"📤 Saída registrada!",movTipo==="entrada"?"#059669":"#f59e0b");
        };
        return (
          <div>
            {/* Header */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
              <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#1e293b"}}>🏪 Controle de Estoque</h2>
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                <div style={{position:"relative"}}>
                  <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",fontSize:14,color:"#94a3b8"}}>🔍</span>
                  <input
                    value={estoqueSearch} onChange={e=>setEstoqueSearch(e.target.value)}
                    placeholder="Pesquisar material..."
                    style={{paddingLeft:32,paddingRight:12,paddingTop:8,paddingBottom:8,border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",width:190,background:"#fff"}}
                  />
                </div>
                <button onClick={()=>{
                  if (estoquePlusDesbloqueado) { setActiveTab("estoque_plus"); }
                  else if (!estoquePlusConfig) { setShowDefinirPin(true); }
                  else { setPinInput(""); setPinErro(""); setShowEstoquePlusPin(true); }
                }} style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",color:"#fff",border:"none",borderRadius:10,padding:"9px 14px",fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:5}}>
                  🔒 Estoque+
                </button>
              </div>
            </div>

            {/* Produtos */}
            {produtosFiltrados.length===0 ? (
              <div style={{background:"#fff",borderRadius:16,padding:40,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:40,marginBottom:8}}>📦</div>
                <div style={{fontSize:16,fontWeight:700,color:"#1e293b",marginBottom:6}}>{estoqueSearch ? "Nenhum resultado para \""+estoqueSearch+"\"" : "Nenhum material cadastrado"}</div>
                <div style={{fontSize:13,color:"#64748b",marginBottom:20}}>{estoqueSearch ? "Tente outro termo de pesquisa." : "Cadastre papel fotográfico, caixas, envelopes e outros insumos."}</div>
                {!estoqueSearch && <button onClick={abrirFormNovo} style={{background:"#0d9488",color:"#fff",border:"none",borderRadius:10,padding:"10px 22px",fontSize:13,fontWeight:700,cursor:"pointer"}}>+ Cadastrar Primeiro Material</button>}
              </div>
            ) : (
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:14,marginBottom:24}}>
                {produtosFiltrados.map(p=>(
                  <div key={p.id} style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",border:`2px solid ${baixo(p)?"#fde68a":"#f1f5f9"}`,position:"relative"}}>
                    {baixo(p)&&<span style={{position:"absolute",top:10,right:12,background:"#f59e0b",color:"#fff",borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700}}>⚠️ Baixo</span>}
                    <div style={{fontSize:15,fontWeight:700,color:"#1e293b",marginBottom:4,paddingRight:60}}>{p.nome}</div>
                    <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:10}}>
                      <span style={{fontSize:32,fontWeight:800,color:baixo(p)?"#f59e0b":"#0d9488",lineHeight:1}}>{p.quantidade_atual}</span>
                      <span style={{fontSize:13,color:"#64748b"}}>{p.unidade}</span>
                      <span style={{fontSize:11,color:"#94a3b8",marginLeft:4}}>mín: {p.quantidade_minima}</span>
                    </div>
                    <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                      <button onClick={()=>abrirMov(p.id,"entrada")} style={{flex:1,padding:"6px 10px",background:"#d1fae5",color:"#065f46",border:"none",borderRadius:8,fontSize:11,fontWeight:700,cursor:"pointer"}}>📥 Entrada</button>
                      <button onClick={()=>abrirMov(p.id,"saida")} style={{flex:1,padding:"6px 10px",background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:8,fontSize:11,fontWeight:700,cursor:"pointer"}}>📤 Saída</button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* ── Pedidos WhatsApp agrupados por fornecedor ───────────────── */}
            {(()=>{
              // Agrupa produtos com estoque baixo que têm fornecedor com telefone
              const baixoComForn = estoqueProdutos.filter(p => baixo(p) && p.fornecedores?.contato);
              if (!baixoComForn.length) return null;

              // Agrupa por fornecedor_id
              const porForn = {};
              baixoComForn.forEach(p => {
                const fId = p.fornecedores.id;
                if (!porForn[fId]) porForn[fId] = { forn: p.fornecedores, itens: [] };
                porForn[fId].itens.push(p);
              });

              const whatsIcon = (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413Z"/>
                </svg>
              );

              return (
                <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20,border:"2px solid #dcfce7"}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#15803d",marginBottom:12,display:"flex",alignItems:"center",gap:6}}>
                    {whatsIcon} Pedidos pendentes — enviar para fornecedores
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:10}}>
                    {Object.values(porForn).map(({forn, itens}) => {
                      const telefone = forn.contato.replace(/\D/g,"");
                      const tel = telefone.startsWith("55") ? telefone : "55"+telefone;

                      // Monta lista de itens para a mensagem
                      const linhasItens = itens.map(p => {
                        const faltam = Math.max(0, (p.quantidade_minima||0) - (p.quantidade_atual||0));
                        const un = p.unidade || "un";
                        return `• ${faltam} ${un} de *${p.nome}*`;
                      }).join("\n");

                      const msg = encodeURIComponent(
                        `Oi, tudo bem?\n\nPreciso repor o seguinte:\n${linhasItens}\n\nPode me passar disponibilidade, preço e prazo para entrega? Obrigado!`
                      );
                      const url = `https://wa.me/${tel}?text=${msg}`;

                      return (
                        <div key={forn.id} style={{background:"#f0fdf4",borderRadius:10,padding:"12px 14px",border:"1px solid #bbf7d0"}}>
                          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:10,flexWrap:"wrap"}}>
                            <div style={{flex:1}}>
                              <div style={{fontSize:12,fontWeight:700,color:"#166534",marginBottom:4}}>{forn.nome}</div>
                              <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                                {itens.map(p => {
                                  const faltam = Math.max(0, (p.quantidade_minima||0) - (p.quantidade_atual||0));
                                  return (
                                    <span key={p.id} style={{background:"#fef3c7",color:"#92400e",borderRadius:20,padding:"2px 9px",fontSize:11,fontWeight:600}}>
                                      {p.nome}: +{faltam} {p.unidade||"un"}
                                    </span>
                                  );
                                })}
                              </div>
                            </div>
                            <a href={url} target="_blank" rel="noopener noreferrer"
                              style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",background:"#16a34a",color:"#fff",borderRadius:9,fontSize:12,fontWeight:700,textDecoration:"none",flexShrink:0}}>
                              {whatsIcon} Enviar pedido
                            </a>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}

            {/* Histórico de movimentos */}
            {estoqueMovimentos.length>0&&(
              <div style={{background:"#fff",borderRadius:16,padding:"18px 20px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:14,fontWeight:700,color:"#1e293b",marginBottom:12}}>📊 Últimos Movimentos</div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr>
                        {["Data","Produto","Tipo","Qtd","Motivo","Quem"].map(h=>(
                          <th key={h} style={{...TH}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {estoqueMovimentos.map(m=>(
                        <tr key={m.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                          <td style={TD}>{new Date(m.criado_em).toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}</td>
                          <td style={TD}><strong>{m.estoque_produtos?.nome||"—"}</strong></td>
                          <td style={TD}>
                            <span style={{background:m.tipo==="entrada"?"#d1fae5":"#fde8e8",color:m.tipo==="entrada"?"#065f46":"#991b1b",borderRadius:20,padding:"2px 10px",fontWeight:700,fontSize:11}}>
                              {m.tipo==="entrada"?"📥 Entrada":"📤 Saída"}
                            </span>
                          </td>
                          <td style={TD}><strong>{m.quantidade}</strong> {m.estoque_produtos?.unidade||""}</td>
                          <td style={TD}>{m.motivo||"—"}</td>
                          <td style={TD}>{m.registrado_por||"—"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Modal: form cadastro/edição */}
            {showEstoqueForm&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
                <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                  <h3 style={{margin:"0 0 18px",fontSize:16,fontWeight:700,color:"#1e293b"}}>{editandoProduto?"✏️ Editar Material":"+ Novo Material"}</h3>
                  <div style={{display:"flex",flexDirection:"column",gap:12}}>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Nome do material *</label>
                      {editandoProduto ? (
                        <div style={{width:"100%",padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,background:"#f8fafc",color:"#64748b",boxSizing:"border-box"}}>{estoqueProdNome}</div>
                      ) : (
                        <input value={estoqueProdNome} onChange={e=>setEstoqueProdNome(e.target.value)} placeholder="Ex: Papel Fotográfico 10x15"
                          style={{width:"100%",padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      )}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Unidade</label>
                        <select value={estoqueProdUnidade} onChange={e=>setEstoqueProdUnidade(e.target.value)}
                          style={{width:"100%",padding:"10px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",background:"#fff"}}>
                          {["un","cx","pct","rolo","m","kg","l"].map(u=><option key={u} value={u}>{u}</option>)}
                        </select>
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Qtd atual</label>
                        <input type="number" min="0" value={estoqueProdQtd} onChange={e=>setEstoqueProdQtd(e.target.value)}
                          style={{width:"100%",padding:"10px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Qtd mínima</label>
                        <input type="number" min="0" value={estoqueProdMin} onChange={e=>setEstoqueProdMin(e.target.value)}
                          style={{width:"100%",padding:"10px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:10,marginTop:20}}>
                    <button onClick={()=>setShowEstoqueForm(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                    <button onClick={salvarProduto} disabled={!estoqueProdNome.trim()||salvandoProduto}
                      style={{flex:1,padding:"10px 0",border:"none",borderRadius:10,background:estoqueProdNome.trim()?"#0d9488":"#e2e8f0",color:estoqueProdNome.trim()?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                      {salvandoProduto?"Salvando...":"Salvar"}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Modal: registrar movimento */}
            {showMovForm&&(()=>{
              const prod = estoqueProdutos.find(p=>p.id===showMovForm);
              return (
                <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
                  <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:400,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                    <h3 style={{margin:"0 0 6px",fontSize:16,fontWeight:700,color:"#1e293b"}}>{movTipo==="entrada"?"📥 Registrar Entrada":"📤 Registrar Saída"}</h3>
                    <p style={{margin:"0 0 16px",fontSize:12,color:"#64748b"}}>{prod?.nome} · saldo atual: <strong>{prod?.quantidade_atual} {prod?.unidade}</strong></p>
                    <div style={{display:"flex",gap:8,marginBottom:14}}>
                      {["entrada","saida"].map(t=>(
                        <button key={t} onClick={()=>setMovTipo(t)}
                          style={{flex:1,padding:"9px",border:`2px solid ${movTipo===t?(t==="entrada"?"#059669":"#ef4444"):"#e2e8f0"}`,borderRadius:10,background:movTipo===t?(t==="entrada"?"#d1fae5":"#fde8e8"):"#fff",fontWeight:700,fontSize:12,cursor:"pointer",color:movTipo===t?(t==="entrada"?"#065f46":"#991b1b"):"#64748b"}}>
                          {t==="entrada"?"📥 Entrada":"📤 Saída"}
                        </button>
                      ))}
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:10}}>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Quantidade *</label>
                        <input type="number" min="1" value={movQtd} onChange={e=>setMovQtd(e.target.value)} autoFocus
                          style={{width:"100%",padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:14,outline:"none",boxSizing:"border-box",fontWeight:600,textAlign:"center"}} />
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Motivo (opcional)</label>
                        <input value={movMotivo} onChange={e=>setMovMotivo(e.target.value)} placeholder="Ex: compra, uso do dia, ajuste..."
                          style={{width:"100%",padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                    </div>
                    <div style={{display:"flex",gap:10,marginTop:18}}>
                      <button onClick={()=>setShowMovForm(null)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                      <button onClick={salvarMov} disabled={!movQtd||Number(movQtd)<=0||salvandoMov}
                        style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:(movQtd&&Number(movQtd)>0)?(movTipo==="entrada"?"#059669":"#ef4444"):"#e2e8f0",color:(movQtd&&Number(movQtd)>0)?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                        {salvandoMov?"Salvando...":movTipo==="entrada"?"✅ Confirmar Entrada":"✅ Confirmar Saída"}
                      </button>
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        );
      })()}

      {/* ── ABA ESTOQUE+ ──────────────────────────────────────────────────── */}
      {activeTab==="estoque_plus"&&isAdminEquipe&&estoquePlusDesbloqueado&&(()=>{
        // helpers
        const hashPin = async (pin) => {
          const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pin));
          return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join("");
        };

        const reloadForn = () => fetchFornecedores(user.id, orgId).then(setFornecedores);
        const reloadMov  = () => fetchMovimentosComPreco(user.id, orgId, mesSelecionado).then(setMovimentosMes);

        const abrirFornForm = (f=null) => {
          setEditandoFornecedor(f);
          setFornNome(f?.nome||""); setFornContato(f?.contato||"");
          setFornSite(f?.site_link||""); setFornObs(f?.observacoes||"");
          setShowFornecedorForm(true);
        };
        const salvarForn = async () => {
          if (!fornNome.trim()) return;
          setSalvandoForn(true);
          const res = await salvarFornecedor(user.id, orgId, {
            id: editandoFornecedor?.id,
            nome: fornNome.trim(), contato: fornContato, site_link: fornSite, observacoes: fornObs,
          });
          setSalvandoForn(false);
          if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
          setShowFornecedorForm(false);
          logOp({tipo:"fornecedores",acao:editandoFornecedor?"editar":"criar",descricao:(editandoFornecedor?"Fornecedor editado: ":"Fornecedor criado: ")+fornNome.trim(),tabela:"fornecedores",registro_id:editandoFornecedor?.id||res.data?.id,antes:editandoFornecedor?pick(editandoFornecedor,COLS_FORNECEDOR):null});
          reloadForn();
          showToast(editandoFornecedor?"✅ Fornecedor atualizado!":"✅ Fornecedor cadastrado!","#059669");
        };
        const excluirForn = async (id) => {
          if (!window.confirm("Remover este fornecedor?")) return;
          const fAntes = fornecedores.find(x=>x.id===id);
          await excluirFornecedor(id);
          logOp({tipo:"fornecedores",acao:"excluir",descricao:"Fornecedor removido: "+(fAntes?.nome||""),tabela:"fornecedores",registro_id:id,antes:pick(fAntes,COLS_FORNECEDOR)});
          reloadForn();
          showToast("🗑️ Fornecedor removido.","#64748b");
        };

        const abrirProdPlus = (p) => {
          setEstoqueProdPlusId(p.id);
          setEstoqueProdFornId(p.fornecedor_id||"");
          setEstoqueProdPreco(p.preco_unitario?String(p.preco_unitario):"");
          setEstoqueProdLink(p.link_produto||"");
          setEstoqueProdObs(p.observacoes||"");
          setEstoqueProdCat(p.categoria||"");
          setShowEstoqueProdPlus(true);
        };
        const salvarProdPlus = async () => {
          if (!estoqueProdPlusId) return;
          setSalvandoProdPlus(true);
          const prodPlusAntes = estoqueProdutos.find(x=>x.id===estoqueProdPlusId);
          const camposPlus = {
            preco_unitario: estoqueProdPreco ? Number(estoqueProdPreco) : null,
            fornecedor_id:  estoqueProdFornId || null,
            link_produto:   estoqueProdLink || null,
            observacoes:    estoqueProdObs || null,
            categoria:      estoqueProdCat || null,
          };
          const res = await atualizarEstoqueProdutoPlus(estoqueProdPlusId, {
            preco_unitario: estoqueProdPreco ? Number(estoqueProdPreco) : null,
            fornecedor_id:  estoqueProdFornId || null,
            link_produto:   estoqueProdLink || null,
            observacoes:    estoqueProdObs || null,
            categoria:      estoqueProdCat || null,
          });
          setSalvandoProdPlus(false);
          if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
          setShowEstoqueProdPlus(false);
          logOp({tipo:"estoque",acao:"editar",descricao:"Preço/fornecedor atualizado: "+(prodPlusAntes?.nome||"material"),tabela:"estoque_produtos",registro_id:estoqueProdPlusId,antes:pick(prodPlusAntes,["preco_unitario","fornecedor_id","link_produto","observacoes","categoria"]),depois:camposPlus});
          fetchEstoqueProdutos(user.id, orgId).then(setEstoqueProdutos);
          showToast("✅ Informações salvas!","#059669");
        };

        // totais do mês
        const totalMes = movimentosMes.reduce((s,m)=>{
          const preco = m.preco_unitario||0;
          return s + (preco * m.quantidade);
        }, 0);
        const porForn = {};
        movimentosMes.forEach(m=>{
          const nome = m.fornecedores?.nome || "Sem fornecedor";
          const val  = (m.preco_unitario||0)*m.quantidade;
          porForn[nome] = (porForn[nome]||0) + val;
        });

        return (
          <div>
            {/* Header */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:10}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <button onClick={()=>setActiveTab("estoque")} style={{background:"#f1f5f9",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,cursor:"pointer",color:"#64748b"}}>← Voltar</button>
                <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#7c3aed"}}>🔒 Estoque+</h2>
                <span style={{background:"#ede9fe",color:"#6d28d9",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>Área restrita</span>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <button onClick={()=>{setAbrirPlusAposSalvar(true);setEditandoProduto(null);setEstoqueProdNome("");setEstoqueProdUnidade("un");setEstoqueProdQtd("0");setEstoqueProdMin("0");setShowEstoqueForm(true);}} style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",color:"#fff",border:"none",borderRadius:10,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:"pointer"}}>+ Novo Material</button>
                <button onClick={abrirNota} style={{background:"linear-gradient(135deg,#0d9488,#0f766e)",color:"#fff",border:"none",borderRadius:10,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📷 Ler nota fiscal</button>
                <button onClick={()=>setShowDefinirPin(true)} style={{background:"#f1f5f9",color:"#64748b",border:"1px solid #e2e8f0",borderRadius:10,padding:"8px 14px",fontSize:12,cursor:"pointer"}}>⚙️ Alterar PIN</button>
                <button onClick={()=>{setEstoquePlusDesbloqueado(false);setActiveTab("estoque");}} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:10,padding:"8px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🔒 Trancar</button>
              </div>
            </div>

            {/* Cards resumo do mês */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:12,marginBottom:20}}>
              <div style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",borderRadius:14,padding:"16px 18px",color:"#fff"}}>
                <div style={{fontSize:11,opacity:.8,marginBottom:4}}>💰 Investido no mês</div>
                <div style={{fontSize:26,fontWeight:800}}>R$ {totalMes.toFixed(2).replace(".",",")}</div>
                <div style={{fontSize:11,opacity:.7}}>{mesSelecionado}</div>
              </div>
              <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:11,color:"#64748b",marginBottom:4}}>📦 Materiais cadastrados</div>
                <div style={{fontSize:26,fontWeight:800,color:"#1e293b"}}>{estoqueProdutos.length}</div>
                <div style={{fontSize:11,color:"#94a3b8"}}>itens no estoque</div>
              </div>
              <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
                <div style={{fontSize:11,color:"#64748b",marginBottom:4}}>🏭 Fornecedores</div>
                <div style={{fontSize:26,fontWeight:800,color:"#1e293b"}}>{fornecedores.length}</div>
                <div style={{fontSize:11,color:"#94a3b8"}}>cadastrados</div>
              </div>
            </div>

            {/* Seletor de mês + totais por fornecedor */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>📊 Compras por fornecedor</div>
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <input type="month" value={mesSelecionado} onChange={e=>{setMesSelecionado(e.target.value); fetchMovimentosComPreco(user.id,orgId,e.target.value).then(setMovimentosMes);}}
                    style={{padding:"6px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none"}} />
                  <button onClick={reloadMov} style={{background:"#ede9fe",color:"#7c3aed",border:"none",borderRadius:8,padding:"6px 12px",fontSize:12,cursor:"pointer"}}>🔄 Atualizar</button>
                </div>
              </div>
              {Object.keys(porForn).length===0 ? (
                <div style={{textAlign:"center",padding:"20px 0",color:"#94a3b8",fontSize:12}}>Nenhuma entrada com preço registrada em {mesSelecionado}.</div>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {Object.entries(porForn).sort((a,b)=>b[1]-a[1]).map(([nome,val])=>(
                    <div key={nome} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 10px",background:"#f8fafc",borderRadius:8}}>
                      <span style={{fontSize:13,color:"#374151"}}>{nome}</span>
                      <span style={{fontSize:13,fontWeight:700,color:"#7c3aed"}}>R$ {val.toFixed(2).replace(".",",")}</span>
                    </div>
                  ))}
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 10px",background:"#ede9fe",borderRadius:8,marginTop:4}}>
                    <span style={{fontSize:13,fontWeight:700,color:"#6d28d9"}}>Total do mês</span>
                    <span style={{fontSize:14,fontWeight:800,color:"#6d28d9"}}>R$ {totalMes.toFixed(2).replace(".",",")}</span>
                  </div>
                </div>
              )}
            </div>

            {/* Materiais com info de preço/fornecedor */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:12}}>🏷️ Preço e fornecedor por material</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:10}}>
                {estoqueProdutos.map(p=>{
                  const forn = fornecedores.find(f=>f.id===p.fornecedor_id);
                  return (
                    <div key={p.id} style={{border:"1.5px solid #e2e8f0",borderRadius:10,padding:"12px 14px"}}>
                      <div style={{fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:6}}>{p.nome}</div>
                      {p.categoria&&<div style={{display:"inline-block",background:"#f1f5f9",borderRadius:20,padding:"1px 8px",fontSize:10,color:"#64748b",marginBottom:6}}>{p.categoria}</div>}
                      <div style={{fontSize:12,color:"#374151",marginBottom:2}}>
                        💰 Preço: {p.preco_unitario ? <strong style={{color:"#059669"}}>R$ {Number(p.preco_unitario).toFixed(2).replace(".",",")} / {p.unidade}</strong> : <span style={{color:"#94a3b8"}}>não informado</span>}
                      </div>
                      <div style={{fontSize:12,color:"#374151",marginBottom:2}}>
                        🏭 Fornecedor: {forn ? <strong>{forn.nome}</strong> : <span style={{color:"#94a3b8"}}>não informado</span>}
                      </div>
                      {p.link_produto&&<div style={{fontSize:11,marginBottom:2}}><a href={p.link_produto} target="_blank" rel="noreferrer" style={{color:"#6d28d9"}}>🔗 Ver produto</a></div>}
                      {p.observacoes&&<div style={{fontSize:11,color:"#64748b",fontStyle:"italic",marginBottom:4}}>"{p.observacoes}"</div>}
                      <button onClick={()=>abrirProdPlus(p)} style={{marginTop:6,width:"100%",padding:"5px 0",background:"#ede9fe",color:"#7c3aed",border:"none",borderRadius:7,fontSize:11,fontWeight:700,cursor:"pointer"}}>✏️ Editar info</button>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Fornecedores */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>🏭 Fornecedores</div>
                <button onClick={()=>abrirFornForm()} style={{background:"#7c3aed",color:"#fff",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>+ Novo Fornecedor</button>
              </div>
              {fornecedores.length===0 ? (
                <div style={{textAlign:"center",padding:"20px 0",color:"#94a3b8",fontSize:12}}>Nenhum fornecedor cadastrado ainda.</div>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:8}}>
                  {fornecedores.map(f=>(
                    <div key={f.id} style={{border:"1.5px solid #e2e8f0",borderRadius:10,padding:"11px 14px",display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                      <div>
                        <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>{f.nome}</div>
                        {f.contato&&<div style={{fontSize:11,color:"#64748b"}}>📞 {f.contato}</div>}
                        {f.site_link&&<div style={{fontSize:11}}><a href={f.site_link} target="_blank" rel="noreferrer" style={{color:"#6d28d9"}}>🔗 {f.site_link.replace(/^https?:\/\//,"")}</a></div>}
                        {f.observacoes&&<div style={{fontSize:11,color:"#94a3b8",fontStyle:"italic"}}>"{f.observacoes}"</div>}
                      </div>
                      <div style={{display:"flex",gap:6,flexShrink:0}}>
                        <button onClick={()=>abrirFornForm(f)} style={{padding:"5px 9px",background:"#f1f5f9",border:"none",borderRadius:7,fontSize:11,cursor:"pointer"}}>✏️</button>
                        <button onClick={()=>excluirForn(f.id)} style={{padding:"5px 9px",background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:7,fontSize:11,cursor:"pointer"}}>🗑️</button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Modal: fornecedor form */}
            {showFornecedorForm&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                  <h3 style={{margin:"0 0 16px",fontSize:15,fontWeight:700}}>{editandoFornecedor?"✏️ Editar Fornecedor":"+ Novo Fornecedor"}</h3>
                  <div style={{display:"flex",flexDirection:"column",gap:11}}>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Nome *</label>
                      <input value={fornNome} onChange={e=>setFornNome(e.target.value)} placeholder="Ex: Shopee Shop Papel, ML Insumos..."
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Contato (WhatsApp / email)</label>
                      <input value={fornContato} onChange={e=>setFornContato(e.target.value)} placeholder="Ex: 21 99999-0000"
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Link da loja / site</label>
                      <input value={fornSite} onChange={e=>setFornSite(e.target.value)} placeholder="https://..."
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Observações</label>
                      <input value={fornObs} onChange={e=>setFornObs(e.target.value)} placeholder="Ex: melhor preço em junho, frete grátis acima de R$100..."
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                  </div>
                  <div style={{display:"flex",gap:10,marginTop:18}}>
                    <button onClick={()=>setShowFornecedorForm(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                    <button onClick={salvarForn} disabled={!fornNome.trim()||salvandoForn}
                      style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:fornNome.trim()?"#7c3aed":"#e2e8f0",color:fornNome.trim()?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                      {salvandoForn?"Salvando...":"Salvar fornecedor"}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Modal: novo material (Estoque+) */}
            {showEstoqueForm&&abrirPlusAposSalvar&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                  <h3 style={{margin:"0 0 14px",fontSize:16,fontWeight:700,color:"#1e293b"}}>+ Novo Material</h3>
                  <div style={{display:"flex",flexDirection:"column",gap:12}}>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Nome do material *</label>
                      <input value={estoqueProdNome} onChange={e=>setEstoqueProdNome(e.target.value)} autoFocus placeholder="Ex: Papel Fotográfico 10x15"
                        style={{width:"100%",padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Unidade</label>
                        <select value={estoqueProdUnidade} onChange={e=>setEstoqueProdUnidade(e.target.value)}
                          style={{width:"100%",padding:"10px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",background:"#fff"}}>
                          {["un","cx","pct","rolo","m","kg","l"].map(u=><option key={u} value={u}>{u}</option>)}
                        </select>
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Qtd atual</label>
                        <input type="number" min="0" value={estoqueProdQtd} onChange={e=>setEstoqueProdQtd(e.target.value)}
                          style={{width:"100%",padding:"10px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"block"}}>Qtd mínima</label>
                        <input type="number" min="0" value={estoqueProdMin} onChange={e=>setEstoqueProdMin(e.target.value)}
                          style={{width:"100%",padding:"10px 8px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:10,marginTop:20}}>
                    <button onClick={()=>{setShowEstoqueForm(false);setAbrirPlusAposSalvar(false);}} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                    <button disabled={!estoqueProdNome.trim()||salvandoProduto} onClick={async()=>{
                      if (!estoqueProdNome.trim()) return;
                      setSalvandoProduto(true);
                      const res = await salvarEstoqueProduto(user.id, orgId, {
                        nome: estoqueProdNome.trim(),
                        unidade: estoqueProdUnidade,
                        quantidade_atual: Number(estoqueProdQtd)||0,
                        quantidade_minima: Number(estoqueProdMin)||0,
                      });
                      setSalvandoProduto(false);
                      if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
                      const produtos = await fetchEstoqueProdutos(user.id, orgId);
                      setEstoqueProdutos(produtos);
                      setShowEstoqueForm(false);
                      setAbrirPlusAposSalvar(false);
                      const novo = res.data?.[0];
                      if (novo) logOp({tipo:"estoque",acao:"criar",descricao:"Material criado: "+estoqueProdNome.trim(),tabela:"estoque_produtos",registro_id:novo.id});
                      if (novo) {
                        setEstoqueProdPlusId(novo.id);
                        setEstoqueProdFornId(""); setEstoqueProdPreco(""); setEstoqueProdLink(""); setEstoqueProdObs(""); setEstoqueProdCat("");
                        setShowEstoqueProdPlus(true);
                      }
                      showToast("✅ Material cadastrado!","#059669");
                    }}
                      style={{flex:1,padding:"10px 0",border:"none",borderRadius:10,background:estoqueProdNome.trim()?"#7c3aed":"#e2e8f0",color:estoqueProdNome.trim()?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                      {salvandoProduto?"Salvando...":"Salvar"}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Modal: editar info plus de um produto */}
            {showEstoqueProdPlus&&(()=>{
              const prod = estoqueProdutos.find(p=>p.id===estoqueProdPlusId);
              const filtForn = fornecedores.filter(f=>f.nome.toLowerCase().includes(fornNome.toLowerCase()));
              return (
                <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                  <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:440,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                    <h3 style={{margin:"0 0 4px",fontSize:15,fontWeight:700}}>✏️ Info financeiras</h3>
                    <p style={{margin:"0 0 14px",fontSize:12,color:"#64748b"}}>{prod?.nome}</p>
                    <div style={{display:"flex",flexDirection:"column",gap:11}}>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Preço de compra por {prod?.unidade||"un"} (R$)</label>
                        <input type="number" min="0" step="0.01" value={estoqueProdPreco} onChange={e=>setEstoqueProdPreco(e.target.value)} placeholder="Ex: 0,80"
                          style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                      <div>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:3}}>
                          <label style={{fontSize:11,fontWeight:600,color:"#64748b"}}>Fornecedor</label>
                          <button type="button" onClick={()=>{setShowEstoqueProdPlus(false);abrirFornForm();}} style={{background:"none",border:"none",color:"#7c3aed",fontSize:11,fontWeight:700,cursor:"pointer",padding:0}}>+ Cadastrar fornecedor</button>
                        </div>
                        {fornecedores.length===0 ? (
                          <div style={{padding:"9px 12px",border:"1.5px dashed #e2e8f0",borderRadius:10,fontSize:12,color:"#94a3b8",textAlign:"center"}}>
                            Nenhum fornecedor cadastrado. Clique em <strong style={{color:"#7c3aed"}}>+ Cadastrar fornecedor</strong> acima.
                          </div>
                        ) : (
                          <select value={estoqueProdFornId} onChange={e=>setEstoqueProdFornId(e.target.value)}
                            style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                            <option value="">— Selecionar fornecedor —</option>
                            {fornecedores.map(f=><option key={f.id} value={f.id}>{f.nome}</option>)}
                          </select>
                        )}
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Categoria</label>
                        <input value={estoqueProdCat} onChange={e=>setEstoqueProdCat(e.target.value)} placeholder="Ex: papel, embalagem, tinta, imã..."
                          style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Link do produto</label>
                        <input value={estoqueProdLink} onChange={e=>setEstoqueProdLink(e.target.value)} placeholder="https://shopee.com.br/..."
                          style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Observações</label>
                        <input value={estoqueProdObs} onChange={e=>setEstoqueProdObs(e.target.value)} placeholder="Ex: comprar só quando em promoção..."
                          style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                    </div>
                    <div style={{display:"flex",gap:10,marginTop:18}}>
                      <button onClick={()=>setShowEstoqueProdPlus(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                      <button onClick={salvarProdPlus} disabled={salvandoProdPlus}
                        style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:"#7c3aed",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                        {salvandoProdPlus?"Salvando...":"Salvar"}
                      </button>
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        );
      })()}

      {/* ── Modais PIN Estoque+ ───────────────────────────────────────────── */}
      {showEstoquePlusPin&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10100,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"#fff",borderRadius:20,padding:32,width:"100%",maxWidth:340,boxShadow:"0 20px 60px rgba(0,0,0,0.25)",textAlign:"center"}}>
            <div style={{fontSize:44,marginBottom:10}}>🔒</div>
            <h3 style={{margin:"0 0 6px",fontSize:17,fontWeight:800,color:"#1e293b"}}>Estoque+</h3>
            <p style={{margin:"0 0 20px",fontSize:12,color:"#64748b"}}>Digite o PIN para acessar a área restrita</p>
            <input type="password" inputMode="numeric" maxLength={8} value={pinInput}
              onChange={e=>{setPinInput(e.target.value);setPinErro("");}}
              onKeyDown={async e=>{
                if (e.key==="Enter") {
                  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pinInput))
                    .then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
                  if (hash===estoquePlusConfig?.pin_hash) {
                    setEstoquePlusDesbloqueado(true);
                    setShowEstoquePlusPin(false);
                    setActiveTab("estoque_plus");
                    fetchMovimentosComPreco(user.id,orgId,mesSelecionado).then(setMovimentosMes);
                  } else { setPinErro("PIN incorreto. Tente novamente."); setPinInput(""); }
                }
              }}
              autoFocus
              placeholder="••••"
              style={{width:"100%",padding:"12px 16px",border:`2px solid ${pinErro?"#ef4444":"#e2e8f0"}`,borderRadius:12,fontSize:22,outline:"none",textAlign:"center",letterSpacing:8,boxSizing:"border-box",marginBottom:8}} />
            {pinErro&&<div style={{fontSize:12,color:"#ef4444",marginBottom:8}}>{pinErro}</div>}
            <button onClick={async()=>{
              const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pinInput))
                .then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
              if (hash===estoquePlusConfig?.pin_hash) {
                setEstoquePlusDesbloqueado(true);
                setShowEstoquePlusPin(false);
                setActiveTab("estoque_plus");
                fetchMovimentosComPreco(user.id,orgId,mesSelecionado).then(setMovimentosMes);
              } else { setPinErro("PIN incorreto. Tente novamente."); setPinInput(""); }
            }} style={{width:"100%",padding:"12px 0",background:"#7c3aed",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:800,cursor:"pointer",marginBottom:10}}>
              Desbloquear
            </button>
            <button onClick={()=>setShowEstoquePlusPin(false)} style={{background:"none",border:"none",color:"#94a3b8",fontSize:12,cursor:"pointer"}}>Cancelar</button>
          </div>
        </div>
      )}

      {/* Modal: definir / alterar PIN */}
      {showDefinirPin&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:10100,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"#fff",borderRadius:20,padding:28,width:"100%",maxWidth:360,boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
            <h3 style={{margin:"0 0 6px",fontSize:16,fontWeight:800,color:"#1e293b"}}>🔐 {estoquePlusConfig?"Alterar PIN":"Definir PIN do Estoque+"}</h3>
            <p style={{margin:"0 0 16px",fontSize:12,color:"#64748b"}}>Crie um PIN numérico. Só você terá acesso ao Estoque+.</p>
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <div>
                <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Novo PIN (4 a 8 dígitos)</label>
                <input type="password" inputMode="numeric" maxLength={8} value={novoPin} onChange={e=>setNovoPin(e.target.value)}
                  placeholder="••••" style={{width:"100%",padding:"10px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:18,outline:"none",textAlign:"center",letterSpacing:6,boxSizing:"border-box"}} />
              </div>
              <div>
                <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Confirmar PIN</label>
                <input type="password" inputMode="numeric" maxLength={8} value={novoPinConfirm} onChange={e=>setNovoPinConfirm(e.target.value)}
                  placeholder="••••" style={{width:"100%",padding:"10px 12px",border:`1.5px solid ${novoPinConfirm&&novoPinConfirm!==novoPin?"#ef4444":"#e2e8f0"}`,borderRadius:10,fontSize:18,outline:"none",textAlign:"center",letterSpacing:6,boxSizing:"border-box"}} />
                {novoPinConfirm&&novoPinConfirm!==novoPin&&<span style={{fontSize:11,color:"#ef4444"}}>PINs não coincidem</span>}
              </div>
            </div>
            <div style={{display:"flex",gap:10,marginTop:18}}>
              <button onClick={()=>{setShowDefinirPin(false);setNovoPin("");setNovoPinConfirm("");}} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
              <button disabled={novoPin.length<4||novoPin!==novoPinConfirm||salvandoPin}
                onClick={async()=>{
                  setSalvandoPin(true);
                  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(novoPin))
                    .then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
                  const res = await salvarEstoquePlusPin(user.id, orgId, hash);
                  setSalvandoPin(false);
                  if (res.error) { showToast("❌ "+res.error,"#991b1b"); return; }
                  setEstoquePlusConfig({pin_hash: hash});
                  setShowDefinirPin(false); setNovoPin(""); setNovoPinConfirm("");
                  setEstoquePlusDesbloqueado(true);
                  setActiveTab("estoque_plus");
                  fetchMovimentosComPreco(user.id,orgId,mesSelecionado).then(setMovimentosMes);
                  showToast("✅ PIN definido! Estoque+ desbloqueado.","#059669");
                }}
                style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:(novoPin.length>=4&&novoPin===novoPinConfirm)?"#7c3aed":"#e2e8f0",color:(novoPin.length>=4&&novoPin===novoPinConfirm)?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                {salvandoPin?"Salvando...":"Salvar PIN"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── ABA FINANCEIRO ──────────────────────────────────────────────── */}
      {activeTab==="upseller"&&isAdminEquipe&&(()=>{
        const inp = {padding:"8px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:13,outline:"none",boxSizing:"border-box",background:"#fff"};
        const STATUS = [["para_imprimir","Para imprimir"],["para_retirada","Para retirada"],["para_emitir","Para emitir"],["enviado","Enviado"],["cancelado","Cancelado"],["sem_status","Sem status"]];
        const corSt = {para_imprimir:["#1d4ed8","#dbeafe"],para_retirada:["#7c3aed","#ede9fe"],para_emitir:["#92400e","#fef3c7"],enviado:["#065f46","#d1fae5"],cancelado:["#991b1b","#fde8e8"]};
        const totalPag = Math.max(1, Math.ceil(ups.total/UPS_POR_PAG));
        const escolhida = opcoesVarTop.find(o=>o.id===upsEscolha);
        const sugestao = upsVincular&&!upsVincular.variacao_id&&!upsVincular.ignorar ? melhorMatch(upsVincular.variacao, opcoesVarTop, 0.45) : null;
        return (
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:10}}>
              <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#1d4ed8"}}>🧾 Pedidos do UpSeller</h2>
              <div style={{display:"flex",gap:8}}>
                <button onClick={abrirVinc} style={{background:"#eff6ff",color:"#1d4ed8",border:"1px solid #bfdbfe",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>⚙️ Baixa automática</button>
                <button onClick={carregarUps} style={{background:"#dbeafe",color:"#1e40af",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,cursor:"pointer"}}>🔄</button>
              </div>
            </div>
            <p style={{margin:"0 0 12px",fontSize:12,color:"#64748b"}}>Ao clicar em <strong>Vincular</strong> num pedido, você liga a variação dele a um produto cadastrado. Todos os pedidos (atuais e futuros) com essa mesma variação ficam vinculados automaticamente e passam a dar baixa no estoque.</p>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
              <input value={upsBusca} onChange={e=>setUpsBusca(e.target.value)} placeholder="Buscar por pedido, cliente ou variação…" style={{...inp,flex:1,minWidth:200}} />
              <select value={upsStatus} onChange={e=>{setUpsStatus(e.target.value);setUpsPag(0);}} style={inp}>
                <option value="">Todos os status</option>
                {STATUS.map(([v,l])=><option key={v} value={v}>{l}</option>)}
              </select>
              <select value={upsLoja} onChange={e=>{setUpsLoja(e.target.value);setUpsPag(0);}} style={inp}>
                <option value="">Todas as lojas</option>
                {lojas.map(l=><option key={l} value={l}>{l}</option>)}
              </select>
              <select value={upsVinculo} onChange={e=>{setUpsVinculo(e.target.value);setUpsPag(0);}} style={inp}>
                <option value="">Vinculados e não</option>
                <option value="sem">Sem vínculo</option>
                <option value="com">Com vínculo</option>
              </select>
            </div>
            <div style={{fontSize:11,color:"#64748b",marginBottom:8}}>{upsCarregando?"Carregando…":`${ups.total} pedido(s)`}</div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {ups.linhas.map(o=>{
                const st = o.status_upseller||"sem_status";
                const [c1,c2] = corSt[st]||["#475569","#f1f5f9"];
                const cancel = o.cancelando||o.excluido_em||st==="cancelado";
                return (
                  <div key={o.id} style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",opacity:cancel?0.7:1}}>
                    <div style={{display:"flex",justifyContent:"space-between",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                      <div style={{fontSize:12,fontWeight:700,color:"#1e293b"}}>{o.id_pedido_plataforma} <span style={{fontWeight:500,color:"#64748b"}}>· {o.destinatario||"—"}</span></div>
                      <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                        {o.loja&&<span style={{fontSize:10,color:"#64748b"}}>{o.loja}</span>}
                        <span style={{fontSize:10,fontWeight:700,color:c1,background:c2,borderRadius:6,padding:"2px 8px"}}>{(STATUS.find(x=>x[0]===st)||[0,st])[1]}{o.cancelando&&st!=="cancelado"?" (cancelando)":""}</span>
                        {o.baixado&&<span style={{fontSize:10,fontWeight:700,color:"#065f46",background:"#d1fae5",borderRadius:6,padding:"2px 8px"}}>📦 Baixado</span>}
                        {o.estornado&&<span style={{fontSize:10,fontWeight:700,color:"#92400e",background:"#fef3c7",borderRadius:6,padding:"2px 8px"}}>↩ Estornado</span>}
                      </div>
                    </div>
                    <div style={{fontSize:10,color:"#94a3b8",margin:"4px 0 8px"}}>{o.data_pedido?new Date(o.data_pedido).toLocaleString("pt-BR"):""} · qtd {o.quantidade||1}</div>
                    {(o.componentes||[]).length===0&&<div style={{fontSize:12,color:"#94a3b8"}}><em>sem variação</em></div>}
                    <div style={{display:"flex",flexDirection:"column",gap:6}}>
                      {(o.componentes||[]).map(cp=>(
                        <div key={cp.chave} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",background:"#f8fafc",borderRadius:8,padding:"6px 10px"}}>
                          <div style={{flex:1,minWidth:180,fontSize:12,color:"#334155",wordBreak:"break-word"}}>{cp.texto}{cp.mult>1?<strong> ×{cp.mult}</strong>:null}</div>
                          {cp.ignorar ? (
                            <>
                              <span style={{fontSize:11,fontWeight:600,color:"#64748b",background:"#f1f5f9",border:"1px solid #e2e8f0",borderRadius:8,padding:"3px 10px"}}>🚫 Sem material</span>
                              <button onClick={()=>{setUpsVincular({chave:cp.chave,variacao:cp.texto,variacao_id:null,ignorar:true});setUpsEscolha("");}} style={{background:"#f1f5f9",color:"#374151",border:"none",borderRadius:8,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>Alterar</button>
                            </>
                          ) : cp.variacao_id ? (
                            <>
                              <span style={{fontSize:11,fontWeight:600,color:"#065f46",background:"#ecfdf5",border:"1px solid #a7f3d0",borderRadius:8,padding:"3px 10px"}}>🔗 {cp.prod_nome}{cp.var_nome?` — ${cp.var_nome}`:""}</span>
                              <button onClick={()=>{setUpsVincular({chave:cp.chave,variacao:cp.texto,variacao_id:cp.variacao_id,ignorar:false});setUpsEscolha(cp.variacao_id);}} style={{background:"#f1f5f9",color:"#374151",border:"none",borderRadius:8,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>Alterar</button>
                            </>
                          ) : (
                            <button onClick={()=>{setUpsVincular({chave:cp.chave,variacao:cp.texto,variacao_id:null,ignorar:false});setUpsEscolha("");}} style={{background:"#1d4ed8",color:"#fff",border:"none",borderRadius:8,padding:"4px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🔗 Vincular</button>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
              {!upsCarregando&&ups.linhas.length===0&&<div style={{textAlign:"center",padding:"28px 0",color:"#94a3b8",fontSize:13}}>Nenhum pedido encontrado.</div>}
            </div>
            <div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:12,marginTop:14}}>
              <button disabled={upsPag===0} onClick={()=>setUpsPag(p=>p-1)} style={{padding:"6px 14px",border:"1px solid #e2e8f0",borderRadius:8,background:"#fff",cursor:upsPag===0?"default":"pointer",opacity:upsPag===0?0.4:1}}>← Anterior</button>
              <span style={{fontSize:12,color:"#64748b"}}>Página {upsPag+1} de {totalPag}</span>
              <button disabled={upsPag+1>=totalPag} onClick={()=>setUpsPag(p=>p+1)} style={{padding:"6px 14px",border:"1px solid #e2e8f0",borderRadius:8,background:"#fff",cursor:upsPag+1>=totalPag?"default":"pointer",opacity:upsPag+1>=totalPag?0.4:1}}>Próxima →</button>
            </div>

            {upsVincular&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:520,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                  <h3 style={{margin:"0 0 6px",fontSize:16,fontWeight:700,color:"#1e293b"}}>🔗 Vincular variação</h3>
                  <div style={{fontSize:12,background:"#f8fafc",borderRadius:8,padding:"8px 10px",marginBottom:12,wordBreak:"break-word"}}><strong>Parte do pedido:</strong> {upsVincular.variacao}</div>
                  <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Produto / variação cadastrada</label>
                  <select value={upsEscolha} onChange={e=>setUpsEscolha(e.target.value)} style={{...inp,width:"100%"}}>
                    <option value="">— Sem vínculo —</option>
                    <option value="__ignorar">🚫 Sem material (não dá baixa, ex: brinde)</option>
                    {opcoesVarTop.map(o=><option key={o.id} value={o.id}>{o.nome}</option>)}
                  </select>
                  {sugestao&&upsEscolha!==sugestao.id&&<button onClick={()=>setUpsEscolha(sugestao.id)} style={{marginTop:8,background:"#f5f3ff",color:"#6d28d9",border:"1px solid #ddd6fe",borderRadius:8,padding:"6px 10px",fontSize:11,fontWeight:700,cursor:"pointer"}}>💡 Sugestão: {sugestao.nome}</button>}
                  {upsEscolha!=="__ignorar"&&escolhida&&!escolhida.temEstoque&&<div style={{marginTop:8,fontSize:11,color:"#92400e",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"6px 10px"}}>⚠️ {escolhida.qtdMp===0?"Essa variação não tem matérias-primas.":"Nenhuma matéria-prima dessa variação está ligada a um item do Estoque."} Sem isso, a baixa não desconta nada. Ajuste em Produtos › 🧱 Matérias-primas.</div>}
                  <div style={{fontSize:11,color:"#64748b",marginTop:10}}>Pedidos com mais de um produto ("A + B") são separados em partes e cada parte tem seu vínculo. Todos os pedidos com essa parte, atuais e futuros, ficam vinculados. A baixa só roda quando todas as partes do pedido estiverem vinculadas (ou marcadas como sem material).</div>
                  <div style={{display:"flex",gap:10,marginTop:16}}>
                    <button onClick={()=>{setUpsVincular(null);setUpsEscolha("");}} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                    <button onClick={confirmarVinculoUps} disabled={upsSalvando||(!upsEscolha&&!upsVincular.variacao_id&&!upsVincular.ignorar)} style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:"#1d4ed8",color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",opacity:(!upsEscolha&&!upsVincular.variacao_id&&!upsVincular.ignorar)?0.4:1}}>{upsSalvando?"Salvando…":upsEscolha?"Vincular todos":"Remover vínculo"}</button>
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })()}

      {activeTab==="financeiro"&&isAdminEquipe&&!estoquePlusDesbloqueado&&(
        <div onClick={()=>{setActiveTab("abertos");setPinInput("");setPinErro("");}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:20,padding:32,maxWidth:340,width:"100%",boxShadow:"0 4px 20px rgba(0,0,0,0.08)",textAlign:"center",position:"relative"}}>
            <button onClick={()=>{setActiveTab("abertos");setPinInput("");setPinErro("");}} style={{position:"absolute",top:12,right:12,background:"none",border:"none",fontSize:18,color:"#94a3b8",cursor:"pointer",lineHeight:1,padding:4}}>✕</button>
            <div style={{fontSize:44,marginBottom:10}}>🔒</div>
            <h3 style={{margin:"0 0 6px",fontSize:17,fontWeight:800,color:"#1e293b"}}>Financeiro</h3>
            <p style={{margin:"0 0 20px",fontSize:12,color:"#64748b"}}>Área protegida. Digite o PIN para acessar.</p>
            {!estoquePlusConfig ? (
              <div>
                <p style={{fontSize:12,color:"#f59e0b",marginBottom:12}}>Configure o PIN pelo botão <strong>🔒 Estoque+</strong> na aba Estoque.</p>
                <button onClick={()=>{setActiveTab("estoque");setShowDefinirPin(true);}} style={{width:"100%",padding:"10px 0",background:"#059669",color:"#fff",border:"none",borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer"}}>Definir PIN agora</button>
              </div>
            ) : (
              <div>
                <input type="password" inputMode="numeric" maxLength={8} value={pinInput}
                  onChange={e=>{setPinInput(e.target.value);setPinErro("");}}
                  onKeyDown={async e=>{
                    if (e.key==="Enter") {
                      const hash = await crypto.subtle.digest("SHA-256",new TextEncoder().encode(pinInput)).then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
                      if (hash===estoquePlusConfig?.pin_hash){setEstoquePlusDesbloqueado(true);setPinInput("");setPinErro("");fetchCustosFixos(user.id,orgId).then(setCustosFixos);fetchMovimentosComPreco(user.id,orgId,finMes).then(setFinMovimentos);}
                      else{setPinErro("PIN incorreto.");setPinInput("");}
                    }
                  }}
                  autoFocus placeholder="••••"
                  style={{width:"100%",padding:"12px 16px",border:`2px solid ${pinErro?"#ef4444":"#e2e8f0"}`,borderRadius:12,fontSize:22,outline:"none",textAlign:"center",letterSpacing:8,boxSizing:"border-box",marginBottom:8}} />
                {pinErro&&<div style={{fontSize:12,color:"#ef4444",marginBottom:8}}>{pinErro}</div>}
                <button onClick={async()=>{
                  const hash = await crypto.subtle.digest("SHA-256",new TextEncoder().encode(pinInput)).then(b=>Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join(""));
                  if (hash===estoquePlusConfig?.pin_hash){setEstoquePlusDesbloqueado(true);setPinInput("");setPinErro("");fetchCustosFixos(user.id,orgId).then(setCustosFixos);fetchMovimentosComPreco(user.id,orgId,finMes).then(setFinMovimentos);}
                  else{setPinErro("PIN incorreto.");setPinInput("");}
                }} style={{width:"100%",padding:"12px 0",background:"#059669",color:"#fff",border:"none",borderRadius:12,fontSize:14,fontWeight:800,cursor:"pointer"}}>
                  Desbloquear
                </button>
              </div>
            )}
          </div>
        </div>
      )}
      {activeTab==="financeiro"&&isAdminEquipe&&estoquePlusDesbloqueado&&(()=>{
        const CATS = [
          {id:"aluguel",     icon:"🏠", label:"Aluguel"},
          {id:"luz",         icon:"💡", label:"Luz"},
          {id:"agua",        icon:"💧", label:"Água"},
          {id:"internet",    icon:"📶", label:"Internet"},
          {id:"funcionario", icon:"👤", label:"Funcionário"},
          {id:"cartao",      icon:"💳", label:"Cartão"},
          {id:"emprestimo",  icon:"🏦", label:"Empréstimo"},
          {id:"outros",      icon:"📋", label:"Outros"},
        ];
        const catInfo = (id) => CATS.find(c=>c.id===id)||{icon:"📋",label:id};

        const totalFixos = custosFixos.reduce((s,c)=>s+Number(c.valor),0);
        const totalMaterial = finMovimentos.reduce((s,m)=>s+(m.preco_unitario||0)*m.quantidade,0);
        const totalGeral = totalFixos + totalMaterial;

        const porCategoria = {};
        custosFixos.forEach(c=>{ porCategoria[c.categoria]=(porCategoria[c.categoria]||0)+Number(c.valor); });
        const porFornecedor = {};
        finMovimentos.forEach(m=>{ const n=m.fornecedores?.nome||"Sem fornecedor"; porFornecedor[n]=(porFornecedor[n]||0)+(m.preco_unitario||0)*m.quantidade; });

        const reloadFin = () => {
          fetchCustosFixos(user.id,orgId).then(setCustosFixos);
          fetchMovimentosComPreco(user.id,orgId,finMes).then(setFinMovimentos);
        };

        const abrirCustoForm = (c=null) => {
          setEditandoCusto(c);
          setCustoNome(c?.nome||""); setCustoCategoria(c?.categoria||"outros");
          setCustoValor(c?.valor?String(c.valor):""); setCustoRecorrente(c?.recorrente!==false);
          setCustoObs(c?.observacoes||"");
          setCustoVenc(c?.data_vencimento||"");
          setShowCustoForm(true);
        };
        // Vencimento: datas "YYYY-MM-DD" tratadas no fuso local (sem deslocar o dia)
        const parseData = (s) => { const [y,m,d]=String(s).slice(0,10).split("-").map(Number); return new Date(y,m-1,d); };
        const fmtData = (d) => String(d.getDate()).padStart(2,"0")+"/"+String(d.getMonth()+1).padStart(2,"0")+"/"+d.getFullYear();
        const infoVenc = (c, pago) => {
          if (!c.data_vencimento) return null;
          const base = parseData(c.data_vencimento);
          const agora = new Date(); const hoje = new Date(agora.getFullYear(),agora.getMonth(),agora.getDate());
          let alvo = base;
          if (c.recorrente!==false) {
            const [fy,fm] = finMes.split("-").map(Number);
            alvo = new Date(fy,fm-1,Math.min(base.getDate(),new Date(fy,fm,0).getDate()));
          }
          const dias = Math.round((alvo-hoje)/86400000);
          let txt, cor, bg;
          if (pago)         { txt="Pago · venc. "+fmtData(alvo).slice(0,5); cor="#065f46"; bg="#d1fae5"; }
          else if (dias<0)  { txt="Venceu em "+fmtData(alvo).slice(0,5); cor="#991b1b"; bg="#fde8e8"; }
          else if (dias===0){ txt="Vence hoje"; cor="#991b1b"; bg="#fde8e8"; }
          else if (dias<=5) { txt="Vence em "+dias+" dia"+(dias>1?"s":"")+" ("+fmtData(alvo).slice(0,5)+")"; cor="#92400e"; bg="#fef3c7"; }
          else              { txt="Vence "+fmtData(alvo).slice(0,5); cor="#475569"; bg="#f1f5f9"; }
          return {txt,cor,bg};
        };
        const alternarPago = async (c) => {
          const jaPago = custosPagos.includes(c.id);
          setCustosPagos(l=>jaPago?l.filter(x=>x!==c.id):[...l,c.id]); // atualiza na hora
          const res = await marcarCustoPago(user.id,orgId,c.id,finMes,!jaPago);
          if (!res.error) logOp({tipo:"financeiro",acao:jaPago?"despagar":"pagar",descricao:(jaPago?"Desmarcou pagamento: ":"Marcou como pago: ")+c.nome+" ("+finMes+")",tabela:"custos_fixos_pagamentos",registro_id:c.id,depois:{custo_id:c.id,mes:finMes}});
          if (res.error) {
            setCustosPagos(l=>jaPago?[...l,c.id]:l.filter(x=>x!==c.id));
            showToast("❌ "+res.error,"#991b1b");
          }
        };
        const salvarCusto = async () => {
          if (!custoNome.trim()||!custoValor||!custoVenc) return;
          setSalvandoCusto(true);
          const res = await salvarCustoFixo(user.id,orgId,{id:editandoCusto?.id,nome:custoNome,categoria:custoCategoria,valor:custoValor,recorrente:custoRecorrente,observacoes:custoObs,data_vencimento:custoVenc});
          setSalvandoCusto(false);
          if (res.error){showToast("❌ "+res.error,"#991b1b");return;}
          setShowCustoForm(false);
          logOp({tipo:"financeiro",acao:editandoCusto?"editar":"criar",descricao:(editandoCusto?"Custo fixo editado: ":"Custo fixo criado: ")+custoNome.trim()+" (R$ "+Number(custoValor).toFixed(2).replace(".",",")+")",tabela:"custos_fixos",registro_id:editandoCusto?.id||res.id,antes:editandoCusto?pick(editandoCusto,COLS_CUSTO):null});
          fetchCustosFixos(user.id,orgId).then(setCustosFixos);
          showToast(editandoCusto?"✅ Custo atualizado!":"✅ Custo cadastrado!","#059669");
        };
        const excluirCusto = async (id) => {
          if (!window.confirm("Remover este custo?")) return;
          const cAntes = custosFixos.find(x=>x.id===id);
          await excluirCustoFixo(id);
          logOp({tipo:"financeiro",acao:"excluir",descricao:"Custo fixo removido: "+(cAntes?.nome||""),tabela:"custos_fixos",registro_id:id,antes:pick(cAntes,COLS_CUSTO)});
          fetchCustosFixos(user.id,orgId).then(setCustosFixos);
          showToast("🗑️ Custo removido.","#64748b");
        };

        // Chat
        const contextoFinanceiro = {
          mes: finMes,
          hoje: new Date().toLocaleDateString("pt-BR"),
          custos_fixos: custosFixos.map(c=>{
            const pago = custosPagos.includes(c.id);
            return {nome:c.nome,categoria:c.categoria,valor:c.valor,recorrente:c.recorrente!==false,pago,situacao:infoVenc(c,pago)?.txt||"sem vencimento"};
          }),
          total_pago: custosFixos.filter(c=>custosPagos.includes(c.id)).reduce((s,c)=>s+Number(c.valor),0),
          total_a_pagar: custosFixos.filter(c=>!custosPagos.includes(c.id)).reduce((s,c)=>s+Number(c.valor),0),
          total_fixos: totalFixos,
          total_material: totalMaterial,
          total_geral: totalGeral,
          por_categoria: porCategoria,
          por_fornecedor: porFornecedor,
          produtos_estoque: estoqueProdutos,
        };

        const gerarAnalise = async () => {
          if (chatAnaliseLoading) return;
          setChatAnaliseLoading(true);
          setChatAnalise("");
          const res = await chamarClaudeFinanceiro([{role:"user",content:`Analise os custos de ${finMes} e me dê um resumo executivo com: 1) situação geral dos custos, 2) maior gasto, 3) um alerta ou dica de otimização. Seja direto, máx 4 linhas.`}], contextoFinanceiro);
          setChatAnaliseLoading(false);
          if (res.reply) setChatAnalise(res.reply);
          else setChatAnalise("⚠️ " + (res.error||"Erro ao gerar análise. Verifique a chave da API."));
        };

        const enviarChat = async () => {
          const msg = chatInput.trim();
          if (!msg || chatCarregando) return;
          setChatInput("");
          const novaMsgs = [...chatMsgs, {role:"user", content:msg}];
          setChatMsgs(novaMsgs);
          await salvarMensagemChat(user.id,orgId,"user",msg);
          setChatCarregando(true);
          const historico = novaMsgs.slice(-10).map(m=>({role:m.role,content:m.content}));
          const res = await chamarClaudeFinanceiro(historico, contextoFinanceiro);
          setChatCarregando(false);
          const reply = res.reply || ("⚠️ " + (res.error||"Erro desconhecido"));
          setChatMsgs(prev=>[...prev,{role:"assistant",content:reply}]);
          await salvarMensagemChat(user.id,orgId,"assistant",reply);
          setTimeout(()=>chatEndRef.current?.scrollIntoView({behavior:"smooth"}),100);
        };

        return (
          <div>
            {/* Header */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
              <h2 style={{margin:0,fontSize:18,fontWeight:800,color:"#059669"}}>💰 Financeiro</h2>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <input type="month" value={finMes} onChange={e=>{setFinMes(e.target.value);fetchMovimentosComPreco(user.id,orgId,e.target.value).then(setFinMovimentos);}}
                  style={{padding:"7px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:12,outline:"none"}} />
                <button onClick={exportarContador} style={{background:"#eff6ff",color:"#1d4ed8",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📊 Exportar p/ contador</button>
                <button onClick={reloadFin} style={{background:"#d1fae5",color:"#065f46",border:"none",borderRadius:8,padding:"7px 12px",fontSize:12,cursor:"pointer"}}>🔄</button>
                <button onClick={()=>{setEstoquePlusDesbloqueado(false);}} style={{background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:10,padding:"7px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>🔒 Trancar</button>
              </div>
            </div>

            {/* Cards resumo */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:12,marginBottom:20}}>
              <div style={{background:"linear-gradient(135deg,#059669,#0d9488)",borderRadius:14,padding:"16px 18px",color:"#fff"}}>
                <div style={{fontSize:11,opacity:.8,marginBottom:4}}>📋 Custos fixos</div>
                <div style={{fontSize:24,fontWeight:800}}>R$ {totalFixos.toFixed(2).replace(".",",")}</div>
              </div>
              <div style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",borderRadius:14,padding:"16px 18px",color:"#fff"}}>
                <div style={{fontSize:11,opacity:.8,marginBottom:4}}>📦 Material (estoque)</div>
                <div style={{fontSize:24,fontWeight:800}}>R$ {totalMaterial.toFixed(2).replace(".",",")}</div>
              </div>
              <div style={{background:"linear-gradient(135deg,#dc2626,#b91c1c)",borderRadius:14,padding:"16px 18px",color:"#fff"}}>
                <div style={{fontSize:11,opacity:.8,marginBottom:4}}>💸 Custo total {finMes}</div>
                <div style={{fontSize:24,fontWeight:800}}>R$ {totalGeral.toFixed(2).replace(".",",")}</div>
              </div>
            </div>

            {/* Análise automática do Claude */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>🤖 Análise do Gemini</div>
                <button onClick={gerarAnalise} disabled={chatAnaliseLoading}
                  style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",color:"#fff",border:"none",borderRadius:8,padding:"6px 14px",fontSize:11,fontWeight:700,cursor:"pointer",opacity:chatAnaliseLoading?.6:1}}>
                  {chatAnaliseLoading?"⏳ Analisando...":"✨ Gerar análise"}
                </button>
              </div>
              {chatAnalise ? (
                <div style={{background:"#f5f3ff",border:"1px solid #ddd6fe",borderRadius:10,padding:"12px 14px",fontSize:13,color:"#374151",lineHeight:1.6,whiteSpace:"pre-wrap"}}>
                  {chatAnalise}
                </div>
              ) : (
                <div style={{textAlign:"center",padding:"12px 0",fontSize:12,color:"#94a3b8"}}>Clique em "Gerar análise" para o Gemini analisar seus custos de {finMes}.</div>
              )}
            </div>

            {/* Custos fixos */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>📋 Custos Fixos</div>
                <button onClick={()=>abrirCustoForm()} style={{background:"#059669",color:"#fff",border:"none",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:700,cursor:"pointer"}}>+ Adicionar</button>
              </div>
              {custosFixos.length===0 ? (
                <div style={{textAlign:"center",padding:"20px 0",color:"#94a3b8",fontSize:12}}>Nenhum custo fixo cadastrado.</div>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {CATS.filter(cat=>custosFixos.some(c=>c.categoria===cat.id)).map(cat=>{
                    const itens = custosFixos.filter(c=>c.categoria===cat.id);
                    const subtotal = itens.reduce((s,c)=>s+Number(c.valor),0);
                    return (
                      <div key={cat.id}>
                        <div style={{fontSize:11,fontWeight:700,color:"#64748b",padding:"4px 0 2px",display:"flex",justifyContent:"space-between"}}>
                          <span>{cat.icon} {cat.label}</span>
                          <span>R$ {subtotal.toFixed(2).replace(".",",")}</span>
                        </div>
                        {itens.map(c=>(
                          <div key={c.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 10px",background:"#f8fafc",borderRadius:8,marginBottom:3}}>
                            <div>
                              <span style={{fontSize:13,color:"#374151"}}>{c.nome}</span>
                              {!c.recorrente&&<span style={{fontSize:10,background:"#fef3c7",color:"#92400e",borderRadius:6,padding:"1px 6px",marginLeft:6}}>único</span>}
                              {(()=>{const v=infoVenc(c,custosPagos.includes(c.id));return v?<div style={{marginTop:3}}><span style={{fontSize:10,fontWeight:700,background:v.bg,color:v.cor,borderRadius:6,padding:"2px 7px"}}>📅 {v.txt}</span></div>:<div style={{marginTop:3}}><span style={{fontSize:10,color:"#94a3b8"}}>sem vencimento — clique em ✏️ para definir</span></div>;})()}
                              {c.observacoes&&<div style={{fontSize:10,color:"#94a3b8",fontStyle:"italic"}}>{c.observacoes}</div>}
                            </div>
                            <div style={{display:"flex",alignItems:"center",gap:8}}>
                              <span style={{fontSize:13,fontWeight:700,color:"#059669"}}>R$ {Number(c.valor).toFixed(2).replace(".",",")}</span>
                              {(()=>{const pago=custosPagos.includes(c.id);return (
                                <button onClick={()=>alternarPago(c)} title={pago?"Pago — clique para desmarcar":"Não pago — clique para marcar como pago"}
                                  style={{padding:"3px 9px",background:pago?"#d1fae5":"#fde8e8",color:pago?"#065f46":"#991b1b",border:`1.5px solid ${pago?"#6ee7b7":"#fecaca"}`,borderRadius:6,fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>
                                  {pago?"✅ Pago":"❌ Não pago"}
                                </button>
                              );})()}
                              <button onClick={()=>abrirCustoForm(c)} style={{padding:"3px 7px",background:"#f1f5f9",border:"none",borderRadius:6,fontSize:11,cursor:"pointer"}}>✏️</button>
                              <button onClick={()=>excluirCusto(c.id)} style={{padding:"3px 7px",background:"#fde8e8",color:"#991b1b",border:"none",borderRadius:6,fontSize:11,cursor:"pointer"}}>🗑️</button>
                            </div>
                          </div>
                        ))}
                      </div>
                    );
                  })}
                  <div style={{display:"flex",justifyContent:"space-between",padding:"10px 10px",background:"#d1fae5",borderRadius:8,marginTop:4}}>
                    <span style={{fontSize:13,fontWeight:700,color:"#065f46"}}>Total fixos</span>
                    <span style={{fontSize:14,fontWeight:800,color:"#065f46"}}>R$ {totalFixos.toFixed(2).replace(".",",")}</span>
                  </div>
                  {(()=>{const pagoTot=custosFixos.filter(c=>custosPagos.includes(c.id)).reduce((s,c)=>s+Number(c.valor),0);return (
                    <div style={{display:"flex",justifyContent:"space-between",gap:8,flexWrap:"wrap",padding:"4px 10px",fontSize:12}}>
                      <span style={{color:"#065f46",fontWeight:700}}>✅ Pago: R$ {pagoTot.toFixed(2).replace(".",",")}</span>
                      <span style={{color:"#991b1b",fontWeight:700}}>❌ Falta pagar: R$ {(totalFixos-pagoTot).toFixed(2).replace(".",",")}</span>
                    </div>
                  );})()}
                </div>
              )}
            </div>

            {/* Material usado no mês (saídas do estoque) */}
            {(()=>{
              const fq = (n) => Number(Number(n).toFixed(3)).toString().replace(".",",");
              const fm = (n) => "R$ "+Number(n||0).toFixed(2).replace(".",",");
              const mapa = {};
              finSaidas.forEach(m=>{
                const nome = m.estoque_produtos?.nome||"(material removido)";
                if(!mapa[nome]) mapa[nome]={nome,un:m.estoque_produtos?.unidade||"un",qtd:0,custo:0,saidas:0};
                const q = Number(m.quantidade)||0;
                mapa[nome].qtd += q; mapa[nome].saidas += 1;
                mapa[nome].custo += q*(Number(m.estoque_produtos?.preco_unitario)||0);
              });
              const lista = Object.values(mapa).sort((a,b)=>a.nome.localeCompare(b.nome));
              const totalCusto = lista.reduce((a,x)=>a+x.custo,0);
              return (
                <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
                  <div style={{fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:4}}>📤 Material usado em {finMes}</div>
                  <div style={{fontSize:11,color:"#94a3b8",marginBottom:12}}>Tudo que saiu do estoque no mês, por material.</div>
                  {lista.length===0 ? (
                    <div style={{textAlign:"center",padding:"16px 0",color:"#94a3b8",fontSize:12}}>Nenhuma saída de estoque registrada neste mês.</div>
                  ) : (
                    <>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))",gap:8}}>
                        {lista.map(x=>(
                          <div key={x.nome} style={{background:"#f0f9ff",border:"1px solid #bae6fd",borderRadius:10,padding:"10px 12px"}}>
                            <div style={{fontSize:20,fontWeight:800,color:"#0369a1"}}>{fq(x.qtd)} <span style={{fontSize:13,fontWeight:700}}>{x.un}</span></div>
                            <div style={{fontSize:12,fontWeight:600,color:"#1e293b",marginTop:2}}>{x.nome}</div>
                            <div style={{fontSize:10,color:"#64748b",marginTop:2}}>{x.saidas} saída{x.saidas>1?"s":""}{x.custo>0?` · ${fm(x.custo)}`:""}</div>
                          </div>
                        ))}
                      </div>
                      {totalCusto>0&&(
                        <div style={{display:"flex",justifyContent:"space-between",marginTop:10,padding:"8px 10px",background:"#e0f2fe",borderRadius:8,fontSize:13}}>
                          <strong style={{color:"#0c4a6e"}}>Custo do material usado</strong>
                          <strong style={{color:"#0369a1"}}>{fm(totalCusto)}</strong>
                        </div>
                      )}
                    </>
                  )}
                </div>
              );
            })()}

            {/* Calculadora de projeção (ponto de equilíbrio) */}
            {(()=>{
              const fm = (n) => "R$ "+Number(n||0).toFixed(2).replace(".",",");
              const estoqueVal = projEstoque!=="" ? (Number(projEstoque)||0) : totalMaterial;
              const outros = Number(projOutros)||0;
              const lucroDes = Number(projLucro)||0;
              const usaCompras = projBase==="compras";
              const jaAuto = pedidosMes?.pedidos||0;
              const ja = projJa!=="" ? (Number(projJa)||0) : jaAuto;
              // Saídas do estoque do mês, agrupadas por material e valorizadas pelo preço cadastrado no Estoque+
              const saidasMap = {};
              finSaidas.forEach(m=>{
                const nome = m.estoque_produtos?.nome||"(material removido)";
                const pr = Number(m.estoque_produtos?.preco_unitario)||0;
                const q = Number(m.quantidade)||0;
                if(!saidasMap[nome]) saidasMap[nome]={nome,un:m.estoque_produtos?.unidade||"",qtd:0,pr,custo:0};
                saidasMap[nome].qtd += q; saidasMap[nome].custo += q*pr;
              });
              const saidasLista = Object.values(saidasMap).sort((a,b)=>b.custo-a.custo);
              const totalSaidas = saidasLista.reduce((a,x)=>a+x.custo,0);
              const semPreco = saidasLista.filter(x=>!x.pr).length;
              const matPorPedido = ja>0 ? totalSaidas/ja : null;
              const saidasComoMeta = projBase==="saidas" && ja<=0;
              const meta = totalFixos + (usaCompras?estoqueVal:0) + (saidasComoMeta?totalSaidas:0) + outros + lucroDes;

              const todas = [];
              produtosCalc.forEach(p=>{
                const vs = p.produto_variacoes||[];
                vs.forEach(v=>{
                  const liq = liquidoDe(p,v.preco_venda);
                  const mp = custoMpDe(v);
                  todas.push({id:v.id, nome:p.nome+(vs.length>1?" · "+v.nome:""), preco:Number(v.preco_venda)||0, liq, mp, contrib: liq-(projBase==="produto"?mp:(projBase==="saidas"&&matPorPedido!==null?matPorPedido:0))});
                });
              });
              const validas = todas.filter(x=>x.preco>0&&x.contrib>0);
              const peso = (x) => { const m=projMix[x.id]; return (m===undefined||m==="")?null:Math.max(0,Number(m)||0); };
              const algumMix = validas.some(x=>peso(x)!==null);
              const pesos = validas.map(x=>algumMix?(peso(x)||0):1);
              const soma = pesos.reduce((a,b)=>a+b,0);
              const contribMedia = soma>0 ? validas.reduce((a,x,i)=>a+x.contrib*pesos[i],0)/soma : 0;
              const precoMedio  = soma>0 ? validas.reduce((a,x,i)=>a+x.preco*pesos[i],0)/soma : 0;
              const unidades = contribMedia>0 ? Math.ceil(meta/contribMedia) : null;
              const faturamento = unidades!==null ? unidades*precoMedio : null;
              const faltam = unidades!==null ? Math.max(0,unidades-ja) : null;
              const [fy,fmo] = finMes.split("-").map(Number);
              const diasMes = new Date(fy,fmo,0).getDate();
              const hj = new Date();
              const ehMesAtual = hj.getFullYear()===fy && hj.getMonth()+1===fmo;
              const diasRest = ehMesAtual ? diasMes-hj.getDate()+1 : diasMes;
              const inp = {width:"100%",padding:"8px 10px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:13,outline:"none",boxSizing:"border-box"};
              const lb = {fontSize:10,fontWeight:600,color:"#64748b",display:"block",marginBottom:2};
              return (
                <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,flexWrap:"wrap",marginBottom:4}}>
                    <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>🎯 Calculadora de projeção — {finMes}</div>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      {projSalvaEm&&<span style={{fontSize:10,color:"#64748b"}}>salva {quandoFmt(projSalvaEm.quando)}{projSalvaEm.quem?` por ${projSalvaEm.quem}`:""}</span>}
                      <button onClick={salvarProjecaoMes} disabled={salvandoProj} style={{background:"#059669",color:"#fff",border:"none",borderRadius:8,padding:"5px 12px",fontSize:11,fontWeight:700,cursor:"pointer"}}>{salvandoProj?"Salvando…":"💾 Salvar projeção"}</button>
                    </div>
                  </div>
                  <div style={{fontSize:11,color:"#94a3b8",marginBottom:12}}>Quanto você precisa vender no mês para cobrir os custos (ponto de equilíbrio).</div>

                  {/* O que precisa ser coberto */}
                  <div style={{background:"#f8fafc",borderRadius:10,padding:"12px 14px",marginBottom:12}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:8}}>1) O que precisa ser pago</div>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:13,marginBottom:6}}>
                      <span>📋 Custos fixos</span><strong>{fm(totalFixos)}</strong>
                    </div>
                    <div style={{marginBottom:6}}>
                      <label style={lb}>Como considerar o material (estoque)?</label>
                      <select value={projBase} onChange={e=>setProjBase(e.target.value)} style={{...inp,background:"#fff"}}>
                        <option value="compras">📦 Compras de estoque do mês (o que gastei comprando)</option>
                        <option value="saidas">📤 Consumo real: saídas do estoque (o que usei nos pedidos)</option>
                        <option value="produto">🧱 Custo de matéria-prima de cada produto (receita)</option>
                      </select>
                    </div>
                    {usaCompras&&(
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,marginBottom:4}}>
                        <span style={{fontSize:13}}>📦 Compras do mês</span>
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          <span style={{fontSize:11,color:"#94a3b8"}}>R$</span>
                          <input type="number" min="0" step="0.01" value={projEstoque} onChange={e=>setProjEstoque(e.target.value)} placeholder={totalMaterial.toFixed(2)} style={{...inp,width:110}} />
                        </div>
                      </div>
                    )}
                    <div style={{fontSize:10,color:"#94a3b8",margin:"0 0 8px"}}>
                      {projBase==="compras" && `Gasto real do mês: ${fm(totalMaterial)} (edite para simular outro valor). O custo de matéria-prima por unidade não é descontado de novo.`}
                      {projBase==="saidas" && (ja>0
                        ? `Saídas do mês ${fm(totalSaidas)} ÷ ${ja} pedidos = ${fm(matPorPedido)} de material por pedido, descontado de cada venda.`
                        : `Informe os pedidos vendidos no mês (seção 2) para virar custo por pedido. Enquanto isso, as saídas (${fm(totalSaidas)}) entram como valor fixo na meta.`)}
                      {projBase==="produto" && "O custo de matéria-prima de cada produto é descontado por unidade vendida."}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                      <div>
                        <label style={lb}>Outros custos do mês (R$)</label>
                        <input type="number" min="0" step="0.01" value={projOutros} onChange={e=>setProjOutros(e.target.value)} placeholder="0,00" style={inp} />
                      </div>
                      <div>
                        <label style={lb}>Lucro que quero tirar (R$)</label>
                        <input type="number" min="0" step="0.01" value={projLucro} onChange={e=>setProjLucro(e.target.value)} placeholder="0,00" style={inp} />
                      </div>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",marginTop:10,paddingTop:8,borderTop:"1px solid #e2e8f0",fontSize:13}}>
                      <strong style={{color:"#1e293b"}}>Meta a cobrir</strong><strong style={{color:"#7c3aed",fontSize:15}}>{fm(meta)}</strong>
                    </div>
                  </div>

                  {/* Mix de produtos */}
                  <div style={{background:"#f0f9ff",border:"1px solid #bae6fd",borderRadius:10,padding:"12px 14px",marginBottom:12}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#0369a1",marginBottom:8}}>2) Material realmente usado (saídas do estoque em {finMes})</div>
                    {saidasLista.length===0 ? (
                      <div style={{fontSize:12,color:"#94a3b8"}}>Nenhuma saída de estoque registrada neste mês.</div>
                    ) : (
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead><tr>{["Material","Saiu","Preço unit.","Custo"].map(h=><th key={h} style={{...TH,padding:"5px 8px"}}>{h}</th>)}</tr></thead>
                          <tbody>
                            {saidasLista.map(x=>(
                              <tr key={x.nome} style={{borderBottom:"1px solid #e0f2fe"}}>
                                <td style={{padding:"5px 8px",fontWeight:600}}>{x.nome}</td>
                                <td style={{padding:"5px 8px"}}>{x.qtd} {x.un}</td>
                                <td style={{padding:"5px 8px"}}>{x.pr?fm(x.pr):<span style={{color:"#f59e0b"}}>sem preço</span>}</td>
                                <td style={{padding:"5px 8px",fontWeight:700,color:"#0369a1"}}>{fm(x.custo)}</td>
                              </tr>
                            ))}
                            <tr style={{background:"#e0f2fe",fontWeight:800}}>
                              <td colSpan={3} style={{padding:"6px 8px"}}>Total usado no mês</td>
                              <td style={{padding:"6px 8px",color:"#0369a1"}}>{fm(totalSaidas)}</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    )}
                    {semPreco>0&&<div style={{fontSize:10,color:"#f59e0b",marginTop:6}}>⚠️ {semPreco} material(is) sem preço no Estoque+ — o custo deles está contando como R$ 0,00.</div>}
                    <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginTop:10}}>
                      <label style={{fontSize:11,fontWeight:600,color:"#0369a1"}}>Pedidos do mês (UpSeller):</label>
                      <input type="number" min="0" value={projJa} onChange={e=>setProjJa(e.target.value)} placeholder={pedidosMes?`${jaAuto} (automático)`:"Ex: 120"} style={{...inp,width:130}} />
                      {pedidosMes&&<span style={{fontSize:10,color:"#64748b"}}>{projJa===""?"usando o valor automático":"valor manual"} · {pedidosMes.cancelados||0} cancelado(s) fora da conta{projJa!==""&&<> · <a href="#" onClick={e=>{e.preventDefault();setProjJa("");}} style={{color:"#0369a1"}}>voltar ao automático</a></>}</span>}
                    </div>
                    {matPorPedido!==null&&(
                      <div style={{marginTop:8,fontSize:13,color:"#0c4a6e"}}>
                        Para <strong>{ja}</strong> pedidos foram usados <strong>{fm(totalSaidas)}</strong> de material → <strong style={{fontSize:15}}>{fm(matPorPedido)} por pedido</strong>
                      </div>
                    )}
                  </div>

                  <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:6}}>3) Quanto cada venda deixa para cobrir os custos</div>
                  {todas.length===0 ? (
                    <div style={{fontSize:12,color:"#94a3b8",padding:"8px 0 12px"}}>Cadastre produtos com preço na aba Produtos para calcular.</div>
                  ) : (
                    <div style={{overflowX:"auto",marginBottom:6}}>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                        <thead><tr>{["Produto / variação","Preço","Sobra por venda","% das vendas","Só vendendo este"].map(h=><th key={h} style={{...TH,padding:"6px 8px"}}>{h}</th>)}</tr></thead>
                        <tbody>
                          {todas.map(x=>{
                            const ok = x.preco>0&&x.contrib>0;
                            const un = ok ? Math.ceil(meta/x.contrib) : null;
                            return (
                              <tr key={x.id} style={{borderBottom:"1px solid #f1f5f9",opacity:ok?1:0.6}}>
                                <td style={{padding:"6px 8px",fontWeight:600}}>{x.nome}</td>
                                <td style={{padding:"6px 8px"}}>{fm(x.preco)}</td>
                                <td style={{padding:"6px 8px",fontWeight:700,color:x.contrib>0?"#059669":"#ef4444"}}>{fm(x.contrib)}</td>
                                <td style={{padding:"6px 8px",width:90}}>
                                  {ok ? <input type="number" min="0" value={projMix[x.id]??""} onChange={e=>setProjMix(m=>({...m,[x.id]:e.target.value}))} placeholder="auto" style={{...inp,padding:"5px 8px",fontSize:12}} /> : <span style={{fontSize:10,color:"#ef4444"}}>sem lucro</span>}
                                </td>
                                <td style={{padding:"6px 8px",color:"#374151"}}>{ok ? <><strong>{un}</strong> un · {fm(un*x.preco)}</> : "—"}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                  <div style={{fontSize:10,color:"#94a3b8",marginBottom:12}}>"Sobra por venda" = preço − comissão, imposto, adiantamento e taxa fixa da Shopee (os % de cada produto){projBase==="produto"?" − custo de matéria-prima do produto":(projBase==="saidas"&&matPorPedido!==null?" − material médio por pedido (saídas do estoque)":"")}. Deixe os % em branco para dividir as vendas igualmente entre os produtos; ou informe a proporção que você vende (ex: 70 e 30).</div>

                  {/* Resultado */}
                  <div style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",borderRadius:12,padding:"14px 16px",color:"#fff"}}>
                    {unidades===null ? (
                      <div style={{fontSize:13}}>Nenhum produto com lucro por venda positivo. Confira preços e custos na aba Produtos.</div>
                    ) : (
                      <>
                        <div style={{fontSize:11,opacity:.85,marginBottom:6}}>Para cobrir {fm(meta)} você precisa vender no mês:</div>
                        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:12}}>
                          <div><div style={{fontSize:22,fontWeight:800}}>{unidades}</div><div style={{fontSize:10,opacity:.85}}>unidades (≈ {(unidades/diasMes).toFixed(1)}/dia no mês todo)</div></div>
                          <div><div style={{fontSize:22,fontWeight:800}}>{fm(faturamento)}</div><div style={{fontSize:10,opacity:.85}}>de faturamento (preço médio {fm(precoMedio)})</div></div>
                          <div><div style={{fontSize:22,fontWeight:800}}>{fm(contribMedia)}</div><div style={{fontSize:10,opacity:.85}}>sobra média por venda</div></div>
                        </div>
                        <div style={{marginTop:12,paddingTop:10,borderTop:"1px solid rgba(255,255,255,.25)",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                          <label style={{fontSize:11,opacity:.9}}>Já vendi no mês (unidades):</label>
                          <input type="number" min="0" value={projJa} onChange={e=>setProjJa(e.target.value)} placeholder={String(jaAuto)} style={{width:90,padding:"6px 8px",border:"none",borderRadius:8,fontSize:13,outline:"none"}} />
                          {ja>0&&(
                            <span style={{fontSize:12,fontWeight:700}}>
                              {faltam===0 ? "✅ Meta batida!" : `Faltam ${faltam} un · ${(faltam/Math.max(diasRest,1)).toFixed(1)}/dia nos ${diasRest} dias restantes`}
                            </span>
                          )}
                        </div>
                      </>
                    )}
                  </div>
                </div>
              );
            })()}

            {/* Material por fornecedor */}
            {Object.keys(porFornecedor).length>0&&(
              <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:10}}>📦 Material por fornecedor</div>
                <div style={{display:"flex",flexDirection:"column",gap:5}}>
                  {Object.entries(porFornecedor).sort((a,b)=>b[1]-a[1]).map(([nome,val])=>(
                    <div key={nome} style={{display:"flex",justifyContent:"space-between",padding:"7px 10px",background:"#f8fafc",borderRadius:8}}>
                      <span style={{fontSize:13,color:"#374151"}}>{nome}</span>
                      <span style={{fontSize:13,fontWeight:700,color:"#7c3aed"}}>R$ {val.toFixed(2).replace(".",",")}</span>
                    </div>
                  ))}
                  <div style={{display:"flex",justifyContent:"space-between",padding:"8px 10px",background:"#ede9fe",borderRadius:8,marginTop:2}}>
                    <span style={{fontSize:13,fontWeight:700,color:"#6d28d9"}}>Total material</span>
                    <span style={{fontSize:14,fontWeight:800,color:"#6d28d9"}}>R$ {totalMaterial.toFixed(2).replace(".",",")}</span>
                  </div>
                </div>
              </div>
            )}

            {/* Upload do ERP — Cálculo de lucro */}
            <div style={{background:"#fff",borderRadius:14,padding:"16px 18px",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",marginBottom:20}}>
              <div style={{fontSize:13,fontWeight:700,color:"#1e293b",marginBottom:4}}>📊 Resultado do mês — Print do ERP</div>
              <div style={{fontSize:11,color:"#64748b",marginBottom:12}}>Carregue um print/screenshot do seu ERP mostrando o total de vendas válidas. O Claude irá ler os valores e calcular seu lucro líquido descontando custos.</div>
              <input ref={erpFileRef} type="file" accept="image/*" style={{display:"none"}} onChange={async e=>{
                const file = e.target.files[0];
                if (!file) return;
                const url = URL.createObjectURL(file);
                setErpImagemPreview(url);
                setErpResultado(null);
                // converte para base64
                const reader = new FileReader();
                reader.onload = ev => {
                  const base64 = ev.target.result.split(",")[1];
                  setErpVendas({base64, tipo: file.type});
                };
                reader.readAsDataURL(file);
              }} />
              <div style={{display:"flex",gap:10,alignItems:"flex-start",flexWrap:"wrap"}}>
                <div>
                  <button onClick={()=>erpFileRef.current?.click()}
                    style={{background:"#ede9fe",color:"#6d28d9",border:"2px dashed #a78bfa",borderRadius:10,padding:"10px 18px",fontSize:12,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:7}}>
                    📎 Carregar print do ERP
                  </button>
                  {erpImagemPreview&&(
                    <div style={{marginTop:8,position:"relative",display:"inline-block"}}>
                      <img src={erpImagemPreview} alt="ERP preview" style={{maxWidth:220,maxHeight:140,borderRadius:10,border:"1.5px solid #e2e8f0",objectFit:"contain"}} />
                      <button onClick={()=>{setErpImagemPreview(null);setErpVendas(null);setErpResultado(null);if(erpFileRef.current)erpFileRef.current.value="";}}
                        style={{position:"absolute",top:4,right:4,background:"#ef4444",color:"#fff",border:"none",borderRadius:"50%",width:20,height:20,fontSize:11,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
                    </div>
                  )}
                </div>
                {erpVendas&&(
                  <button onClick={async()=>{
                    setErpAnalisando(true);
                    setErpResultado(null);
                    // Chama Claude com a imagem + contexto financeiro
                    const prompt = `Analise este print de ERP/sistema de vendas. Extraia o total de vendas do mês e calcule o lucro líquido usando estas informações da loja:\n\n- Mês: ${finMes}\n- Total de custos fixos: R$ ${totalFixos.toFixed(2)}\n- Total gasto com material/estoque: R$ ${totalMaterial.toFixed(2)}\n- Custo total: R$ ${totalGeral.toFixed(2)}\n\nResponda em JSON com o formato:\n{"total_vendas": número, "lucro_bruto": número, "lucro_liquido": número, "margem_percent": número, "observacao": "texto"}\n\nOnde:\n- total_vendas = valor total de vendas que você leu na imagem\n- lucro_bruto = total_vendas - total_material\n- lucro_liquido = total_vendas - custo_total (fixos + material)\n- margem_percent = (lucro_liquido / total_vendas) * 100\n\nSe não conseguir ler o valor na imagem, coloque 0 em total_vendas e explique na observação.`;
                    try {
                      const { data: { session } } = await supabase.auth.getSession();
                      const token = session?.access_token;
                      const resp = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/claude-financeiro`, {
                        method:"POST",
                        headers:{"Content-Type":"application/json","Authorization":`Bearer ${token}`},
                        body: JSON.stringify({
                          messages:[{role:"user",content:[
                            {type:"image",source:{type:"base64",media_type:erpVendas.tipo,data:erpVendas.base64}},
                            {type:"text",text:prompt}
                          ]}],
                          contexto:{mes:finMes,total_fixos:totalFixos,total_material:totalMaterial,total_geral:totalGeral}
                        })
                      });
                      const json = await resp.json();
                      const reply = json.reply || "";
                      // tenta extrair JSON da resposta
                      const match = reply.match(/\{[\s\S]*\}/);
                      if (match) {
                        try { setErpResultado(JSON.parse(match[0])); }
                        catch { setErpResultado({total_vendas:0,lucro_bruto:0,lucro_liquido:0,margem_percent:0,observacao:reply}); }
                      } else {
                        setErpResultado({total_vendas:0,lucro_bruto:0,lucro_liquido:0,margem_percent:0,observacao:reply||"Não foi possível analisar a imagem."});
                      }
                    } catch(err) {
                      setErpResultado({total_vendas:0,lucro_bruto:0,lucro_liquido:0,margem_percent:0,observacao:"Erro: "+err.message});
                    }
                    setErpAnalisando(false);
                  }} disabled={erpAnalisando}
                    style={{padding:"10px 18px",background:erpAnalisando?"#e2e8f0":"linear-gradient(135deg,#7c3aed,#059669)",color:erpAnalisando?"#94a3b8":"#fff",border:"none",borderRadius:10,fontSize:12,fontWeight:700,cursor:"pointer",alignSelf:"flex-start"}}>
                    {erpAnalisando?"⏳ Calculando...":"🤖 Calcular lucro"}
                  </button>
                )}
              </div>
              {erpResultado&&(
                <div style={{marginTop:14}}>
                  {erpResultado.total_vendas>0 ? (
                    <div>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:10,marginBottom:12}}>
                        <div style={{background:"#d1fae5",borderRadius:12,padding:"12px 14px"}}>
                          <div style={{fontSize:10,color:"#065f46",fontWeight:600,marginBottom:3}}>💰 Vendas totais</div>
                          <div style={{fontSize:20,fontWeight:800,color:"#065f46"}}>R$ {Number(erpResultado.total_vendas).toFixed(2).replace(".",",")}</div>
                        </div>
                        <div style={{background:erpResultado.lucro_liquido>=0?"#d1fae5":"#fee2e2",borderRadius:12,padding:"12px 14px"}}>
                          <div style={{fontSize:10,color:erpResultado.lucro_liquido>=0?"#065f46":"#991b1b",fontWeight:600,marginBottom:3}}>💸 Lucro líquido</div>
                          <div style={{fontSize:20,fontWeight:800,color:erpResultado.lucro_liquido>=0?"#065f46":"#991b1b"}}>R$ {Number(erpResultado.lucro_liquido).toFixed(2).replace(".",",")}</div>
                        </div>
                        <div style={{background:"#ede9fe",borderRadius:12,padding:"12px 14px"}}>
                          <div style={{fontSize:10,color:"#6d28d9",fontWeight:600,marginBottom:3}}>📈 Margem</div>
                          <div style={{fontSize:20,fontWeight:800,color:"#6d28d9"}}>{Number(erpResultado.margem_percent).toFixed(1)}%</div>
                        </div>
                        <div style={{background:"#fef3c7",borderRadius:12,padding:"12px 14px"}}>
                          <div style={{fontSize:10,color:"#92400e",fontWeight:600,marginBottom:3}}>📦 Lucro bruto</div>
                          <div style={{fontSize:20,fontWeight:800,color:"#92400e"}}>R$ {Number(erpResultado.lucro_bruto).toFixed(2).replace(".",",")}</div>
                        </div>
                      </div>
                      {erpResultado.observacao&&<div style={{background:"#f5f3ff",border:"1px solid #ddd6fe",borderRadius:10,padding:"10px 14px",fontSize:12,color:"#374151",fontStyle:"italic"}}>{erpResultado.observacao}</div>}
                    </div>
                  ) : (
                    <div style={{background:"#fef3c7",border:"1px solid #fde68a",borderRadius:10,padding:"10px 14px",fontSize:12,color:"#92400e"}}>
                      ⚠️ {erpResultado.observacao||"Não foi possível ler os valores na imagem. Tente um print mais nítido."}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Chat com Claude */}
            <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 3px rgba(0,0,0,0.07)",overflow:"hidden",marginBottom:20}}>
              <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{fontSize:13,fontWeight:700,color:"#1e293b"}}>💬 Chat com Claude</div>
                {chatMsgs.length>0&&<button onClick={async()=>{if(window.confirm("Limpar histórico do chat?")){{await limparChatFinanceiro(user.id,orgId);setChatMsgs([]);setChatAnalise("");}}}} style={{background:"#f1f5f9",color:"#64748b",border:"none",borderRadius:7,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>🗑️ Limpar</button>}
              </div>
              <div style={{height:280,overflowY:"auto",padding:"14px 16px",display:"flex",flexDirection:"column",gap:10,background:"#fafafa"}}>
                {chatMsgs.length===0&&(
                  <div style={{textAlign:"center",color:"#94a3b8",fontSize:12,margin:"auto"}}>
                    <div style={{fontSize:32,marginBottom:8}}>🤖</div>
                    <div>Pergunte sobre seus custos, rentabilidade ou qualquer coisa financeira da loja!</div>
                    <div style={{marginTop:8,display:"flex",flexDirection:"column",gap:4}}>
                      {["Qual meu custo operacional este mês?","Quanto preciso vender para cobrir os custos?","Qual fornecedor estou gastando mais?"].map(s=>(
                        <button key={s} onClick={()=>setChatInput(s)}
                          style={{background:"#ede9fe",color:"#6d28d9",border:"none",borderRadius:8,padding:"6px 10px",fontSize:11,cursor:"pointer",textAlign:"left"}}>
                          {s}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {chatMsgs.map((m,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:m.role==="user"?"flex-end":"flex-start"}}>
                    <div style={{maxWidth:"80%",background:m.role==="user"?"#7c3aed":"#fff",color:m.role==="user"?"#fff":"#374151",borderRadius:m.role==="user"?"16px 16px 4px 16px":"16px 16px 16px 4px",padding:"10px 14px",fontSize:12,lineHeight:1.6,boxShadow:"0 1px 3px rgba(0,0,0,0.1)",whiteSpace:"pre-wrap"}}>
                      {m.content}
                    </div>
                  </div>
                ))}
                {chatCarregando&&(
                  <div style={{display:"flex",justifyContent:"flex-start"}}>
                    <div style={{background:"#fff",borderRadius:"16px 16px 16px 4px",padding:"10px 14px",fontSize:12,color:"#94a3b8",boxShadow:"0 1px 3px rgba(0,0,0,0.1)"}}>⏳ Claude está pensando...</div>
                  </div>
                )}
                <div ref={chatEndRef} />
              </div>
              <div style={{padding:"10px 14px",borderTop:"1px solid #f1f5f9",display:"flex",gap:8}}>
                <input value={chatInput} onChange={e=>setChatInput(e.target.value)}
                  onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&enviarChat()}
                  placeholder="Pergunte sobre seus custos..." disabled={chatCarregando}
                  style={{flex:1,padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none"}} />
                <button onClick={enviarChat} disabled={!chatInput.trim()||chatCarregando}
                  style={{padding:"9px 16px",background:chatInput.trim()?"#7c3aed":"#e2e8f0",color:chatInput.trim()?"#fff":"#94a3b8",border:"none",borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer"}}>
                  Enviar
                </button>
              </div>
            </div>

            {/* Modal custo fixo */}
            {showCustoForm&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>
                  <h3 style={{margin:"0 0 16px",fontSize:15,fontWeight:700}}>{editandoCusto?"✏️ Editar custo":"+ Novo custo fixo"}</h3>
                  <div style={{display:"flex",flexDirection:"column",gap:11}}>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Nome *</label>
                      <input value={custoNome} onChange={e=>setCustoNome(e.target.value)} placeholder="Ex: Aluguel sala, Salário Mylena..."
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Categoria</label>
                        <select value={custoCategoria} onChange={e=>setCustoCategoria(e.target.value)}
                          style={{width:"100%",padding:"9px 10px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",background:"#fff",boxSizing:"border-box"}}>
                          {[{id:"aluguel",l:"🏠 Aluguel"},{id:"luz",l:"💡 Luz"},{id:"agua",l:"💧 Água"},{id:"internet",l:"📶 Internet"},{id:"funcionario",l:"👤 Funcionário"},{id:"cartao",l:"💳 Cartão"},{id:"emprestimo",l:"🏦 Empréstimo"},{id:"outros",l:"📋 Outros"}].map(c=><option key={c.id} value={c.id}>{c.l}</option>)}
                        </select>
                      </div>
                      <div>
                        <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Valor mensal (R$) *</label>
                        <input type="number" min="0" step="0.01" value={custoValor} onChange={e=>setCustoValor(e.target.value)} placeholder="0,00"
                          style={{width:"100%",padding:"9px 10px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      </div>
                    </div>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Data de vencimento *</label>
                      <input type="date" value={custoVenc} onChange={e=>setCustoVenc(e.target.value)}
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                      {custoRecorrente&&<span style={{fontSize:10,color:"#94a3b8"}}>Como é recorrente, o vencimento se repete todo mês nesse mesmo dia.</span>}
                    </div>
                    <div>
                      <label style={{fontSize:11,fontWeight:600,color:"#64748b",display:"block",marginBottom:3}}>Observações</label>
                      <input value={custoObs} onChange={e=>setCustoObs(e.target.value)} placeholder="Ex: vence dia 10, parcela 3/12..."
                        style={{width:"100%",padding:"9px 12px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:13,outline:"none",boxSizing:"border-box"}} />
                    </div>
                    <label style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:"#64748b",cursor:"pointer"}}>
                      <input type="checkbox" checked={custoRecorrente} onChange={e=>setCustoRecorrente(e.target.checked)} style={{width:16,height:16}} />
                      Custo recorrente (aparece todo mês)
                    </label>
                  </div>
                  <div style={{display:"flex",gap:10,marginTop:18}}>
                    <button onClick={()=>setShowCustoForm(false)} style={{flex:1,padding:"10px 0",border:"1px solid #e2e8f0",borderRadius:10,background:"#f8fafc",color:"#64748b",fontSize:13,cursor:"pointer"}}>Cancelar</button>
                    <button onClick={salvarCusto} disabled={!custoNome.trim()||!custoValor||!custoVenc||salvandoCusto}
                      style={{flex:2,padding:"10px 0",border:"none",borderRadius:10,background:(custoNome.trim()&&custoValor&&custoVenc)?"#059669":"#e2e8f0",color:(custoNome.trim()&&custoValor&&custoVenc)?"#fff":"#94a3b8",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                      {salvandoCusto?"Salvando...":"Salvar custo"}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })()}

      {showFinanceGate&&<FinanceGate
        pinHash={estoquePlusConfig?.pin_hash}
        onUnlock={()=>{setEstoquePlusDesbloqueado(true);setFinanceUnlocked(true);setActiveTab("financeiro");setShowFinanceGate(false);fetchCustosFixos(user.id,orgId).then(setCustosFixos);fetchMovimentosComPreco(user.id,orgId,finMes).then(setFinMovimentos);}}
        onDefinirPin={()=>{setShowFinanceGate(false);setActiveTab("estoque");setShowDefinirPin(true);}}
      />}
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
