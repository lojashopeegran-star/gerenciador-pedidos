import { useState, useRef, useEffect } from "react";
// XLSX carregado sob demanda para não bloquear o primeiro carregamento
let _XLSX = null;
export const getXLSX = async () => { if (!_XLSX) _XLSX = await import("xlsx"); return _XLSX; };
// ── Column maps — matched to REAL Shopee spreadsheet columns ─────────────────
export const PEDIDO_COLS = {
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

export const DEVOLUCAO_COLS = {
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
export function norm(s) {
  return (s||"").toString().trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"");
}

// ── Parse spreadsheet with flexible column matching ───────────────────────────
export function parseSheet(sheet, colMap, loja) {
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
export function normStatus(s) { return norm(s||""); }

// Calcula o mês de referência (formato "2026-06") a partir da Data de criação do pedido
export function mesReferencia(dataCriacao) {
  if (!dataCriacao) return "";
  // Aceita formatos "2026-06-01 00:29" ou "2026-06-01"
  const m = String(dataCriacao).match(/^(\d{4})-(\d{2})/);
  return m ? `${m[1]}-${m[2]}` : "";
}

export function mesLabel(mesRef) {
  if (!mesRef) return "—";
  const [ano, mes] = mesRef.split("-");
  const nomes = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  return `${nomes[parseInt(mes,10)-1]}/${ano}`;
}

// Cores fortes e bem visíveis de destaque para cliente recorrente, por loja.
// lojas[0] = laranja forte, lojas[1] = rosa forte (pink vibrante).
export function corRecorrencia(loja, lojas) {
  if (lojas && loja === lojas[0]) return { bg:"#fb923c", bgLight:"#fed7aa", color:"#7c2d12", border:"#ea580c" };
  if (lojas && loja === lojas[1]) return { bg:"#ec4899", bgLight:"#fbcfe8", color:"#831843", border:"#db2777" };
  return { bg:"#2dd4bf", bgLight:"#99f6e4", color:"#134e4a", border:"#0d9488" };
}

// Simplifica e categoriza o texto do motivo de cancelamento da Shopee
export function formatMotivoCancelamento(motivo) {
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

export function isEnviado(r) {
  const s = normStatus(r.statusPedido);
  return s === "enviado" || s === "entregue" || s === "concluido" ||
         s.startsWith("o comprador pode pedir") || s === "completed" || s === "delivered" ||
         s === "order received";
}
export function isCancelado(r) {
  const s = normStatus(r.statusPedido);
  return s === "cancelado" || s === "nao pago" || s === "unpaid" ||
         s === "cancelamento solicitado" || s.startsWith("cancelad");
}
export function isAberto(r) {
  // Cancelled and sent must never appear in "abertos"
  if (isCancelado(r) || isEnviado(r)) return false;
  const s = normStatus(r.statusPedido);
  return s.includes("a enviar") || s === "a" || s === "pendente";
}
export function isHistorico(r) { return isEnviado(r) || isCancelado(r); }

// ── Deadline color ────────────────────────────────────────────────────────────
export function deadlineInfo(dateStr) {
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

export function fmtBRL(v) { return Number(v||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"}); }
export function currentMes() { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; }

// ── Status badge ──────────────────────────────────────────────────────────────
export function StatusBadge({status}) {
  const s = normStatus(status);
  let bg="#f1f5f9", color="#64748b", label=status||"—";
  if (isAberto({statusPedido:status}))    { bg="#fde8e8"; color="#991b1b"; label="🔴 A ENVIAR"; }
  else if (isEnviado({statusPedido:status}))  { bg="#d1fae5"; color="#065f46"; label="✅ ENVIADO"; }
  else if (isCancelado({statusPedido:status})){ bg="#f3f4f6"; color="#374151"; label="❌ CANCELADO"; }
  return <span style={{background:bg,color,borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{label}</span>;
}

// ── UI Components ─────────────────────────────────────────────────────────────
export function StyledCheckbox({checked,onChange}) {
  return (
    <div onClick={onChange} style={{width:18,height:18,borderRadius:5,border:`2px solid ${checked?"#1d4ed8":"#cbd5e1"}`,background:checked?"#1d4ed8":"#fff",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",transition:"all 0.15s",flexShrink:0}}>
      {checked&&<svg width="10" height="8" viewBox="0 0 10 8" fill="none"><path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>}
    </div>
  );
}

// Botão "Feito" que usa a cor do funcionário que marcou (sistema de times)
export function FeitoButton({isFeito,cor,nome,onClick,small}) {
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
export function RevisaoButton({isRevisao,cor,nome,onClick,small}) {
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
export function VazioButton({isVazio,cor,nome,onClick,small}) {
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

export function ExpandCell({value,maxLen=22}) {
  const [exp,setExp] = useState(false);
  const long = value && value.length > maxLen;
  return (
    <div style={{display:"flex",alignItems:"flex-start",gap:3,maxWidth:180}}>
      <span style={{fontSize:12,color:"#374151",lineHeight:1.4,whiteSpace:exp?"normal":"nowrap",overflow:exp?"visible":"hidden",textOverflow:exp?"clip":"ellipsis",flex:1}}>{value||"—"}</span>
      {long&&<button onClick={()=>setExp(v=>!v)} style={{background:"none",border:"none",cursor:"pointer",color:"#94a3b8",fontSize:12,padding:0,flexShrink:0}}>{exp?"▲":"▼"}</button>}
    </div>
  );
}

export function DropZone({label,sublabel,onFile,file,color,disabled}) {
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

export function Toast({msg,color}) {
  if(!msg)return null;
  return <div style={{position:"fixed",bottom:24,right:24,zIndex:20000,background:color||"#1e293b",color:"#fff",borderRadius:12,padding:"12px 20px",fontSize:13,fontWeight:600,boxShadow:"0 8px 32px rgba(0,0,0,0.2)"}}>{msg}</div>;
}

export function RevisaoModal({order,onConfirm,onClose,onDesmarcar}) {
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

export function FinanceGate({onUnlock, pinHash, onDefinirPin}) {
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

export function TabBar({tabs,active,onChange}) {
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

export const TH={padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:10,color:"#64748b",letterSpacing:0.5,textTransform:"uppercase",whiteSpace:"nowrap",borderBottom:"2px solid #e2e8f0"};
export const TD={padding:"7px 11px",verticalAlign:"middle"};
// ── Helpers de cálculo de produto (usados na projeção do Financeiro) ──
export const _n = (s,d) => (s===""||s==null||isNaN(Number(s))) ? d : Number(s);
// ── Leitura de nota fiscal: utilitários ──
export const _norm = (t) => String(t||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
export const _tokens = (t) => new Set(_norm(t).split(" ").filter(w=>w.length>1));
export const similaridade = (a,b) => {
  const A=_tokens(a), B=_tokens(b); if(!A.size||!B.size) return 0;
  let inter=0; A.forEach(w=>{ if(B.has(w)) inter++; });
  return inter/Math.max(A.size,B.size);
};
export const melhorMatch = (texto, lista, minimo=0.5) => {
  let best=null, bs=0;
  for (const x of lista) { const sc=similaridade(texto,x.nome); if (sc>bs) { bs=sc; best=x; } }
  return bs>=minimo ? best : null;
};
export const mapUnidade = (u) => {
  const x=_norm(u);
  if (["cx","caixa","cxa"].includes(x)) return "cx";
  if (["pct","pc","pac","pacote"].includes(x)) return "pct";
  if (["rl","rolo"].includes(x)) return "rolo";
  if (["m","mt","metro","metros"].includes(x)) return "m";
  if (["kg","quilo"].includes(x)) return "kg";
  if (["l","lt","litro","litros"].includes(x)) return "l";
  return "un";
};
export const arquivoParaBase64 = (file) => new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(String(r.result).split(",")[1]); r.onerror=rej; r.readAsDataURL(file); });
// Reduz fotos grandes (celular) para acelerar o envio; se o navegador não abrir o formato, envia o original
export const prepararImagemNota = async (file) => {
  try {
    const url = URL.createObjectURL(file);
    const img = await new Promise((res,rej)=>{ const i=new Image(); i.onload=()=>res(i); i.onerror=rej; i.src=url; });
    const max = 1600; const k = Math.min(1, max/Math.max(img.width,img.height));
    const c = document.createElement("canvas"); c.width=Math.round(img.width*k); c.height=Math.round(img.height*k);
    c.getContext("2d").drawImage(img,0,0,c.width,c.height);
    URL.revokeObjectURL(url);
    return { mime:"image/jpeg", base64:c.toDataURL("image/jpeg",0.8).split(",")[1] };
  } catch { return { mime:file.type||"image/jpeg", base64:await arquivoParaBase64(file) }; }
};
export const fmtBRLn = (n) => "R$ "+Number(n||0).toFixed(2).replace(".",",");
export const quandoFmt = (iso) => { const d=new Date(iso); return String(d.getDate()).padStart(2,"0")+"/"+String(d.getMonth()+1).padStart(2,"0")+" "+String(d.getHours()).padStart(2,"0")+":"+String(d.getMinutes()).padStart(2,"0"); };
export const liquidoDe = (p,preco) => Number(preco||0)*(1-(_n(p.taxa_plataforma_pct,0)+_n(p.taxa_imposto_pct,0)+_n(p.taxa_adiantamento_pct,0))/100) - _n(p.taxa_fixa,0);
export const custoMpDe = (v) => (v.produto_materias_primas||[]).reduce((a,l)=>{
  const m = l.materias_primas || l;
  const un = Number(m.qtd_por_unidade_compra)>0 ? Number(m.custo_unidade_compra)/Number(m.qtd_por_unidade_compra) : 0;
  return a + un*Number(l.qtd_usada_por_produto||0);
},0);

// ── MAIN APP ──────────────────────────────────────────────────────────────────

// Link do WhatsApp a partir de um telefone/contato (assume Brasil quando vier sem DDI)
export const waUrl = (contato, msg = "") => {
  let d = String(contato || "").replace(/\D/g, "");
  if (d.length < 10) return null;
  if (d.length <= 11) d = "55" + d;
  return `https://wa.me/${d}${msg ? "?text=" + encodeURIComponent(msg) : ""}`;
};
