import { useState, useEffect, useMemo } from "react"
import {
  registrarMovimentoEstoque, fetchVariacoesUpseller, fetchErrosApp, fetchBackups, gerarBackupAgora, fetchBackupDados,
  fetchResumoSemanal, fetchConsumoRecente, fetchCustosFixos, fetchPagamentosCustos,
} from "./supabase.js"

const fmtN = n => String(Number(Number(n).toFixed(2))).replace(".", ",")
const fmtDH = iso => { try { return new Date(iso).toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo", day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) } catch (e) { return "" } }
const baixar = (nome, conteudo, tipo) => {
  const url = URL.createObjectURL(new Blob([conteudo], { type: tipo }))
  const a = document.createElement("a"); a.href = url; a.download = nome; a.click(); setTimeout(() => URL.revokeObjectURL(url), 2000)
}

const box = { background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 12, padding: "10px 14px", marginBottom: 10 }
const btn = (cor = "#2563eb") => ({ background: cor, color: "#fff", border: "none", borderRadius: 10, padding: "7px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" })
const btnClaro = { background: "#fff", color: "#334155", border: "1px solid #cbd5e1", borderRadius: 10, padding: "6px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }

// ─── Pendências ──────────────────────────────────────────────────────────────
function Pendencias({ orgId, estoqueProdutos, materiasPrimas, produtosCalc, onIrPara }) {
  const [vars, setVars] = useState(null)
  useEffect(() => { fetchVariacoesUpseller(orgId).then(r => setVars(Array.isArray(r) ? r : [])) }, [orgId])
  const ativos = estoqueProdutos.filter(p => p.ativo !== false)
  const grupos = useMemo(() => {
    const mpsAtivas = materiasPrimas.filter(m => m.ativo !== false)
    const prodSemMat = []
    produtosCalc.forEach(p => (p.produto_variacoes || []).forEach(v => {
      if (!(v.produto_materias_primas || []).length) prodSemMat.push(`${p.nome} › ${v.nome || "variação"}`)
    }))
    return [
      { t: "Estoque negativo", icon: "🔴", tab: "estoque", itens: ativos.filter(p => Number(p.quantidade_atual) < 0).map(p => `${p.nome}: ${fmtN(p.quantidade_atual)} ${p.unidade}`), dica: "Provável saída lançada a mais. Corrija em Estoque (use ↩ Desfazer)." },
      { t: "Materiais sem preço", icon: "💲", tab: "estoque", itens: ativos.filter(p => !Number(p.preco_unitario)).map(p => p.nome), dica: "Sem preço o custo dos produtos e o gasto do mês ficam errados. Abra o material e preencha em Info financeiras." },
      { t: "Materiais sem quantidade mínima", icon: "📉", tab: "estoque", itens: ativos.filter(p => !Number(p.quantidade_minima)).map(p => p.nome), dica: "Sem mínimo o sistema não avisa quando está acabando." },
      { t: "Matérias-primas sem item de estoque", icon: "🧱", tab: "produtos", itens: mpsAtivas.filter(m => !m.estoque_produto_id).map(m => m.nome), dica: "Sem vínculo com o estoque a baixa automática não desconta esse material. Produtos › Matérias-primas." },
      { t: "Matérias-primas sem custo", icon: "🧮", tab: "produtos", itens: mpsAtivas.filter(m => !Number(m.custo_unidade_compra)).map(m => m.nome), dica: "Custo zero distorce o lucro do produto." },
      { t: "Variações de produto sem materiais", icon: "📦", tab: "produtos", itens: prodSemMat, dica: "Sem materiais, o custo dessa variação é zero e a baixa não consome nada." },
    ]
  }, [estoqueProdutos, materiasPrimas, produtosCalc])
  const semVinculo = (vars || []).filter(v => !v.variacao_id && !v.ignorar)
  const total = grupos.reduce((s, g) => s + g.itens.length, 0) + semVinculo.length

  const Grupo = ({ icon, t, itens, dica, tab, extra }) => itens.length === 0 ? null : (
    <div style={box}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 13, color: "#1e293b" }}>{icon} {t} <span style={{ background: "#fee2e2", color: "#b91c1c", borderRadius: 10, padding: "1px 8px", fontSize: 11 }}>{itens.length}</span></div>
        <button style={btnClaro} onClick={() => onIrPara(tab)}>Ir para a tela →</button>
      </div>
      <div style={{ fontSize: 11, color: "#64748b", margin: "4px 0 6px" }}>{dica}</div>
      <div style={{ fontSize: 12, color: "#334155", maxHeight: 110, overflowY: "auto", lineHeight: 1.6 }}>
        {itens.slice(0, 40).map((x, i) => <div key={i}>• {x}</div>)}
        {itens.length > 40 && <div style={{ color: "#94a3b8" }}>… e mais {itens.length - 40}</div>}
      </div>
      {extra}
    </div>
  )
  return (
    <div>
      <div style={{ fontSize: 13, color: total ? "#b45309" : "#15803d", fontWeight: 700, marginBottom: 10 }}>
        {vars === null ? "Verificando…" : total ? `${total} pendência${total > 1 ? "s" : ""} para resolver` : "✅ Tudo certo, nenhuma pendência!"}
      </div>
      {grupos.map(g => <Grupo key={g.t} {...g} />)}
      <Grupo icon="🧾" t="Partes de pedidos UpSeller sem vínculo" tab="upseller"
        dica="Pedidos com essas partes não dão baixa de estoque. Vincule na aba Pedidos UpSeller (os mais frequentes primeiro)."
        itens={semVinculo.slice(0, 200).map(v => `${v.variacao} — ${v.pedidos} pedido${v.pedidos > 1 ? "s" : ""}`)} />
    </div>
  )
}

// ─── Contagem física ─────────────────────────────────────────────────────────
function Contagem({ user, orgId, nome, estoqueProdutos, onRecarregar, showToast }) {
  const [busca, setBusca] = useState("")
  const [cont, setCont] = useState({})
  const [salvando, setSalvando] = useState(false)
  const lista = estoqueProdutos.filter(p => p.ativo !== false && (!busca || p.nome.toLowerCase().includes(busca.toLowerCase())))
  const alterados = estoqueProdutos.filter(p => cont[p.id] !== undefined && cont[p.id] !== "" && Number(String(cont[p.id]).replace(",", ".")) !== Number(p.quantidade_atual))
  const dif = p => Number(String(cont[p.id]).replace(",", ".")) - Number(p.quantidade_atual)

  const aplicar = async () => {
    if (!alterados.length) return
    if (!window.confirm(`Ajustar o estoque de ${alterados.length} item(ns) para a quantidade contada?\n\nCada ajuste fica registrado nas movimentações e pode ser desfeito com ↩ Desfazer. Ajustes de contagem NÃO entram como compra nem como consumo nos relatórios.`)) return
    setSalvando(true)
    let ok = 0; const erros = []
    for (const p of alterados) {
      const d = dif(p)
      const r = await registrarMovimentoEstoque(user.id, orgId, nome, {
        produto_id: p.id, tipo: d > 0 ? "entrada" : "saida", quantidade: Math.abs(d),
        motivo: `Ajuste de contagem (sistema ${fmtN(p.quantidade_atual)} → contado ${fmtN(Number(String(cont[p.id]).replace(",", ".")))})`,
        preco_unitario: null, fornecedor_id: null,
      })
      if (r.error) erros.push(`${p.nome}: ${r.error}`); else ok++
    }
    setSalvando(false)
    await onRecarregar()
    setCont({})
    showToast(erros.length ? `⚠️ ${ok} ajustado(s). Falhou: ${erros.join("; ")}` : `✅ ${ok} item(ns) ajustado(s) pela contagem.`, erros.length ? "#b45309" : "#059669", 6000)
  }
  return (
    <div>
      <div style={{ fontSize: 12, color: "#64748b", marginBottom: 8 }}>Digite o que você contou na prateleira. Só os itens preenchidos serão ajustados.</div>
      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <input value={busca} onChange={e => setBusca(e.target.value)} placeholder="Buscar material…" style={{ flex: 1, border: "1px solid #cbd5e1", borderRadius: 10, padding: "7px 10px", fontSize: 13 }} />
        <button disabled={!alterados.length || salvando} onClick={aplicar} style={{ ...btn("#059669"), opacity: !alterados.length || salvando ? .5 : 1 }}>
          {salvando ? "Salvando…" : `Aplicar ${alterados.length || ""} ajuste${alterados.length === 1 ? "" : "s"}`}
        </button>
      </div>
      <div style={{ maxHeight: 380, overflowY: "auto", border: "1px solid #e2e8f0", borderRadius: 12 }}>
        {lista.map(p => {
          const tem = cont[p.id] !== undefined && cont[p.id] !== ""
          const d = tem ? dif(p) : 0
          return (
            <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", borderBottom: "1px solid #f1f5f9" }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: "#1e293b", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.nome}</div>
                <div style={{ fontSize: 11, color: "#94a3b8" }}>Sistema: {fmtN(p.quantidade_atual)} {p.unidade}</div>
              </div>
              <input inputMode="decimal" value={cont[p.id] ?? ""} placeholder="contado" onChange={e => setCont(c => ({ ...c, [p.id]: e.target.value }))}
                style={{ width: 90, border: "1px solid #cbd5e1", borderRadius: 8, padding: "6px 8px", fontSize: 13, textAlign: "right" }} />
              <div style={{ width: 80, fontSize: 12, fontWeight: 700, textAlign: "right", color: !tem || d === 0 ? "#94a3b8" : d > 0 ? "#15803d" : "#b91c1c" }}>
                {!tem ? "" : d === 0 ? "✓ igual" : (d > 0 ? "+" : "") + fmtN(d)}
              </div>
            </div>
          )
        })}
        {!lista.length && <div style={{ padding: 20, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>Nenhum material encontrado.</div>}
      </div>
    </div>
  )
}

// ─── Resumo semanal ──────────────────────────────────────────────────────────
export async function montarResumoSemanal(user, orgId, estoqueProdutos) {
  const desde = new Date(Date.now() - 7 * 86400000).toISOString()
  const hoje = new Date()
  const mes = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, "0")}`
  const [sem, cons, custos, pagos] = await Promise.all([
    fetchResumoSemanal(user.id, orgId, desde), fetchConsumoRecente(user.id, orgId, 30),
    fetchCustosFixos(user.id, orgId), fetchPagamentosCustos(user.id, orgId, mes),
  ])
  const h0 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate())
  const contas = []
  custos.forEach(c => {
    if (!c.data_vencimento || pagos.indexOf(c.id) >= 0) return
    const pp = String(c.data_vencimento).slice(0, 10).split("-").map(Number)
    const alvo = c.recorrente !== false ? new Date(hoje.getFullYear(), hoje.getMonth(), Math.min(pp[2], new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getDate())) : new Date(pp[0], pp[1] - 1, pp[2])
    const dias = Math.round((alvo - h0) / 86400000)
    if (dias <= 7) contas.push({ nome: c.nome, valor: Number(c.valor) || 0, dias, data: alvo })
  })
  contas.sort((a, b) => a.dias - b.dias)
  const repor = []
  estoqueProdutos.filter(p => p.ativo !== false).forEach(p => {
    const c = Number(cons[p.id]) || 0, q = Number(p.quantidade_atual) || 0
    const dias = c > 0 ? q / (c / 30) : null
    const baixo = Number(p.quantidade_minima) > 0 && q <= Number(p.quantidade_minima)
    if (!baixo && !(dias !== null && dias <= 7)) return
    repor.push({ nome: p.nome, unidade: p.unidade, q, dias, sug: c > 0 ? Math.max(0, Math.ceil((c - q) * 10) / 10) : 0 })
  })
  return { ...sem, contas, repor }
}

function textoResumo(r) {
  const L = ["📋 RESUMO DA SEMANA", `Pedidos: ${r.pedidos} (cancelados: ${r.cancelados})`, "", "Material usado:"]
  r.uso.slice(0, 12).forEach(u => L.push(`• ${u.nome}: ${fmtN(u.qtd)} ${u.unidade}`))
  if (!r.uso.length) L.push("• (nenhuma baixa registrada)")
  L.push("", "Contas (vencidas ou nos próximos 7 dias):")
  r.contas.forEach(c => L.push(`• ${c.nome} — R$ ${c.valor.toFixed(2).replace(".", ",")} — ${c.dias < 0 ? "vencida" : c.dias === 0 ? "hoje" : "em " + c.dias + " dia(s)"}`))
  if (!r.contas.length) L.push("• nenhuma")
  L.push("", "Repor:")
  r.repor.forEach(x => L.push(`• ${x.nome}: restam ${fmtN(x.q)} ${x.unidade}${x.dias !== null ? ` (~${Math.max(0, Math.floor(x.dias))} dias)` : ""}${x.sug > 0 ? ` · comprar ~${fmtN(x.sug)}` : ""}`))
  if (!r.repor.length) L.push("• nada urgente")
  return L.join("\n")
}

function ResumoSemanal({ user, orgId, estoqueProdutos, showToast }) {
  const [r, setR] = useState(null)
  useEffect(() => { montarResumoSemanal(user, orgId, estoqueProdutos).then(setR).catch(() => setR({ pedidos: 0, cancelados: 0, uso: [], contas: [], repor: [] })) }, [])  // eslint-disable-line react-hooks/exhaustive-deps
  if (!r) return <div style={{ color: "#94a3b8", fontSize: 13 }}>Montando o resumo…</div>
  const Sec = ({ t, children }) => <div style={box}><div style={{ fontWeight: 700, fontSize: 13, color: "#1e293b", marginBottom: 4 }}>{t}</div><div style={{ fontSize: 12, color: "#334155", lineHeight: 1.6 }}>{children}</div></div>
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 12, color: "#64748b" }}>Últimos 7 dias</div>
        <button style={btnClaro} onClick={() => { navigator.clipboard?.writeText(textoResumo(r)).then(() => showToast("📋 Resumo copiado! Cole no WhatsApp ou e-mail.", "#059669")).catch(() => showToast("Não consegui copiar.", "#991b1b")) }}>📋 Copiar resumo</button>
      </div>
      <Sec t="🧾 Pedidos">{r.pedidos} pedidos · {r.cancelados} cancelados</Sec>
      <Sec t="🧱 Material usado">{r.uso.length ? r.uso.slice(0, 12).map(u => <div key={u.nome}>• {u.nome}: {fmtN(u.qtd)} {u.unidade}</div>) : "Nenhuma baixa registrada."}</Sec>
      <Sec t="💸 Contas (vencidas ou nos próximos 7 dias)">{r.contas.length ? r.contas.map(c => <div key={c.nome}>• {c.nome} — R$ {c.valor.toFixed(2).replace(".", ",")} — {c.dias < 0 ? "vencida" : c.dias === 0 ? "hoje" : `em ${c.dias} dia(s)`}</div>) : "Nenhuma."}</Sec>
      <Sec t="📦 O que repor">{r.repor.length ? r.repor.map(x => <div key={x.nome}>• {x.nome}: restam {fmtN(x.q)} {x.unidade}{x.dias !== null ? ` (~${Math.max(0, Math.floor(x.dias))} dias)` : ""}{x.sug > 0 ? ` · comprar ~${fmtN(x.sug)}` : ""}</div>) : "Nada urgente."}</Sec>
    </div>
  )
}

// ─── Backup ──────────────────────────────────────────────────────────────────
function Backup({ orgId, showToast }) {
  const [lista, setLista] = useState(null)
  const [busy, setBusy] = useState(false)
  const carregar = () => fetchBackups(orgId).then(setLista)
  useEffect(() => { carregar() }, [])  // eslint-disable-line react-hooks/exhaustive-deps
  const gerar = async () => {
    setBusy(true); const r = await gerarBackupAgora(orgId); setBusy(false)
    if (r.error) { showToast("❌ " + r.error, "#991b1b", 6000); return }
    showToast("✅ Backup criado.", "#059669"); carregar()
  }
  const baixarJson = async b => {
    const d = await fetchBackupDados(b.id)
    if (d.error) { showToast("❌ " + d.error, "#991b1b"); return }
    baixar(`backup-${String(b.criado_em).slice(0, 10)}.json`, JSON.stringify(d.dados, null, 2), "application/json")
  }
  const baixarXlsx = async b => {
    const d = await fetchBackupDados(b.id)
    if (d.error) { showToast("❌ " + d.error, "#991b1b"); return }
    const X = await import("xlsx"); const wb = X.utils.book_new()
    Object.entries(d.dados).forEach(([nome, linhas]) => {
      const plano = (linhas || []).map(l => Object.fromEntries(Object.entries(l).map(([k, v]) => [k, v !== null && typeof v === "object" ? JSON.stringify(v) : v])))
      X.utils.book_append_sheet(wb, X.utils.json_to_sheet(plano.length ? plano : [{}]), nome.slice(0, 31))
    })
    X.writeFile(wb, `backup-${String(b.criado_em).slice(0, 10)}.xlsx`)
  }
  const semTabela = lista && lista.error
  return (
    <div>
      <div style={{ fontSize: 12, color: "#64748b", marginBottom: 10, lineHeight: 1.5 }}>
        Todo domingo à noite o sistema guarda uma cópia de estoque, movimentações, produtos, matérias-primas, vínculos, custos e projeções (guarda as 12 mais recentes). Baixe um arquivo de vez em quando para ter uma cópia fora do sistema.
      </div>
      <button onClick={gerar} disabled={busy} style={{ ...btn(), marginBottom: 12, opacity: busy ? .6 : 1 }}>{busy ? "Gerando…" : "💾 Criar backup agora"}</button>
      {semTabela && <div style={{ ...box, background: "#fffbeb", borderColor: "#fcd34d", fontSize: 12 }}>O backup ainda não foi ativado no banco de dados (falta rodar o SQL "robustez"). {lista.error}</div>}
      {Array.isArray(lista) && !lista.length && <div style={{ fontSize: 12, color: "#94a3b8" }}>Nenhum backup ainda.</div>}
      {Array.isArray(lista) && lista.map(b => (
        <div key={b.id} style={{ ...box, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#1e293b" }}>{fmtDH(b.criado_em)} <span style={{ fontSize: 10, color: "#94a3b8", fontWeight: 600 }}>({b.origem})</span></div>
            <div style={{ fontSize: 11, color: "#64748b" }}>{b.resumo ? `${b.resumo.estoque_produtos || 0} materiais · ${b.resumo.estoque_movimentos || 0} movimentos · ${b.resumo.produtos || 0} produtos · ${b.resumo.variacao_vinculos || 0} vínculos` : ""}</div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <button style={btnClaro} onClick={() => baixarXlsx(b)}>⬇ Planilha</button>
            <button style={btnClaro} onClick={() => baixarJson(b)}>⬇ Arquivo (JSON)</button>
          </div>
        </div>
      ))}
    </div>
  )
}

// ─── Erros registrados ───────────────────────────────────────────────────────
function Erros({ orgId }) {
  const [l, setL] = useState(null)
  useEffect(() => { fetchErrosApp(orgId).then(setL) }, [orgId])
  if (!l) return <div style={{ color: "#94a3b8", fontSize: 13 }}>Carregando…</div>
  if (l.error) return <div style={{ ...box, background: "#fffbeb", fontSize: 12 }}>O registro de erros ainda não foi ativado no banco (falta rodar o SQL "robustez"). {l.error}</div>
  if (!l.length) return <div style={{ color: "#15803d", fontSize: 13, fontWeight: 700 }}>✅ Nenhum erro registrado.</div>
  return <div>{l.map(e => (
    <div key={e.id} style={box}>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#b91c1c" }}>{e.mensagem}</div>
      <div style={{ fontSize: 11, color: "#94a3b8" }}>{fmtDH(e.criado_em)} · {e.usuario_nome || ""}</div>
      {e.contexto && <div style={{ fontSize: 10, color: "#64748b", whiteSpace: "pre-wrap", maxHeight: 60, overflow: "hidden" }}>{e.contexto}</div>}
    </div>))}</div>
}

// ─── Modal ───────────────────────────────────────────────────────────────────
const ABAS = [["pend", "✅ Pendências"], ["cont", "🔢 Contagem"], ["sem", "📋 Resumo semanal"], ["bkp", "💾 Backup"], ["err", "🐞 Erros"]]

export default function Ferramentas(props) {
  const [aba, setAba] = useState(props.abaInicial || "pend")
  return (
    <div onClick={props.onClose} style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,.55)", zIndex: 9500, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: 18, width: "min(720px,100%)", maxHeight: "90vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,.3)" }}>
        <div style={{ padding: "14px 18px 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ fontSize: 16, fontWeight: 800, color: "#1e293b" }}>🧰 Ferramentas</div>
          <button onClick={props.onClose} style={{ background: "#f1f5f9", border: "none", borderRadius: 8, width: 28, height: 28, cursor: "pointer", fontSize: 15 }}>✕</button>
        </div>
        <div style={{ display: "flex", gap: 6, padding: "10px 18px", flexWrap: "wrap", borderBottom: "1px solid #f1f5f9" }}>
          {ABAS.map(([id, t]) => (
            <button key={id} onClick={() => setAba(id)} style={{ border: "none", borderRadius: 20, padding: "6px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer", background: aba === id ? "#2563eb" : "#f1f5f9", color: aba === id ? "#fff" : "#475569" }}>{t}</button>
          ))}
        </div>
        <div style={{ padding: 18, overflowY: "auto" }}>
          {aba === "pend" && <Pendencias {...props} onIrPara={t => { props.onIrPara(t); props.onClose() }} />}
          {aba === "cont" && <Contagem {...props} />}
          {aba === "sem" && <ResumoSemanal {...props} />}
          {aba === "bkp" && <Backup {...props} />}
          {aba === "err" && <Erros {...props} />}
        </div>
      </div>
    </div>
  )
}
