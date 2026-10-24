import React from "react"
import { supabase, registrarErroApp } from "./supabase.js"

async function logar(erro, contexto) {
  try {
    const { data } = await supabase.auth.getSession()
    const u = data?.session?.user
    if (!u) return
    const { data: m } = await supabase.from("membros").select("organizacao_id, nome").eq("user_id", u.id).maybeSingle()
    await registrarErroApp({ userId: u.id, orgId: m?.organizacao_id, nome: m?.nome }, erro, contexto)
  } catch (e) { /* ignora */ }
}

// Erros fora do React (promessas rejeitadas, scripts) também são registrados, sem derrubar a tela
if (typeof window !== "undefined" && !window.__erroGlobalAtivo) {
  window.__erroGlobalAtivo = true
  let enviados = 0
  window.addEventListener("error", e => { if (enviados++ < 5) logar(e.error || e.message, "window.error") })
  window.addEventListener("unhandledrejection", e => { if (enviados++ < 5) logar(e.reason, "promessa rejeitada") })
}

export default class ErrorBoundary extends React.Component {
  constructor(p) { super(p); this.state = { erro: null } }
  static getDerivedStateFromError(erro) { return { erro } }
  componentDidCatch(erro, info) {
    const msg = String(erro?.message || "")
    // Depois de uma atualização do sistema, a aba aberta pode tentar carregar um arquivo antigo: recarrega sozinho 1 vez
    if (/dynamically imported module|Importing a module script failed|Loading chunk/i.test(msg)) {
      try {
        if (!sessionStorage.getItem("recarregou_chunk")) { sessionStorage.setItem("recarregou_chunk", "1"); location.reload(); return }
      } catch (e) { /* ignora */ }
    }
    logar(erro, (info?.componentStack || "").slice(0, 800))
  }
  render() {
    if (!this.state.erro) return this.props.children
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "linear-gradient(135deg,#1e3a8a,#3b82f6)", fontFamily: "sans-serif", padding: 20 }}>
        <div style={{ background: "#fff", borderRadius: 18, padding: "32px 28px", maxWidth: 420, textAlign: "center", boxShadow: "0 12px 40px rgba(0,0,0,.25)" }}>
          <div style={{ fontSize: 44, marginBottom: 8 }}>😕</div>
          <div style={{ fontSize: 18, fontWeight: 800, color: "#1e293b", marginBottom: 8 }}>Algo deu errado</div>
          <div style={{ fontSize: 13, color: "#64748b", lineHeight: 1.5, marginBottom: 18 }}>
            Seus dados estão salvos e nada foi perdido. O erro foi registrado para ser corrigido. Clique abaixo para recarregar a página.
          </div>
          <button onClick={() => { try { sessionStorage.removeItem("recarregou_chunk") } catch (e) { /* ignora */ } location.reload() }}
            style={{ background: "#2563eb", color: "#fff", border: "none", borderRadius: 12, padding: "11px 26px", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
            🔄 Recarregar
          </button>
          <div style={{ fontSize: 10, color: "#cbd5e1", marginTop: 14, wordBreak: "break-word" }}>{String(this.state.erro?.message || "").slice(0, 140)}</div>
        </div>
      </div>
    )
  }
}
