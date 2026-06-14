import { useState } from "react"
import { signIn, signUp, resetPassword } from "./supabase.js"

export default function Auth({ onLogin }) {
  const [mode,     setMode]     = useState("login") // login | register | forgot
  const [email,    setEmail]    = useState("")
  const [password, setPassword] = useState("")
  const [loading,  setLoading]  = useState(false)
  const [error,    setError]    = useState("")
  const [success,  setSuccess]  = useState("")

  const reset = () => { setError(""); setSuccess(""); }

  const handle = async () => {
    reset(); setLoading(true)
    if (!email) { setError("Preencha o e-mail."); setLoading(false); return }

    if (mode === "forgot") {
      const { error } = await resetPassword(email)
      if (error) setError("Não foi possível enviar o e-mail. Verifique o endereço.")
      else setSuccess("E-mail de recuperação enviado! Verifique sua caixa de entrada.")
      setLoading(false); return
    }

    if (!password) { setError("Preencha a senha."); setLoading(false); return }

    if (mode === "register") {
      const { error } = await signUp(email, password)
      if (error) setError(error.message)
      else setSuccess("Conta criada! Verifique seu e-mail para confirmar e depois faça login.")
    } else {
      const { data, error } = await signIn(email, password)
      if (error) setError("E-mail ou senha incorretos.")
      else onLogin(data.session.user)
    }
    setLoading(false)
  }

  const TABS = [
    { id: "login",    label: "Entrar" },
    { id: "register", label: "Criar conta" },
  ]

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#1e3a8a 0%,#1d4ed8 60%,#3b82f6 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:20,fontFamily:"'DM Sans','Segoe UI',sans-serif"}}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
      <div style={{background:"#fff",borderRadius:24,padding:"40px 36px",width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,0.2)"}}>

        {/* Logo */}
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{fontSize:44,marginBottom:8}}>📦</div>
          <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#1e293b"}}>Gerenciador de Pedidos</h1>
          <p style={{margin:"6px 0 0",fontSize:13,color:"#64748b"}}>
            {mode==="forgot" ? "Recuperar senha" : mode==="login" ? "Entre na sua conta" : "Crie sua conta gratuitamente"}
          </p>
        </div>

        {/* Tab bar — only show on login/register */}
        {mode!=="forgot" && (
          <div style={{display:"flex",background:"#f1f5f9",borderRadius:12,padding:4,marginBottom:24}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>{setMode(t.id);reset();}} style={{flex:1,padding:"9px 0",border:"none",borderRadius:9,fontSize:13,fontWeight:600,cursor:"pointer",transition:"all 0.2s",background:mode===t.id?"#fff":"transparent",color:mode===t.id?"#1d4ed8":"#64748b",boxShadow:mode===t.id?"0 1px 4px rgba(0,0,0,0.1)":"none"}}>
                {t.label}
              </button>
            ))}
          </div>
        )}

        {/* Fields */}
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div>
            <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>E-mail</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="seu@email.com"
              onKeyDown={e=>e.key==="Enter"&&handle()}
              style={{width:"100%",padding:"11px 14px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:14,outline:"none",boxSizing:"border-box"}} />
          </div>
          {mode!=="forgot" && (
            <div>
              <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>Senha</label>
              <input type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••"
                onKeyDown={e=>e.key==="Enter"&&handle()}
                style={{width:"100%",padding:"11px 14px",border:"1.5px solid #e2e8f0",borderRadius:10,fontSize:14,outline:"none",boxSizing:"border-box"}} />
              {mode==="register" && <div style={{fontSize:11,color:"#94a3b8",marginTop:4}}>Mínimo 6 caracteres</div>}
            </div>
          )}
        </div>

        {/* Forgot password link */}
        {mode==="login" && (
          <div style={{textAlign:"right",marginTop:8}}>
            <button onClick={()=>{setMode("forgot");reset();}} style={{background:"none",border:"none",fontSize:12,color:"#1d4ed8",cursor:"pointer",fontWeight:500}}>
              Esqueci minha senha
            </button>
          </div>
        )}

        {/* Error / Success */}
        {error  && <div style={{marginTop:14,background:"#fde8e8",borderRadius:8,padding:"10px 14px",fontSize:12,color:"#991b1b",fontWeight:500}}>❌ {error}</div>}
        {success && <div style={{marginTop:14,background:"#d1fae5",borderRadius:8,padding:"10px 14px",fontSize:12,color:"#065f46",fontWeight:500}}>✅ {success}</div>}

        {/* Button */}
        <button onClick={handle} disabled={loading} style={{width:"100%",marginTop:20,padding:"13px 0",background:loading?"#e2e8f0":"linear-gradient(135deg,#1d4ed8,#7c3aed)",color:loading?"#94a3b8":"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:700,cursor:loading?"not-allowed":"pointer"}}>
          {loading ? "Aguarde..." : mode==="login" ? "Entrar →" : mode==="register" ? "Criar conta →" : "Enviar e-mail de recuperação →"}
        </button>

        {/* Back link on forgot mode */}
        {mode==="forgot" && (
          <button onClick={()=>{setMode("login");reset();}} style={{width:"100%",marginTop:12,padding:"10px 0",background:"none",border:"1px solid #e2e8f0",borderRadius:12,fontSize:13,color:"#64748b",cursor:"pointer",fontWeight:500}}>
            ← Voltar para o login
          </button>
        )}
      </div>
    </div>
  )
}
