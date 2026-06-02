import { useState, useEffect } from "react"
import { supabase, getSession, signOut } from "./supabase.js"
import Auth from "./Auth.jsx"
import App from "./App.jsx"

export default function AppShell() {
  const [user,    setUser]    = useState(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    getSession().then(session => {
      if (session?.user) setUser(session.user)
      setLoading(false)
    })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event, session) => {
      if (event === 'SIGNED_IN' && session?.user) setUser(session.user)
      else if (event === 'SIGNED_OUT') setUser(null)
    })
    return () => subscription.unsubscribe()
  }, [])

  if (loading) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",
      background:"linear-gradient(135deg,#1e3a8a,#3b82f6)",fontFamily:"sans-serif"}}>
      <div style={{color:"#fff",fontSize:16,fontWeight:600}}>📦 Carregando...</div>
    </div>
  )

  if (!user) return <Auth onLogin={setUser} />
  return <App user={user} onLogout={() => signOut()} />
}
