import { useState, useEffect } from "react"
import { supabase, signOut } from "./supabase.js"
import Auth from "./Auth.jsx"
import App from "./App.jsx"

export default function AppShell() {
  const [user,    setUser]    = useState(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    // onAuthStateChange fires immediately with current session on mount
    // This is the single source of truth — no need for getSession separately
    const { data: { subscription } } = supabase.auth.onAuthStateChange((event, session) => {
      console.log("Auth event:", event, "user:", session?.user?.email)
      if (session?.user) {
        setUser(session.user)
      } else {
        setUser(null)
      }
      setLoading(false)
    })

    return () => subscription.unsubscribe()
  }, [])

  if (loading) return (
    <div style={{
      minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center",
      background:"linear-gradient(135deg,#1e3a8a,#3b82f6)", fontFamily:"sans-serif"
    }}>
      <div style={{color:"#fff", fontSize:16, fontWeight:600}}>📦 Carregando...</div>
    </div>
  )

  if (!user) return <Auth onLogin={setUser} />
  return <App user={user} onLogout={() => signOut()} />
}
