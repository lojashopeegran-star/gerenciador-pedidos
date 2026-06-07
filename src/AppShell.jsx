import { useState, useEffect } from "react"
import { supabase, signOut } from "./supabase.js"
import Auth from "./Auth.jsx"
import App from "./App.jsx"

export default function AppShell() {
  const [user,    setUser]    = useState(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    // Get initial session immediately
    supabase.auth.getSession().then(({ data: { session } }) => {
      console.log("Initial session:", session?.user?.email || "none")
      setUser(session?.user ?? null)
      setLoading(false)
    })

    // Listen for auth changes
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      console.log("Auth change:", _event, session?.user?.email || "none")
      setUser(session?.user ?? null)
    })

    return () => subscription.unsubscribe()
  }, [])

  if (loading) return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center",
      justifyContent: "center", background: "linear-gradient(135deg,#1e3a8a,#3b82f6)",
      fontFamily: "sans-serif"
    }}>
      <div style={{color:"#fff", fontSize:16, fontWeight:600}}>📦 Carregando...</div>
    </div>
  )

  if (!user) return <Auth onLogin={u => setUser(u)} />
  return <App user={user} onLogout={() => signOut()} />
}
