// Edge Function: google-calendar-connect
// Recebe o "code" de autorização do Google (após o admin autorizar) e troca
// por um refresh_token permanente, salvando no banco para uso futuro.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const { code, redirectUri } = await req.json();
    if (!code) {
      return new Response(JSON.stringify({ error: "code é obrigatório" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const authHeader = req.headers.get("Authorization");
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const googleClientId = Deno.env.get("GOOGLE_CLIENT_ID");
    const googleClientSecret = Deno.env.get("GOOGLE_CLIENT_SECRET");

    const callerClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authHeader } },
    });
    const { data: { user: callerUser }, error: callerError } = await callerClient.auth.getUser();
    if (callerError || !callerUser) {
      return new Response(JSON.stringify({ error: "Não autenticado" }), {
        status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const adminClient = createClient(supabaseUrl, serviceKey);

    // Confirma que o usuário é admin de uma organização
    const { data: membro } = await adminClient
      .from("membros").select("organizacao_id, is_admin").eq("user_id", callerUser.id).maybeSingle();
    if (!membro || !membro.is_admin) {
      return new Response(JSON.stringify({ error: "Apenas o admin pode conectar o Google Calendar" }), {
        status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Troca o code por tokens
    const tokenResp = await fetch("https://oauth2.googleapis.com/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        code,
        client_id: googleClientId,
        client_secret: googleClientSecret,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
      }),
    });
    const tokenData = await tokenResp.json();

    if (!tokenData.refresh_token) {
      return new Response(JSON.stringify({ error: tokenData.error_description || "Não foi possível obter o refresh_token. Tente desconectar e conectar novamente." }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Salva (ou atualiza) o refresh_token
    await adminClient.from("google_calendar_tokens").delete().eq("organizacao_id", membro.organizacao_id);
    const { error: insertError } = await adminClient.from("google_calendar_tokens").insert({
      organizacao_id: membro.organizacao_id,
      user_id: callerUser.id,
      refresh_token: tokenData.refresh_token,
      ativo: true,
    });
    if (insertError) {
      return new Response(JSON.stringify({ error: insertError.message }), {
        status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    return new Response(JSON.stringify({ success: true }), {
      status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
