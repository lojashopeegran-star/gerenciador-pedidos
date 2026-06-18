// Edge Function: delete-user
// Apaga de vez um usuário do Supabase Auth (auth.users), liberando o e-mail
// para ser reutilizado depois. Só pode ser chamada pelo admin da organização.
//
// Deploy: supabase functions deploy delete-user
// (veja instruções completas no final do arquivo supabase_teams.sql)

import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";

Deno.serve(async (req) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  };

  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const { userIdToDelete } = await req.json();
    if (!userIdToDelete) {
      return new Response(JSON.stringify({ error: "userIdToDelete é obrigatório" }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Client autenticado com o token de quem está chamando (para validar admin)
    const authHeader = req.headers.get("Authorization");
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

    const callerClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authHeader } },
    });
    const { data: { user: callerUser }, error: callerError } = await callerClient.auth.getUser();
    if (callerError || !callerUser) {
      return new Response(JSON.stringify({ error: "Não autenticado" }), {
        status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Client com privilégio de admin (service role) para checar/excluir
    const adminClient = createClient(supabaseUrl, serviceKey);

    // Confirma que quem está chamando é realmente admin da organização
    // do funcionário que está sendo removido
    const { data: targetMembro } = await adminClient
      .from("membros").select("organizacao_id").eq("user_id", userIdToDelete).maybeSingle();
    if (!targetMembro) {
      return new Response(JSON.stringify({ error: "Funcionário não encontrado" }), {
        status: 404, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const { data: org } = await adminClient
      .from("organizacoes").select("admin_id").eq("id", targetMembro.organizacao_id).maybeSingle();
    if (!org || org.admin_id !== callerUser.id) {
      return new Response(JSON.stringify({ error: "Você não tem permissão para remover este funcionário" }), {
        status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    // Remove o registro de membro primeiro
    await adminClient.from("membros").delete().eq("user_id", userIdToDelete);

    // Apaga o login de vez (auth.users) — libera o e-mail para reuso
    const { error: deleteError } = await adminClient.auth.admin.deleteUser(userIdToDelete);
    if (deleteError) {
      return new Response(JSON.stringify({ error: deleteError.message }), {
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
