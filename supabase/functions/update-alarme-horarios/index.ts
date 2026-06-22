// Edge Function: update-alarme-horarios
// Recebe lista de horários (ex: ["09:30","13:00","17:00"]) e
// recria os cron jobs automaticamente no Supabase pg_cron.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

// Converte horário BRT "HH:MM" para cron UTC (BRT = UTC-3)
function horaParaCronUTC(hora: string): string {
  const [h, m] = hora.split(":").map(Number);
  let hUtc = h + 3; // BRT -> UTC
  if (hUtc >= 24) hUtc -= 24;
  return `${m} ${hUtc} * * *`;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const { horarios } = await req.json(); // ex: ["09:30","13:00","17:00"]
    if (!Array.isArray(horarios) || horarios.length === 0) {
      return new Response(JSON.stringify({ error: "Lista de horários inválida." }), {
        status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const authHeader = req.headers.get("Authorization");
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY");
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

    // Valida que é admin
    const callerClient = createClient(supabaseUrl, anonKey, {
      global: { headers: { Authorization: authHeader } },
    });
    const { data: { user }, error: userError } = await callerClient.auth.getUser();
    if (userError || !user) {
      return new Response(JSON.stringify({ error: "Não autenticado." }), {
        status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const adminClient = createClient(supabaseUrl, serviceKey);
    const { data: membro } = await adminClient
      .from("membros").select("organizacao_id, is_admin").eq("user_id", user.id).maybeSingle();
    if (!membro?.is_admin) {
      return new Response(JSON.stringify({ error: "Apenas o admin pode alterar os horários." }), {
        status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const orgId = membro.organizacao_id;
    const funcUrl = `${supabaseUrl}/functions/v1/check-prazos-calendar`;

    // 1. Remove todos os cron jobs antigos desse prefixo
    await adminClient.rpc("exec_sql", {
      sql: `
        DO $$
        DECLARE r record;
        BEGIN
          FOR r IN SELECT jobname FROM cron.job WHERE jobname LIKE 'alarme-${orgId.slice(0,8)}-%'
          LOOP
            PERFORM cron.unschedule(r.jobname);
          END LOOP;
        END $$;
      `
    });

    // 2. Cria um cron job para cada horário
    for (const hora of horarios) {
      const cronExpr = horaParaCronUTC(hora);
      const jobName = `alarme-${orgId.slice(0,8)}-${hora.replace(":","h")}`;
      await adminClient.rpc("exec_sql", {
        sql: `
          SELECT cron.schedule(
            '${jobName}',
            '${cronExpr}',
            $q$SELECT net.http_post(url := '${funcUrl}', headers := jsonb_build_object('Content-Type','application/json'));$q$
          );
        `
      });
    }

    // 3. Salva os horários na tabela
    const { data: existing } = await adminClient
      .from("alarme_horarios").select("id").eq("organizacao_id", orgId).maybeSingle();
    if (existing) {
      await adminClient.from("alarme_horarios")
        .update({ horarios, atualizado_em: new Date().toISOString() })
        .eq("organizacao_id", orgId);
    } else {
      await adminClient.from("alarme_horarios").insert({ organizacao_id: orgId, horarios });
    }

    return new Response(JSON.stringify({ success: true, horarios, jobs: horarios.length }), {
      status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
