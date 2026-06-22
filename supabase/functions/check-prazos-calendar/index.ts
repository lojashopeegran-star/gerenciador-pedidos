// Edge Function: check-prazos-calendar
// Roda automaticamente (via Cron Job) nos horários: 09:30, 11:00, 13:00, 14:30
// Verifica pedidos vencendo hoje e cria um evento no Google Calendar do admin,
// com lembrete, para cada organização que tiver o Google Calendar conectado.

import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";

async function refreshAccessToken(refreshToken: string, clientId: string, clientSecret: string) {
  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "refresh_token",
    }),
  });
  const data = await resp.json();
  return data.access_token;
}

async function criarEventoCalendar(accessToken: string, calendarId: string, titulo: string, descricao: string) {
  const agora = new Date();
  const inicio = agora.toISOString();
  const fim = new Date(agora.getTime() + 15 * 60 * 1000).toISOString(); // evento de 15 min

  const resp = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events`, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      summary: titulo,
      description: descricao,
      start: { dateTime: inicio },
      end: { dateTime: fim },
      reminders: {
        useDefault: false,
        overrides: [{ method: "popup", minutes: 0 }],
      },
    }),
  });
  return await resp.json();
}

Deno.serve(async (req) => {
  try {
    const supabaseUrl = Deno.env.get("SUPABASE_URL");
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
    const googleClientId = Deno.env.get("GOOGLE_CLIENT_ID");
    const googleClientSecret = Deno.env.get("GOOGLE_CLIENT_SECRET");

    const adminClient = createClient(supabaseUrl, serviceKey);

    // Busca todas as organizações com Google Calendar conectado e ativo
    const { data: conexoes, error: conexoesError } = await adminClient
      .from("google_calendar_tokens")
      .select("*")
      .eq("ativo", true);

    if (conexoesError) throw conexoesError;
    if (!conexoes || conexoes.length === 0) {
      return new Response(JSON.stringify({ message: "Nenhuma organização com Google Calendar conectado." }), {
        status: 200, headers: { "Content-Type": "application/json" },
      });
    }

    const hoje = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
    const resultados = [];

    for (const conexao of conexoes) {
      // Busca pedidos "em aberto" (não feito) com prazo vencendo hoje ou já vencido
      const { data: pedidos, error: pedidosError } = await adminClient
        .from("pedidos")
        .select("id_pedido, destinatario, data_envio, status_interno, status_pedido")
        .eq("organizacao_id", conexao.organizacao_id)
        .neq("status_interno", "feito")
        .lte("data_envio", hoje + " 23:59:59");

      if (pedidosError) { resultados.push({ org: conexao.organizacao_id, error: pedidosError.message }); continue; }

      // Filtra: só pedidos realmente "em aberto" (não enviado/cancelado)
      const abertosVencendo = (pedidos || []).filter(p => {
        const s = (p.status_pedido || "").toLowerCase();
        const isEnviadoOuCancelado = /enviado|entregue|concluido|cancelado|nao pago|order received/.test(s);
        return !isEnviadoOuCancelado;
      });

      if (abertosVencendo.length === 0) {
        resultados.push({ org: conexao.organizacao_id, status: "sem pedidos vencendo" });
        continue;
      }

      try {
        const accessToken = await refreshAccessToken(conexao.refresh_token, googleClientId, googleClientSecret);
        if (!accessToken) throw new Error("Não foi possível renovar o acesso ao Google Calendar.");

        const pendentes = abertosVencendo.filter(p => p.status_interno !== "revisao");
        const emRevisao = abertosVencendo.filter(p => p.status_interno === "revisao");

        const hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit", timeZone: "America/Sao_Paulo" });
        const titulo = `🚨 ${abertosVencendo.length} pedido(s) vencendo - verificação ${hora}`;

        const formatarLista = (lista) => lista.slice(0, 15).map(p => `• ${p.id_pedido} — ${p.destinatario || "sem nome"}`).join("\n")
          + (lista.length > 15 ? `\n... e mais ${lista.length - 15}` : "");

        let descricao = `Pedidos com prazo vencendo hoje ou já vencidos:\n\n`;
        if (pendentes.length > 0) {
          descricao += `⏳ PENDENTES (${pendentes.length}):\n${formatarLista(pendentes)}\n\n`;
        }
        if (emRevisao.length > 0) {
          descricao += `📋 EM REVISÃO (${emRevisao.length}):\n${formatarLista(emRevisao)}`;
        }

        const evento = await criarEventoCalendar(accessToken, conexao.calendar_id || "primary", titulo, descricao);
        resultados.push({ org: conexao.organizacao_id, status: "evento criado", total: abertosVencendo.length, eventoId: evento.id });
      } catch (err) {
        resultados.push({ org: conexao.organizacao_id, error: err.message });
      }
    }

    return new Response(JSON.stringify({ resultados }), {
      status: 200, headers: { "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500, headers: { "Content-Type": "application/json" },
    });
  }
});
