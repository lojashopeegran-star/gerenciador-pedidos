import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const key = import.meta.env.VITE_SUPABASE_ANON_KEY

export const supabase = url && key ? createClient(url, key) : null

// ── Auth ──────────────────────────────────────────────────────────────────────
export async function signUp(email, password) { return await supabase.auth.signUp({ email, password }) }
export async function signIn(email, password) { return await supabase.auth.signInWithPassword({ email, password }) }
export async function signOut() { await supabase.auth.signOut() }
export async function getSession() { const { data } = await supabase.auth.getSession(); return data.session }

// ── Pedidos ───────────────────────────────────────────────────────────────────
export async function fetchPedidos(userId) {
  if (!supabase) return []
  const { data, error } = await supabase.from('pedidos').select('*').eq('user_id', userId).order('criado_em', { ascending: false })
  if (error) { console.error(error); return [] }
  return data.map(dbToRow)
}

export async function upsertPedidos(rows, userId) {
  if (!supabase || !rows.length) return { count: 0 }
  const records = rows.map(r => ({ ...rowToDb(r), user_id: userId }))
  const { data, error } = await supabase.from('pedidos').upsert(records, { onConflict: 'user_id,id_pedido', ignoreDuplicates: true }).select()
  if (error) { console.error(error); return { count: 0 } }
  return { count: data?.length ?? 0 }
}

export async function updatePedidoEnviado(ids, userId) {
  if (!supabase || !ids.length) return
  const { error } = await supabase.from('pedidos').update({ status_pedido: 'ENVIADO' }).in('id_pedido', ids).eq('user_id', userId)
  if (error) console.error(error)
}

export async function updatePedidoStatus(idPedido, userId, status, nota = '') {
  if (!supabase) return
  const { error } = await supabase.from('pedidos').update({ status_interno: status, nota_revisao: nota }).eq('id_pedido', idPedido).eq('user_id', userId)
  if (error) console.error(error)
}

// ── Faturamento ───────────────────────────────────────────────────────────────
export async function fetchFaturamento(userId) {
  if (!supabase) return []
  const { data, error } = await supabase.from('faturamento').select('*').eq('user_id', userId).order('mes', { ascending: true })
  if (error) { console.error(error); return [] }
  return data
}

export async function upsertFaturamento(userId, mes, loja, valor) {
  if (!supabase) return
  // sum into existing record for that month+loja
  const { data: existing } = await supabase.from('faturamento').select('*').eq('user_id', userId).eq('mes', mes).eq('loja', loja).maybeSingle()
  if (existing) {
    await supabase.from('faturamento').update({ valor: existing.valor + valor }).eq('id', existing.id)
  } else {
    await supabase.from('faturamento').insert({ user_id: userId, mes, loja, valor })
  }
}

// ── Conversões ────────────────────────────────────────────────────────────────
function rowToDb(r) {
  return {
    id_pedido:       r.idPedido      || '',
    status_pedido:   r.statusPedido  || '',
    data_envio:      r.dataEnvio     || null,
    hora_pagamento:  r.horaPagamento || null,
    produto:         r.produto       || '',
    preco:           parseFloat(r.preco) || 0,
    quantidade:      r.quantidade    || '',
    variacao:        r.variacao      || '',
    destinatario:    r.destinatario  || '',
    loja:            r.loja          || '',
    status_interno:  r.statusInterno || '',
    nota_revisao:    r.notaRevisao   || '',
  }
}

export function dbToRow(d) {
  return {
    idPedido:      d.id_pedido      || '',
    statusPedido:  d.status_pedido  || '',
    dataEnvio:     d.data_envio     || '',
    horaPagamento: d.hora_pagamento || '',
    produto:       d.produto        || '',
    preco:         d.preco          || 0,
    quantidade:    d.quantidade     || '',
    variacao:      d.variacao       || '',
    destinatario:  d.destinatario   || '',
    loja:          d.loja           || '',
    statusInterno: d.status_interno || '',
    notaRevisao:   d.nota_revisao   || '',
    _status: 'existing',
  }
}
