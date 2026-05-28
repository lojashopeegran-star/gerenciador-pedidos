import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const key = import.meta.env.VITE_SUPABASE_ANON_KEY

export const supabase = url && key ? createClient(url, key) : null

// ── Auth ──────────────────────────────────────────────────────────────────────
export async function signUp(email, password) {
  return await supabase.auth.signUp({ email, password })
}
export async function signIn(email, password) {
  return await supabase.auth.signInWithPassword({ email, password })
}
export async function signOut() {
  await supabase.auth.signOut()
}
export async function getSession() {
  const { data } = await supabase.auth.getSession()
  return data.session
}

// ── Pedidos ───────────────────────────────────────────────────────────────────
export async function fetchPedidos(userId) {
  if (!supabase) return []
  const { data, error } = await supabase
    .from('pedidos')
    .select('*')
    .eq('user_id', userId)
    .order('criado_em', { ascending: false })
  if (error) { console.error('Erro ao carregar:', error); return [] }
  return data.map(dbToRow)
}

export async function upsertPedidos(rows, userId) {
  if (!supabase || !rows.length) return { count: 0 }
  const records = rows.map(r => ({ ...rowToDb(r), user_id: userId }))
  const { data, error } = await supabase
    .from('pedidos')
    .upsert(records, { onConflict: 'user_id,id_plataforma', ignoreDuplicates: true })
    .select()
  if (error) { console.error('Erro ao salvar:', error); return { count: 0 } }
  return { count: data?.length ?? 0 }
}

export async function deletePedidos(ids, userId) {
  if (!supabase || !ids.length) return
  const { error } = await supabase
    .from('pedidos')
    .delete()
    .in('id_plataforma', ids)
    .eq('user_id', userId)
  if (error) console.error('Erro ao remover:', error)
}

export async function updatePedidoStatus(idPlataforma, userId, status, nota = '') {
  if (!supabase) return
  const { error } = await supabase
    .from('pedidos')
    .update({ status_interno: status, nota_revisao: nota })
    .eq('id_plataforma', idPlataforma)
    .eq('user_id', userId)
  if (error) console.error('Erro ao atualizar status:', error)
}

// ── Conversões ────────────────────────────────────────────────────────────────
function rowToDb(r) {
  return {
    id_plataforma:  r.idPlataforma  || '',
    subpedido:      r.subpedido     || '',
    numero_pedido:  r.numeroPedido  || '',
    plataforma:     r.plataforma    || '',
    loja:           r.loja          || '',
    estado:         r.estado        || '',
    hora_pedido:    r.horaPedido    || null,
    hora_pagamento: r.horaPagamento || null,
    prazo_envio:    r.prazoEnvio    || null,
    notas:          r.notas         || '',
    produto:        r.produto       || '',
    variacao:       r.variacao      || '',
    quantidade:     r.quantidade    || '',
    unidade:        r.unidade       || '',
    nome_comprador: r.nomeComprador || '',
    id_comprador:   r.idComprador   || '',
    destinatario:   r.destinatario  || '',
    status_interno: r.statusInterno || '',
    nota_revisao:   r.notaRevisao   || '',
  }
}

export function dbToRow(d) {
  return {
    idPlataforma:  d.id_plataforma  || '',
    subpedido:     d.subpedido      || '',
    numeroPedido:  d.numero_pedido  || '',
    plataforma:    d.plataforma     || '',
    loja:          d.loja           || '',
    estado:        d.estado         || '',
    horaPedido:    d.hora_pedido    || '',
    horaPagamento: d.hora_pagamento || '',
    prazoEnvio:    d.prazo_envio    || '',
    notas:         d.notas          || '',
    produto:       d.produto        || '',
    variacao:      d.variacao       || '',
    quantidade:    d.quantidade     || '',
    unidade:       d.unidade        || '',
    nomeComprador: d.nome_comprador || '',
    idComprador:   d.id_comprador   || '',
    destinatario:  d.destinatario   || '',
    statusInterno: d.status_interno || '',
    notaRevisao:   d.nota_revisao   || '',
    _status: 'existing',
  }
}
