import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const key = import.meta.env.VITE_SUPABASE_ANON_KEY

if (!url || !key) {
  console.warn('⚠️ Variáveis do Supabase não configuradas. O sistema funcionará sem persistência.')
}

export const supabase = url && key ? createClient(url, key) : null

// ── Carregar todos os pedidos do banco ────────────────────────────────────────
export async function fetchPedidos() {
  if (!supabase) return []
  const { data, error } = await supabase
    .from('pedidos')
    .select('*')
    .order('hora_pedido', { ascending: false })
  if (error) { console.error('Erro ao carregar pedidos:', error); return [] }
  return data.map(dbToRow)
}

// ── Inserir pedidos novos (ignora duplicatas pelo id_plataforma) ──────────────
export async function upsertPedidos(rows) {
  if (!supabase) return { count: 0 }
  const records = rows.map(rowToDb)
  const { data, error } = await supabase
    .from('pedidos')
    .upsert(records, { onConflict: 'id_plataforma', ignoreDuplicates: true })
    .select()
  if (error) { console.error('Erro ao salvar pedidos:', error); return { count: 0 } }
  return { count: data?.length ?? 0 }
}

// ── Remover pedidos por id_plataforma ─────────────────────────────────────────
export async function deletePedidos(ids) {
  if (!supabase || !ids.length) return
  const { error } = await supabase
    .from('pedidos')
    .delete()
    .in('id_plataforma', ids)
  if (error) console.error('Erro ao remover pedidos:', error)
}

// ── Conversões entre formato da app e formato do banco ───────────────────────
function rowToDb(r) {
  return {
    id_plataforma:   r.idPlataforma   || '',
    subpedido:       r.subpedido      || '',
    numero_pedido:   r.numeroPedido   || '',
    plataforma:      r.plataforma     || '',
    loja:            r.loja           || '',
    estado:          r.estado         || '',
    hora_pedido:     r.horaPedido     || null,
    hora_pagamento:  r.horaPagamento  || null,
    prazo_envio:     r.prazoEnvio     || null,
    notas:           r.notas          || '',
    produto:         r.produto        || '',
    variacao:        r.variacao       || '',
    quantidade:      r.quantidade     || '',
    unidade:         r.unidade        || '',
    nome_comprador:  r.nomeComprador  || '',
    id_comprador:    r.idComprador    || '',
    destinatario:    r.destinatario   || '',
  }
}

function dbToRow(d) {
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
    _status: 'existing',
  }
}
