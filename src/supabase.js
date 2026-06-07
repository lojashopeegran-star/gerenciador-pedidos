import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const key = import.meta.env.VITE_SUPABASE_ANON_KEY

// Configure client with persistent session storage
export const supabase = url && key ? createClient(url, key, {
  auth: {
    persistSession: true,
    storageKey: 'gerenciador-pedidos-auth',
    storage: window.localStorage,
    autoRefreshToken: true,
    detectSessionInUrl: false,
  }
}) : null

export async function signUp(email, password) { return await supabase.auth.signUp({ email, password }) }
export async function signIn(email, password) { return await supabase.auth.signInWithPassword({ email, password }) }
export async function signOut() { await supabase.auth.signOut() }
export async function getSession() { const { data } = await supabase.auth.getSession(); return data.session }

// ── Pedidos — paginated to get ALL records ────────────────────────────────────
export async function fetchPedidos(userId) {
  if (!supabase || !userId) return []
  let allData = []
  let from = 0
  const pageSize = 500
  let page = 0
  while (true) {
    page++
    console.log(`fetchPedidos page ${page}: fetching rows ${from} to ${from + pageSize - 1}`)
    const { data, error, count } = await supabase
      .from('pedidos').select('*', { count: 'exact' }).eq('user_id', userId)
      .order('id', { ascending: true })
      .range(from, from + pageSize - 1)
    if (error) { console.error('fetchPedidos error:', error); break }
    console.log(`fetchPedidos page ${page}: got ${data?.length} rows, total in DB: ${count}`)
    if (!data || data.length === 0) break
    allData = allData.concat(data)
    if (data.length < pageSize) break
    from += pageSize
    if (page > 20) { console.error('fetchPedidos: too many pages, stopping'); break }
  }
  console.log('fetchPedidos total loaded:', allData.length)
  return allData.map(dbToRow)
}

export async function upsertPedidos(rows, userId) {
  if (!supabase || !rows.length || !userId) return { count: 0 }

  // Deduplicate by id_pedido keeping LAST occurrence
  // (last row in spreadsheet wins for same ID)
  const dedupMap = new Map()
  for (const r of rows) {
    if (r.idPedido) dedupMap.set(r.idPedido, r)
  }
  const uniqueRows = Array.from(dedupMap.values())
  console.log(`upsertPedidos: ${rows.length} rows -> ${uniqueRows.length} unique`)

  let totalCount = 0
  // Process one at a time to avoid duplicate conflicts within same batch
  for (const r of uniqueRows) {
    const record = { ...rowToDb(r), user_id: userId }
    const { error } = await supabase
      .from('pedidos')
      .upsert(record, { onConflict: 'user_id,id_pedido', ignoreDuplicates: false })
    if (error) {
      console.error('upsertPedidos error for', record.id_pedido, ':', error.message)
    } else {
      totalCount++
    }
  }
  console.log('upsertPedidos total saved:', totalCount)
  return { count: totalCount }
}

export async function updatePedidoStatus(idPedido, userId, statusInterno, nota = '') {
  if (!supabase || !userId) return
  await supabase.from('pedidos')
    .update({ status_interno: statusInterno, nota_revisao: nota })
    .eq('id_pedido', idPedido).eq('user_id', userId)
}

// ── Devoluções ────────────────────────────────────────────────────────────────
export async function fetchDevolucoes(userId) {
  if (!supabase || !userId) return []
  let allData = []
  let from = 0
  while (true) {
    const { data, error } = await supabase
      .from('devolucoes').select('*').eq('user_id', userId)
      .order('criado_em', { ascending: false })
      .range(from, from + 999)
    if (error) { console.error('fetchDevolucoes error:', error); break }
    if (!data || data.length === 0) break
    allData = allData.concat(data)
    if (data.length < 1000) break
    from += 1000
  }
  return allData
}

export async function upsertDevolucoes(rows, userId) {
  if (!supabase || !rows.length || !userId) return
  for (let i = 0; i < rows.length; i += 50) {
    const batch = rows.slice(i, i + 50)
    const records = batch.map(r => ({ ...r, user_id: userId }))
    const { error } = await supabase
      .from('devolucoes')
      .upsert(records, { onConflict: 'user_id,id_pedido', ignoreDuplicates: false })
    if (error) console.error('upsertDevolucoes error:', error)
  }
}

// ── Faturamento ───────────────────────────────────────────────────────────────
export async function fetchFaturamento(userId) {
  if (!supabase || !userId) return []
  const { data, error } = await supabase
    .from('faturamento').select('*').eq('user_id', userId).order('mes', { ascending: true })
  if (error) { console.error('fetchFaturamento error:', error); return [] }
  return data || []
}

export async function upsertFaturamento(userId, mes, loja, valor) {
  if (!supabase || !userId) return
  const { data: ex } = await supabase.from('faturamento')
    .select('*').eq('user_id', userId).eq('mes', mes).eq('loja', loja).maybeSingle()
  if (ex) await supabase.from('faturamento').update({ valor: Number(ex.valor) + Number(valor) }).eq('id', ex.id)
  else await supabase.from('faturamento').insert({ user_id: userId, mes, loja, valor })
}

// ── Delete all ────────────────────────────────────────────────────────────────
export async function deleteAllPedidos(userId) {
  if (!supabase || !userId) return
  await supabase.from('pedidos').delete().eq('user_id', userId)
}

export async function deleteAllDevolucoes(userId) {
  if (!supabase || !userId) return
  await supabase.from('devolucoes').delete().eq('user_id', userId)
}

export async function deleteAllFaturamento(userId) {
  if (!supabase || !userId) return
  await supabase.from('faturamento').delete().eq('user_id', userId)
}

function rowToDb(r) {
  return {
    id_pedido:      r.idPedido      || '',
    status_pedido:  r.statusPedido  || '',
    data_envio:     r.dataEnvio     || null,
    hora_pagamento: r.horaPagamento || null,
    produto:        r.produto       || '',
    preco:          parseFloat(r.preco) || 0,
    quantidade:     r.quantidade    || '',
    variacao:       r.variacao      || '',
    destinatario:   r.destinatario  || '',
    loja:           r.loja          || '',
    status_interno: r.statusInterno || '',
    nota_revisao:   r.notaRevisao   || '',
    notas:          r.notas         || '',
  }
}

export function dbToRow(d) {
  return {
    idPedido:      d.id_pedido      || '',
    statusPedido:  d.status_pedido  || '',
    dataEnvio:     d.data_envio     || '',
    horaPagamento: d.hora_pagamento || '',
    produto:       d.produto        || '',
    preco:         Number(d.preco)  || 0,
    quantidade:    d.quantidade     || '',
    variacao:      d.variacao       || '',
    destinatario:  d.destinatario   || '',
    loja:          d.loja           || '',
    statusInterno: d.status_interno || '',
    notaRevisao:   d.nota_revisao   || '',
    notas:         d.notas          || '',
    _status: 'existing',
  }
}
