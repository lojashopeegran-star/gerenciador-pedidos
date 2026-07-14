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

// Cliente isolado, sem persistência de sessão — usado SOMENTE para criar
// logins de funcionários sem trocar a sessão ativa do admin no navegador.
const supabaseAdminAux = url && key ? createClient(url, key, {
  auth: {
    persistSession: false,
    autoRefreshToken: false,
    detectSessionInUrl: false,
  }
}) : null

export async function signUp(email, password) { return await supabase.auth.signUp({ email, password }) }
export async function signIn(email, password) { return await supabase.auth.signInWithPassword({ email, password }) }
export async function signOut() { await supabase.auth.signOut() }
export async function resetPassword(email) {
  return await supabase.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin,
  })
}
export async function getSession() { const { data } = await supabase.auth.getSession(); return data.session }

// ── Pedidos — paginated to get ALL records ────────────────────────────────────
export async function fetchPedidos(userId, orgId) {
  if (!supabase || !userId) return []
  let allData = []
  let from = 0
  const pageSize = 500
  let page = 0
  while (true) {
    page++
    let query = supabase.from('pedidos').select('*', { count: 'exact' })
    query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
    query = query.is('arquivado_em', null)  // ignora soft-deleted
    const { data, error, count } = await query
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

export async function upsertPedidos(rows, userId, orgId, onProgress) {
  if (!supabase || !rows.length || !userId) return { count: 0 }

  // Deduplica por id_pedido — última ocorrência vence
  var dedupMap = new Map()
  for (var i0 = 0; i0 < rows.length; i0++) {
    var r0 = rows[i0]
    if (r0.idPedido) dedupMap.set(r0.idPedido, r0)
  }
  var uniqueRows = Array.from(dedupMap.values())
  var total = uniqueRows.length
  var batchSize = 50

  // 1. Busca IDs existentes em batch (máx 500 por query)
  var ids = uniqueRows.map(function(r) { return r.idPedido })
  var existingMap = new Map()
  for (var i1 = 0; i1 < ids.length; i1 += 500) {
    var idChunk = ids.slice(i1, i1 + 500)
    var q = supabase.from('pedidos').select('id, id_pedido, status_interno, nota_revisao').in('id_pedido', idChunk)
    q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
    var res = await q
    if (res.data) res.data.forEach(function(row) { existingMap.set(row.id_pedido, row) })
  }

  // 2. Separa em novos vs existentes
  var toInsert = []
  var toUpdate = []
  for (var i2 = 0; i2 < uniqueRows.length; i2++) {
    var rec = Object.assign({}, rowToDb(uniqueRows[i2]), { user_id: userId })
    if (orgId) rec.organizacao_id = orgId
    var ex = existingMap.get(uniqueRows[i2].idPedido)
    if (ex) {
      toUpdate.push({ dbId: ex.id, rec: rec })
    } else {
      toUpdate.push(null) // placeholder
      toInsert.push(Object.assign({}, rec, { status_interno: '', nota_revisao: '' }))
    }
  }
  // Remove placeholders
  toUpdate = toUpdate.filter(function(x) { return x !== null })

  var done = 0
  function report() { done++; if (onProgress) onProgress(Math.round((done / total) * 100)) }

  // 3. Inserts em batch
  for (var i3 = 0; i3 < toInsert.length; i3 += batchSize) {
    var insertChunk = toInsert.slice(i3, i3 + batchSize)
    var ins = await supabase.from('pedidos').insert(insertChunk)
    if (ins.error) console.error('batch insert error:', ins.error.message)
    for (var k = 0; k < insertChunk.length; k++) report()
  }

  // 4. Updates em paralelo por chunk
  for (var i4 = 0; i4 < toUpdate.length; i4 += batchSize) {
    var updateChunk = toUpdate.slice(i4, i4 + batchSize)
    await Promise.all(updateChunk.map(async function(item) {
      var upd = await supabase.from('pedidos').update({
        status_pedido:       item.rec.status_pedido,
        data_envio:          item.rec.data_envio,
        hora_pagamento:      item.rec.hora_pagamento,
        produto:             item.rec.produto,
        preco:               item.rec.preco,
        quantidade:          item.rec.quantidade,
        variacao:            item.rec.variacao,
        destinatario:        item.rec.destinatario,
        notas:               item.rec.notas,
        loja:                item.rec.loja,
        nome_usuario:        item.rec.nome_usuario,
        data_criacao:        item.rec.data_criacao,
        mes_referencia:      item.rec.mes_referencia,
        motivo_cancelamento: item.rec.motivo_cancelamento,
      }).eq('id', item.dbId)
      if (upd.error) console.error('update error:', upd.error.message)
      report()
    }))
  }

  console.log('upsertPedidos done: ' + toInsert.length + ' inserted, ' + toUpdate.length + ' updated')
  return { count: toInsert.length + toUpdate.length }
}

export async function updatePedidoStatus(idPedido, userId, statusInterno, nota = '', orgId, membroNome) {
  if (!supabase || !userId) return
  const updateData = { status_interno: statusInterno, nota_revisao: nota }
  if (statusInterno === 'feito' || statusInterno === 'revisao' || statusInterno === 'vazio') {
    updateData.feito_por_user_id = userId
    updateData.feito_por_nome = membroNome || null
    updateData.feito_em = new Date().toISOString()
  } else {
    // Ao desmarcar, mantém feito_por mas limpa o status
    updateData.feito_por_user_id = null
    updateData.feito_por_nome = null
    updateData.feito_em = null
  }
  let query = supabase.from('pedidos').update(updateData).eq('id_pedido', idPedido)
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  await query
}

// ── Devoluções ────────────────────────────────────────────────────────────────
export async function fetchDevolucoes(userId, orgId) {
  if (!supabase || !userId) return []
  let allData = []
  let from = 0
  while (true) {
    let query = supabase.from('devolucoes').select('*')
    query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
    const { data, error } = await query
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

export async function upsertDevolucoes(rows, userId, orgId) {
  if (!supabase || !rows.length || !userId) return
  for (let i = 0; i < rows.length; i += 50) {
    const batch = rows.slice(i, i + 50)
    const records = batch.map(r => ({ ...r, user_id: userId, ...(orgId ? { organizacao_id: orgId } : {}) }))
    const { error } = await supabase
      .from('devolucoes')
      .upsert(records, { onConflict: 'user_id,id_pedido', ignoreDuplicates: false })
    if (error) console.error('upsertDevolucoes error:', error)
  }
}

// ── Faturamento ───────────────────────────────────────────────────────────────
export async function fetchFaturamento(userId, orgId) {
  if (!supabase || !userId) return []
  let query = supabase.from('faturamento').select('*')
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  const { data, error } = await query.order('mes', { ascending: true })
  if (error) { console.error('fetchFaturamento error:', error); return [] }
  return data || []
}

export async function upsertFaturamento(userId, mes, loja, valor, orgId) {
  if (!supabase || !userId) return
  let existQuery = supabase.from('faturamento').select('*').eq('mes', mes).eq('loja', loja)
  existQuery = orgId ? existQuery.eq('organizacao_id', orgId) : existQuery.eq('user_id', userId)
  const { data: ex } = await existQuery.maybeSingle()
  if (ex) await supabase.from('faturamento').update({ valor: Number(ex.valor) + Number(valor) }).eq('id', ex.id)
  else await supabase.from('faturamento').insert({ user_id: userId, mes, loja, valor, ...(orgId ? { organizacao_id: orgId } : {}) })
}

// ── Delete all ────────────────────────────────────────────────────────────────
export async function deleteAllPedidos(userId, orgId) {
  // Soft delete: marca arquivado_em em vez de deletar permanentemente
  if (!supabase || !userId) return
  const agora = new Date().toISOString()
  let query = supabase.from('pedidos').update({ arquivado_em: agora }).is('arquivado_em', null)
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  await query
}

export async function restaurarPedidos(userId, orgId) {
  if (!supabase || !userId) return
  let query = supabase.from('pedidos').update({ arquivado_em: null }).not('arquivado_em', 'is', null)
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  await query
}

export async function contarPedidosArquivados(userId, orgId) {
  if (!supabase || !userId) return 0
  let query = supabase.from('pedidos').select('id', { count: 'exact', head: true }).not('arquivado_em', 'is', null)
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  const { count } = await query
  return count || 0
}

export async function deleteAllDevolucoes(userId, orgId) {
  if (!supabase || !userId) return
  let query = supabase.from('devolucoes').delete()
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  await query
}

export async function deleteAllFaturamento(userId, orgId) {
  if (!supabase || !userId) return
  let query = supabase.from('faturamento').delete()
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  await query
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
    nome_usuario:   r.nomeUsuario   || '',
    data_criacao:   r.dataCriacao   || null,
    mes_referencia: r.mesReferencia || '',
    motivo_cancelamento: r.motivoCancelamento || '',
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
    feitoPorNome:  d.feito_por_nome || '',
    feitoEm:       d.feito_em       || '',
    nomeUsuario:   d.nome_usuario   || '',
    dataCriacao:   d.data_criacao   || '',
    mesReferencia: d.mes_referencia || '',
    motivoCancelamento: d.motivo_cancelamento || '',
    _status: 'existing',
  }
}

// ── Configurações do usuário (lojas) ─────────────────────────────────────────
export async function fetchConfig(userId, orgId) {
  if (!supabase || !userId) return null
  let query = supabase.from('user_config').select('*')
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  const { data, error } = await query.maybeSingle()
  if (error) { console.error('fetchConfig error:', error); return null }
  return data
}

export async function saveConfig(userId, loja1, loja2, orgId) {
  if (!supabase || !userId) return
  let existQuery = supabase.from('user_config').select('id')
  existQuery = orgId ? existQuery.eq('organizacao_id', orgId) : existQuery.eq('user_id', userId)
  const { data: ex } = await existQuery.maybeSingle()
  if (ex) {
    await supabase.from('user_config')
      .update({ loja1, loja2 })
      .eq('id', ex.id)
  } else {
    await supabase.from('user_config')
      .insert({ user_id: userId, loja1, loja2, ...(orgId ? { organizacao_id: orgId } : {}) })
  }
}

// ── Organizações & Membros (sistema de times) ─────────────────────────────────
export async function fetchMembro(userId) {
  if (!supabase || !userId) return null
  const { data, error } = await supabase
    .from('membros')
    .select('*')
    .eq('user_id', userId)
    .maybeSingle()
  if (error) { console.error('fetchMembro error:', error); return null }
  return data
}

export async function fetchOrganizacao(orgId) {
  if (!supabase || !orgId) return null
  const { data, error } = await supabase
    .from('organizacoes')
    .select('*')
    .eq('id', orgId)
    .maybeSingle()
  if (error) { console.error('fetchOrganizacao error:', error); return null }
  return data
}

export async function fetchMembrosDaOrganizacao(orgId) {
  if (!supabase || !orgId) return []
  const { data, error } = await supabase
    .from('membros')
    .select('*')
    .eq('organizacao_id', orgId)
    .order('criado_em', { ascending: true })
  if (error) { console.error('fetchMembrosDaOrganizacao error:', error); return [] }
  return data || []
}

// Cria um novo funcionário: cria o login (auth) e o registro de membro
// Usa um client isolado (supabaseAdminAux) para NÃO substituir a sessão ativa do admin.
export async function criarFuncionario(orgId, { nome, email, password, cor, permissoes }) {
  if (!supabase || !supabaseAdminAux) return { error: 'Supabase não configurado' }
  const { data: signUpData, error: signUpError } = await supabaseAdminAux.auth.signUp({ email, password })
  if (signUpError) return { error: signUpError.message }
  const newUserId = signUpData?.user?.id
  if (!newUserId) return { error: 'Não foi possível criar o login do funcionário.' }

  // Importante: encerra a sessão do client auxiliar para não deixar resíduo
  await supabaseAdminAux.auth.signOut()

  const { error: membroError } = await supabase.from('membros').insert({
    organizacao_id: orgId,
    user_id: newUserId,
    nome,
    email,
    cor: cor || '#3b82f6',
    is_admin: false,
    pode_ver_financeiro:    permissoes?.pode_ver_financeiro    ?? false,
    pode_zerar_sistema:     permissoes?.pode_zerar_sistema     ?? false,
    pode_carregar_planilha: permissoes?.pode_carregar_planilha ?? true,
    pode_editar_status:     permissoes?.pode_editar_status     ?? true,
  })
  if (membroError) return { error: membroError.message }
  return { success: true, userId: newUserId }
}

export async function atualizarPermissoesMembro(membroId, permissoes) {
  if (!supabase || !membroId) return
  const { error } = await supabase.from('membros').update(permissoes).eq('id', membroId)
  if (error) console.error('atualizarPermissoesMembro error:', error)
  return { error }
}

export async function removerMembro(membroId) {
  if (!supabase || !membroId) return
  const { error } = await supabase.from('membros').delete().eq('id', membroId)
  if (error) console.error('removerMembro error:', error)
  return { error }
}

// Relatório de produtividade: conta quantos pedidos cada membro marcou como feito/revisão
export async function fetchProdutividade(orgId) {
  if (!supabase || !orgId) return []
  const { data, error } = await supabase
    .from('pedidos')
    .select('feito_por_user_id, feito_por_nome, status_interno, feito_em')
    .eq('organizacao_id', orgId)
    .not('feito_por_user_id', 'is', null)
    .not('feito_em', 'is', null)
    .is('arquivado_em', null)
  if (error) { console.error('fetchProdutividade error:', error); return [] }
  return data || []
}

// Apaga o funcionário de vez (incluindo o login/e-mail) via Edge Function,
// que usa a service role key no backend para isso de forma segura.
export async function removerFuncionarioCompleto(userIdToDelete) {
  if (!supabase) return { error: 'Supabase não configurado' }
  const { data, error } = await supabase.functions.invoke('delete-user', {
    body: { userIdToDelete },
  })
  if (error) return { error: error.message || 'Erro ao remover funcionário.' }
  if (data?.error) return { error: data.error }
  return { success: true }
}

// Envia e-mail de redefinição de senha para o funcionário
export async function enviarResetSenhaFuncionario(email) {
  return await resetPassword(email)
}

// ── Auditoria — registra ações importantes da equipe ──────────────────────────
export async function registrarAuditoria(orgId, userId, nomeUsuario, acao, detalhes = '') {
  if (!supabase || !userId) return
  const { error } = await supabase.from('auditoria').insert({
    organizacao_id: orgId,
    user_id: userId,
    nome_usuario: nomeUsuario,
    acao,
    detalhes,
  })
  if (error) console.error('registrarAuditoria error:', error.message)
}

export async function fetchAuditoria(orgIdOrUserId, isOrg) {
  if (!supabase || !orgIdOrUserId) return []
  let query = supabase.from('auditoria').select('*')
  query = isOrg ? query.eq('organizacao_id', orgIdOrUserId) : query.eq('user_id', orgIdOrUserId)
  const { data, error } = await query.order('criado_em', { ascending: false }).limit(200)
  if (error) { console.error('fetchAuditoria error:', error.message); return [] }
  return data || []
}

// ── Google Calendar — conexão e status ─────────────────────────────────────────
export async function fetchGoogleCalendarStatus(orgId) {
  if (!supabase || !orgId) return null
  const { data, error } = await supabase
    .from('google_calendar_tokens')
    .select('id, ativo, criado_em')
    .eq('organizacao_id', orgId)
    .maybeSingle()
  if (error) { console.error('fetchGoogleCalendarStatus error:', error.message); return null }
  return data
}

export async function desconectarGoogleCalendar(orgId) {
  if (!supabase || !orgId) return
  await supabase.from('google_calendar_tokens').delete().eq('organizacao_id', orgId)
}

// ── Alarme horários configuráveis ─────────────────────────────────────────────
export async function fetchAlarmeHorarios(orgId) {
  if (!supabase || !orgId) return ["09:30","11:00","13:00","14:30"]
  const { data, error } = await supabase
    .from('alarme_horarios')
    .select('horarios')
    .eq('organizacao_id', orgId)
    .maybeSingle()
  if (error || !data) return ["09:30","11:00","13:00","14:30"]
  return data.horarios || ["09:30","11:00","13:00","14:30"]
}

export async function salvarAlarmeHorarios(horarios) {
  if (!supabase) return { error: 'Supabase não configurado' }
  const { data, error } = await supabase.functions.invoke('update-alarme-horarios', {
    body: { horarios },
  })
  if (error) return { error: error.message }
  if (data?.error) return { error: data.error }
  return { success: true }
}
