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
  const pageSize = 1000
  const base = () => {
    let q = supabase.from('pedidos').select('*')
    q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
    return q.is('arquivado_em', null) // ignora soft-deleted
  }
  const pagina = async (n) => {
    const { data, error } = await base().order('id', { ascending: true }).range(n * pageSize, n * pageSize + pageSize - 1)
    if (error) { console.error('fetchPedidos error:', error); return null }
    return data || []
  }
  // 1ª página já mostra quantas existem; as demais vêm em paralelo (antes era uma por vez + contagem exata em cada)
  let cq = supabase.from('pedidos').select('id', { count: 'estimated', head: true })
  cq = orgId ? cq.eq('organizacao_id', orgId) : cq.eq('user_id', userId)
  cq = cq.is('arquivado_em', null)
  const [primeira, cnt] = await Promise.all([pagina(0), cq])
  if (!primeira) return []
  let all = primeira
  if (primeira.length === pageSize) {
    const estimado = Number(cnt?.count) || 0
    let n = 1
    // estimativa só orienta quantas páginas pedir; continua até vir página incompleta
    const lote = Math.max(4, Math.ceil(estimado / pageSize) + 2 - 1)
    while (true) {
      const idx = Array.from({ length: Math.min(lote, 40) }, (_, i) => n + i)
      const res = await Promise.all(idx.map(pagina))
      let fim = false
      for (const r of res) {
        if (!r) { fim = true; break }
        all = all.concat(r)
        if (r.length < pageSize) { fim = true; break }
      }
      if (fim || n > 200) break
      n += idx.length
    }
  }
  return all.map(dbToRow)
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
  const trintaDiasAtras = new Date()
  trintaDiasAtras.setDate(trintaDiasAtras.getDate() - 30)
  // Pagina até buscar todos os registros dos últimos 30 dias
  let allData = []
  let from = 0
  const pageSize = 1000
  while (true) {
    const { data, error } = await supabase
      .from('pedidos')
      .select('feito_por_user_id, feito_por_nome, status_interno, feito_em')
      .eq('organizacao_id', orgId)
      .not('feito_por_user_id', 'is', null)
      .not('feito_em', 'is', null)
      .gte('feito_em', trintaDiasAtras.toISOString())
      .range(from, from + pageSize - 1)
    if (error) { console.error('fetchProdutividade error:', error); break }
    if (!data || data.length === 0) break
    allData = allData.concat(data)
    if (data.length < pageSize) break
    from += pageSize
  }
  return allData
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

// ── Estoque ────────────────────────────────────────────────────────────────────
export async function fetchEstoqueProdutos(userId, orgId) {
  if (!supabase || !userId) return []
  let query = supabase.from('estoque_produtos').select('*, fornecedores(id, nome, contato)').eq('ativo', true)
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  const { data, error } = await query.order('nome', { ascending: true })
  if (error) { console.error('fetchEstoqueProdutos error:', error.message); return [] }
  return data || []
}

export async function salvarEstoqueProduto(userId, orgId, produto) {
  if (!supabase || !userId) return { error: 'Não autenticado' }
  const rec = {
    nome: produto.nome,
    unidade: produto.unidade || 'un',
    quantidade_atual: Number(produto.quantidade_atual) || 0,
    quantidade_minima: Number(produto.quantidade_minima) || 0,
    ativo: true,
    user_id: userId,
    ...(orgId ? { organizacao_id: orgId } : {}),
  }
  if (produto.id) {
    const { error } = await supabase.from('estoque_produtos').update({ ...rec, atualizado_em: new Date().toISOString() }).eq('id', produto.id)
    if (error) return { error: error.message }
    return { success: true }
  } else {
    const { data, error } = await supabase.from('estoque_produtos').insert(rec).select()
    if (error) return { error: error.message }
    return { success: true, data }
  }
}

export async function excluirEstoqueProduto(produtoId) {
  if (!supabase) return { error: 'Supabase não configurado' }
  const { error } = await supabase.from('estoque_produtos').update({ ativo: false }).eq('id', produtoId)
  if (error) return { error: error.message }
  return { success: true }
}

export async function registrarMovimentoEstoque(userId, orgId, membroNome, movimento) {
  if (!supabase || !userId) return { error: 'Não autenticado' }
  const quantidade = Math.abs(Number(movimento.quantidade))
  if (!quantidade) return { error: 'Quantidade inválida' }

  // Busca saldo atual
  const { data: prod, error: prodErr } = await supabase
    .from('estoque_produtos').select('quantidade_atual, preco_unitario, fornecedor_id').eq('id', movimento.produto_id).maybeSingle()
  if (prodErr || !prod) return { error: 'Produto não encontrado' }

  const delta = movimento.tipo === 'entrada' ? quantidade : -quantidade
  const novoSaldo = Number(prod.quantidade_atual) + delta
  if (novoSaldo < 0) return { error: 'Estoque insuficiente para saída' }

  // Atualiza saldo
  await supabase.from('estoque_produtos')
    .update({ quantidade_atual: novoSaldo, atualizado_em: new Date().toISOString() })
    .eq('id', movimento.produto_id)

  // Entradas sem preço informado usam o preço/fornecedor cadastrados no Estoque+,
  // para que a compra apareça no gasto de material do Financeiro.
  const ehEntrada = movimento.tipo === 'entrada'
  const preco = movimento.preco_unitario !== undefined ? movimento.preco_unitario
    : (ehEntrada && prod.preco_unitario ? Number(prod.preco_unitario) : null)
  const fornecedorId = movimento.fornecedor_id !== undefined ? movimento.fornecedor_id
    : (ehEntrada ? (prod.fornecedor_id || null) : null)

  // Registra movimento
  const row = {
    produto_id: movimento.produto_id,
    organizacao_id: orgId || null,
    user_id: userId,
    tipo: movimento.tipo,
    quantidade,
    motivo: movimento.motivo || '',
    referencia_pedido: movimento.referencia_pedido || null,
    registrado_por: membroNome || '',
    preco_unitario: preco,
    fornecedor_id: fornecedorId,
    nota_fiscal: movimento.nota_fiscal || null,
  }
  if (movimento.criado_em) row.criado_em = movimento.criado_em
  const { data: mov, error } = await supabase.from('estoque_movimentos').insert(row).select().single()
  if (error) {
    // desfaz o ajuste de saldo se o movimento não pôde ser gravado
    await supabase.from('estoque_produtos').update({ quantidade_atual: Number(prod.quantidade_atual) }).eq('id', movimento.produto_id)
    return { error: error.message }
  }
  return { success: true, novoSaldo, movimento: mov }
}

export async function fetchMovimentosEstoque(userId, orgId, produtoId = null, limite = 50) {
  if (!supabase || !userId) return []
  let query = supabase.from('estoque_movimentos')
    .select('*, estoque_produtos(nome, unidade)')
    .order('criado_em', { ascending: false })
    .limit(limite)
  if (orgId) query = query.eq('organizacao_id', orgId)
  else query = query.eq('user_id', userId)
  if (produtoId) query = query.eq('produto_id', produtoId)
  const { data, error } = await query
  if (error) { console.error('fetchMovimentosEstoque error:', error.message); return [] }
  return data || []
}

// ── PRODUTOS / CALCULADORA ──────────────────────────────────────────────────

// ── Biblioteca de matérias-primas (compartilhada entre produtos) ─────────────
export async function fetchMateriasPrimas(userId, orgId) {
  let q = supabase.from('materias_primas').select('*').eq('ativo', true)
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data, error } = await q.order('nome', { ascending: true })
  if (error) { console.error('fetchMateriasPrimas error:', error.message); return [] }
  return data || []
}

export async function salvarMateriaPrimaLib(userId, orgId, mp) {
  const payload = {
    nome: mp.nome,
    custo_unidade_compra: mp.custo_unidade_compra || 0,
    qtd_por_unidade_compra: mp.qtd_por_unidade_compra || 1,
    unidade: mp.unidade || 'un',
    estoque_produto_id: mp.estoque_produto_id || null,
    rendimento_estoque: mp.rendimento_estoque || null,
    perda_pct: Number(mp.perda_pct) || 0,
  }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  if (mp.id) {
    const { error } = await supabase.from('materias_primas').update(payload).eq('id', mp.id)
    return error ? { error: error.message } : { success: true, id: mp.id }
  }
  const { data, error } = await supabase.from('materias_primas').insert(payload).select('id').single()
  return error ? { error: error.message } : { success: true, id: data.id }
}

// Exclusão lógica: produtos que já usam a matéria-prima continuam calculando o custo
export async function atualizarPerdaMp(id, pct) {
  const { error } = await supabase.from('materias_primas').update({ perda_pct: Number(pct) || 0 }).eq('id', id)
  return error ? { error: error.message } : { success: true }
}

export async function excluirMateriaPrimaLib(id) {
  const { error } = await supabase.from('materias_primas').update({ ativo: false }).eq('id', id)
  return error ? { error: error.message } : { success: true }
}

// Salva produto + variações + vínculos de matérias-primas de cada variação
// variacoes: [{ id?, nome, preco_venda, mps: [{ materia_prima_id, nome, qtd_usada }] }]
export async function salvarProdutoCompleto(userId, orgId, produto, variacoes) {
  const base = await salvarProduto(userId, orgId, {
    ...produto,
    preco_venda: variacoes[0]?.preco_venda || 0,
  })
  if (base.error) return base
  const produtoId = base.id
  const owner = orgId ? { organizacao_id: orgId } : { user_id: userId }

  const { data: existentes } = await supabase.from('produto_variacoes').select('id').eq('produto_id', produtoId)
  const idsMantidos = new Set(variacoes.filter(v => v.id).map(v => v.id))
  const remover = (existentes || []).map(v => v.id).filter(id => !idsMantidos.has(id))
  if (remover.length) {
    const { error } = await supabase.from('produto_variacoes').delete().in('id', remover)
    if (error) return { error: error.message }
  }

  for (let i = 0; i < variacoes.length; i++) {
    const v = variacoes[i]
    const payload = { produto_id: produtoId, nome: v.nome || 'Padrão', preco_venda: v.preco_venda || 0, ordem: i, ...owner }
    let variacaoId = v.id
    if (variacaoId) {
      const { error } = await supabase.from('produto_variacoes').update(payload).eq('id', variacaoId)
      if (error) return { error: error.message }
    } else {
      const { data, error } = await supabase.from('produto_variacoes').insert(payload).select('id').single()
      if (error) return { error: error.message }
      variacaoId = data.id
    }
    const { error: eDel } = await supabase.from('produto_materias_primas').delete().eq('variacao_id', variacaoId)
    if (eDel) return { error: eDel.message }
    const linhas = (v.mps || []).filter(m => m.materia_prima_id && Number(m.qtd_usada) > 0).map((m, idx) => ({
      produto_id: produtoId,
      variacao_id: variacaoId,
      materia_prima_id: m.materia_prima_id,
      nome: m.nome || '',
      qtd_usada_por_produto: Number(m.qtd_usada),
      ordem: idx,
      ...owner,
    }))
    if (linhas.length) {
      const { error } = await supabase.from('produto_materias_primas').insert(linhas)
      if (error) return { error: error.message }
    }
  }
  return { success: true, id: produtoId }
}

export async function fetchProdutos(userId, orgId) {
  let query = supabase.from('produtos')
    .select('*, produto_variacoes(*, produto_materias_primas(*, materias_primas(*)))')
    .eq('ativo', true)
  query = orgId ? query.eq('organizacao_id', orgId) : query.eq('user_id', userId)
  const { data, error } = await query.order('nome', { ascending: true })
  if (error) { console.error('fetchProdutos error:', error.message); return [] }
  return (data || []).map(p => ({
    ...p,
    produto_variacoes: (p.produto_variacoes || []).slice().sort((a, b) => (a.ordem || 0) - (b.ordem || 0)),
  }))
}

export async function salvarProduto(userId, orgId, produto) {
  const payload = {
    nome: produto.nome,
    descricao: produto.descricao || null,
    preco_venda: produto.preco_venda || 0,
    taxa_plataforma_pct: produto.taxa_plataforma_pct ?? 20,
    taxa_imposto_pct: produto.taxa_imposto_pct ?? 4,
    taxa_adiantamento_pct: produto.taxa_adiantamento_pct ?? 3.5,
    taxa_fixa: produto.taxa_fixa ?? 4,
    ativo: true,
  }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  if (produto.id) {
    const { error } = await supabase.from('produtos').update(payload).eq('id', produto.id)
    return error ? { error: error.message } : { success: true, id: produto.id }
  } else {
    const { data, error } = await supabase.from('produtos').insert(payload).select('id').single()
    return error ? { error: error.message } : { success: true, id: data.id }
  }
}

export async function excluirProduto(produtoId) {
  const { error } = await supabase.from('produtos').update({ ativo: false }).eq('id', produtoId)
  return error ? { error: error.message } : { success: true }
}

export async function salvarMateriaPrima(userId, orgId, materiaPrima) {
  const payload = {
    produto_id: materiaPrima.produto_id,
    nome: materiaPrima.nome,
    custo_unidade_compra: materiaPrima.custo_unidade_compra || 0,
    qtd_por_unidade_compra: materiaPrima.qtd_por_unidade_compra || 1,
    qtd_usada_por_produto: materiaPrima.qtd_usada_por_produto || 1,
    unidade: materiaPrima.unidade || 'un',
    ordem: materiaPrima.ordem || 0,
  }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  if (materiaPrima.id) {
    const { error } = await supabase.from('produto_materias_primas').update(payload).eq('id', materiaPrima.id)
    return error ? { error: error.message } : { success: true }
  } else {
    const { error } = await supabase.from('produto_materias_primas').insert(payload)
    return error ? { error: error.message } : { success: true }
  }
}

export async function excluirMateriaPrima(id) {
  const { error } = await supabase.from('produto_materias_primas').delete().eq('id', id)
  return error ? { error: error.message } : { success: true }
}

// ─── Fornecedores ────────────────────────────────────────────────────────────

export async function fetchFornecedores(userId, orgId) {
  let q = supabase.from('fornecedores').select('*').eq('ativo', true)
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data } = await q.order('nome', { ascending: true })
  return data || []
}

export async function salvarFornecedor(userId, orgId, f) {
  const payload = {
    nome: f.nome?.trim(),
    contato: f.contato?.trim() || null,
    site_link: f.site_link?.trim() || null,
    observacoes: f.observacoes?.trim() || null,
  }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  if (f.id) {
    const { error } = await supabase.from('fornecedores').update(payload).eq('id', f.id)
    return error ? { error: error.message } : { success: true }
  } else {
    const { data, error } = await supabase.from('fornecedores').insert(payload).select().single()
    return error ? { error: error.message } : { success: true, data }
  }
}

export async function excluirFornecedor(id) {
  const { error } = await supabase.from('fornecedores').update({ ativo: false }).eq('id', id)
  return error ? { error: error.message } : { success: true }
}

// ─── Estoque+ config (PIN) ────────────────────────────────────────────────────

export async function fetchEstoquePlusConfig(userId, orgId) {
  let q = supabase.from('estoque_plus_config').select('*')
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data } = await q.maybeSingle()
  return data || null
}

export async function salvarEstoquePlusPin(userId, orgId, pinHash) {
  // Primeiro verifica se já existe
  const atual = await fetchEstoquePlusConfig(userId, orgId)
  const payload = { pin_hash: pinHash, atualizado_em: new Date().toISOString() }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  if (atual) {
    const { error } = await supabase.from('estoque_plus_config').update(payload)
      .eq(orgId ? 'organizacao_id' : 'user_id', orgId || userId)
    return error ? { error: error.message } : { success: true }
  } else {
    const { error } = await supabase.from('estoque_plus_config').insert(payload)
    return error ? { error: error.message } : { success: true }
  }
}

// ─── Estoque Produtos com campos extras ──────────────────────────────────────

export async function atualizarEstoqueProdutoPlus(produtoId, campos) {
  // campos: { preco_unitario, fornecedor_id, link_produto, observacoes, categoria }
  const { error } = await supabase.from('estoque_produtos').update(campos).eq('id', produtoId)
  return error ? { error: error.message } : { success: true }
}

// ─── Movimentos com preço ─────────────────────────────────────────────────────

export async function fetchMovimentosComPreco(userId, orgId, mes) {
  // mes = "2026-10" (YYYY-MM)
  const inicio = mes + '-01'
  const fim = new Date(new Date(inicio).getFullYear(), new Date(inicio).getMonth()+1, 1).toISOString().slice(0,10)
  let q = supabase.from('estoque_movimentos')
    .select('*, estoque_produtos(nome, unidade), fornecedores(nome)')
    .gte('criado_em', inicio).lt('criado_em', fim)
    .eq('tipo', 'entrada')
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data } = await q.order('criado_em', { ascending: false })
  return (data || []).filter(m => !String(m.motivo || '').startsWith('Estorno') && !String(m.motivo || '').startsWith('Ajuste de contagem'))
}

// Saídas de estoque do mês (material realmente consumido) — mes = "YYYY-MM"
export async function fetchSaidasMes(userId, orgId, mes) {
  const [y, m] = mes.split('-').map(Number)
  const inicio = new Date(y, m - 1, 1).toISOString()
  const fim = new Date(y, m, 1).toISOString()
  // saídas + estornos (pedido cancelado devolve material: entra como saída negativa)
  let q = supabase.from('estoque_movimentos')
    .select('*, estoque_produtos(nome, unidade, preco_unitario)')
    .gte('criado_em', inicio).lt('criado_em', fim)
    .or('tipo.eq.saida,and(tipo.eq.entrada,motivo.like.Estorno%)')
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data, error } = await q.order('criado_em', { ascending: false })
  if (error) { console.error('fetchSaidasMes error:', error.message); return [] }
  // Material usado = só o que veio dos pedidos do UpSeller vinculados (baixa automática/retroativa e seus estornos).
  // Saídas manuais do estoque não entram.
  const dePedido = m => /^(Baixa autom|Baixa retroativa|Estorno)/.test(String(m.motivo || ''))
  return (data || []).filter(dePedido).map(m => m.tipo === 'entrada' ? { ...m, quantidade: -Number(m.quantidade) } : m)
}

// ─── Histórico de operações e desfazer ───────────────────────────────────────

export const pick = (o, cols) => Object.fromEntries(cols.filter(c => o && c in o).map(c => [c, o[c]]))
export const COLS_ESTOQUE_PROD = ['nome', 'unidade', 'quantidade_atual', 'quantidade_minima', 'preco_unitario', 'fornecedor_id', 'link_produto', 'alternativas', 'rendimento', 'unidade_uso', 'observacoes', 'categoria']
export const COLS_FORNECEDOR = ['nome', 'contato', 'site_link', 'observacoes']
export const COLS_CUSTO = ['nome', 'categoria', 'valor', 'recorrente', 'observacoes', 'data_vencimento']
export const COLS_MP = ['nome', 'custo_unidade_compra', 'qtd_por_unidade_compra', 'unidade', 'estoque_produto_id', 'rendimento_estoque', 'perda_pct']
export const COLS_MOV = ['id', 'produto_id', 'tipo', 'quantidade', 'motivo', 'preco_unitario', 'fornecedor_id', 'nota_fiscal']

// Foto completa de um produto (taxas + variações + matérias-primas) para poder restaurar
export function snapProduto(p) {
  return {
    produto: pick(p, ['nome', 'descricao', 'taxa_plataforma_pct', 'taxa_imposto_pct', 'taxa_adiantamento_pct', 'taxa_fixa']),
    variacoes: (p.produto_variacoes || []).map(v => ({
      id: v.id, nome: v.nome, preco_venda: Number(v.preco_venda) || 0,
      mps: (v.produto_materias_primas || []).map(l => ({
        materia_prima_id: l.materia_prima_id, nome: l.nome || '', qtd_usada: Number(l.qtd_usada_por_produto) || 0,
      })),
    })),
  }
}

// Registra uma operação no histórico. Nunca lança erro (não pode atrapalhar a operação em si).
export async function registrarOperacao(orgId, userId, nomeUsuario, op) {
  try {
    if (!supabase || !userId) return null
    const id = crypto.randomUUID()
    const { error } = await supabase.from('operacoes_log').insert({
      id,
      organizacao_id: orgId || null,
      user_id: userId,
      usuario_nome: nomeUsuario || null,
      tipo: op.tipo || 'geral',
      acao: op.acao,
      descricao: op.descricao,
      tabela: op.tabela || null,
      registro_id: op.registro_id ? String(op.registro_id) : null,
      antes: op.antes ?? null,
      depois: op.depois ?? null,
      lote_id: op.lote_id || null,
    })
    if (error) { console.error('registrarOperacao error:', error.message); return null }
    return id
  } catch (e) { console.error('registrarOperacao', e); return null }
}

export async function fetchOperacoes(userId, orgId, limite = 300) {
  let q = supabase.from('operacoes_log').select('*')
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data, error } = await q.order('seq', { ascending: false }).limit(limite)
  if (error) { console.error('fetchOperacoes error:', error.message); return [] }
  return data || []
}

async function _setAtivo(tabela, id, ativo) {
  const { error } = await supabase.from(tabela).update({ ativo }).eq('id', id)
  return error ? { error: error.message } : { success: true }
}

async function _reverter(op, ctx) {
  const { tabela, acao, registro_id: id, antes, depois } = op
  const { userId, orgId } = ctx

  if (tabela === 'nota_fiscal' || tabela === 'baixa_pedidos') return { success: true } // linha-resumo do lote; as ações filhas é que revertem

  if (tabela === 'estoque_movimentos') {
    const m = depois
    const { data: prod } = await supabase.from('estoque_produtos').select('quantidade_atual, nome').eq('id', m.produto_id).maybeSingle()
    if (!prod) return { error: 'O material desse movimento não existe mais.' }
    const delta = m.tipo === 'entrada' ? -Number(m.quantidade) : Number(m.quantidade)
    const novo = Number(prod.quantidade_atual) + delta
    if (novo < 0) return { error: `Não dá para desfazer: o saldo de "${prod.nome}" ficaria negativo (parte desse material já foi usada).` }
    const { error: e1 } = await supabase.from('estoque_produtos').update({ quantidade_atual: novo, atualizado_em: new Date().toISOString() }).eq('id', m.produto_id)
    if (e1) return { error: e1.message }
    const { error: e2 } = await supabase.from('estoque_movimentos').delete().eq('id', m.id)
    if (e2) {
      await supabase.from('estoque_produtos').update({ quantidade_atual: Number(prod.quantidade_atual) }).eq('id', m.produto_id)
      return { error: e2.message }
    }
    return { success: true }
  }

  if (tabela === 'estoque_produtos') {
    if (acao === 'criar') return _setAtivo(tabela, id, false)
    if (acao === 'excluir') return _setAtivo(tabela, id, true)
    if (acao === 'editar') {
      const patch = { ...antes }
      if (depois && 'quantidade_atual' in depois && 'quantidade_atual' in antes && Number(depois.quantidade_atual) !== Number(antes.quantidade_atual)) {
        // a edição mexeu no saldo: desfaz só a diferença, preservando movimentos posteriores
        const { data: prod } = await supabase.from('estoque_produtos').select('quantidade_atual').eq('id', id).maybeSingle()
        patch.quantidade_atual = Number(prod?.quantidade_atual || 0) - (Number(depois.quantidade_atual) - Number(antes.quantidade_atual))
      } else {
        delete patch.quantidade_atual
      }
      const { error } = await supabase.from(tabela).update({ ...patch, atualizado_em: new Date().toISOString() }).eq('id', id)
      return error ? { error: error.message } : { success: true }
    }
  }

  if (tabela === 'fornecedores' || tabela === 'custos_fixos' || tabela === 'materias_primas') {
    if (acao === 'criar') return _setAtivo(tabela, id, false)
    if (acao === 'excluir') return _setAtivo(tabela, id, true)
    if (acao === 'editar') {
      const { error } = await supabase.from(tabela).update(antes).eq('id', id)
      return error ? { error: error.message } : { success: true }
    }
  }

  if (tabela === 'custos_fixos_pagamentos') {
    const r = await marcarCustoPago(userId, orgId, depois.custo_id, depois.mes, acao !== 'pagar')
    return r
  }

  if (tabela === 'produtos') {
    if (acao === 'criar') return _setAtivo(tabela, id, false)
    if (acao === 'excluir') return _setAtivo(tabela, id, true)
    if (acao === 'editar') {
      // variações que foram removidas depois precisam ser recriadas (sem o id antigo)
      const { data: atuais } = await supabase.from('produto_variacoes').select('id').eq('produto_id', id)
      const ids = new Set((atuais || []).map(v => v.id))
      const variacoes = (antes.variacoes || []).map(v => ({ ...v, id: ids.has(v.id) ? v.id : undefined }))
      const r = await salvarProdutoCompleto(userId, orgId, { id, ...antes.produto }, variacoes)
      return r.error ? { error: r.error } : { success: true }
    }
  }

  return { error: 'Esse tipo de operação não pode ser desfeito automaticamente.' }
}

// Desfaz uma lista de operações (do mais recente para o mais antigo). Para na primeira falha.
export async function desfazerOperacoes(ops, ctx, nomeUsuario) {
  const ordenadas = [...ops].filter(o => !o.desfeito_em).sort((a, b) => b.seq - a.seq)
  let feitas = 0
  for (const op of ordenadas) {
    let r
    try { r = await _reverter(op, ctx) } catch (e) { r = { error: e.message } }
    if (r.error) return { error: r.error, feitas }
    await supabase.from('operacoes_log').update({ desfeito_em: new Date().toISOString(), desfeito_por: nomeUsuario || null }).eq('id', op.id)
    feitas++
  }
  return { success: true, feitas }
}

// Leitura de nota fiscal (foto ou PDF) pelo Gemini
// Print de tabela de variações: 1ª coluna = nome, última coluna = preço
export async function lerVariacoesPrint(mime_type, data) {
  const { data: { session } } = await supabase.auth.getSession()
  const token = session?.access_token
  if (!token) return { error: 'Não autenticado' }
  const res = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/ler-nota-fiscal`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ mime_type, data, modo: 'variacoes' }),
  })
  let json = null
  try { json = await res.json() } catch { /* ignore */ }
  if (!res.ok || json?.error) return { error: json?.error || `Erro ${res.status} ao ler o print` }
  return { variacoes: json.dados?.variacoes || [] }
}

export async function lerNotaFiscal(mime_type, data) {
  const { data: { session } } = await supabase.auth.getSession()
  const token = session?.access_token
  if (!token) return { error: 'Não autenticado' }
  const res = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/ler-nota-fiscal`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ mime_type, data }),
  })
  let json = null
  try { json = await res.json() } catch { /* ignore */ }
  if (!res.ok || json?.error) return { error: json?.error || `Erro ${res.status} ao ler a nota` }
  return { dados: json.dados }
}

// Já existe movimento com esse número de nota? (evita lançar a mesma nota duas vezes)
export async function notaJaLancada(userId, orgId, numero) {
  if (!numero) return false
  let q = supabase.from('estoque_movimentos').select('id', { count: 'exact', head: true }).eq('nota_fiscal', numero)
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { count } = await q
  return (count || 0) > 0
}

// ─── Custos Fixos ─────────────────────────────────────────────────────────────

export async function fetchCustosFixos(userId, orgId) {
  let q = supabase.from('custos_fixos').select('*').eq('ativo', true)
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data } = await q.order('categoria').order('nome')
  return data || []
}

export async function salvarCustoFixo(userId, orgId, custo) {
  const payload = {
    nome:        custo.nome?.trim(),
    categoria:   custo.categoria || 'outros',
    valor:       Number(custo.valor) || 0,
    recorrente:  custo.recorrente !== false,
    observacoes: custo.observacoes?.trim() || null,
    data_vencimento: custo.data_vencimento || null,
  }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  if (custo.id) {
    const { error } = await supabase.from('custos_fixos').update({ ...payload, atualizado_em: new Date().toISOString() }).eq('id', custo.id)
    return error ? { error: error.message } : { success: true }
  } else {
    const { data, error } = await supabase.from('custos_fixos').insert(payload).select('id').single()
    return error ? { error: error.message } : { success: true, id: data.id }
  }
}

// IDs dos custos fixos já pagos no mês (formato YYYY-MM)
export async function fetchPagamentosCustos(userId, orgId, mes) {
  let q = supabase.from('custos_fixos_pagamentos').select('custo_id').eq('mes', mes)
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data, error } = await q
  if (error) { console.error('fetchPagamentosCustos error:', error.message); return [] }
  return (data || []).map(r => r.custo_id)
}

export async function marcarCustoPago(userId, orgId, custoId, mes, pago) {
  if (pago) {
    const rec = { custo_id: custoId, mes }
    if (orgId) rec.organizacao_id = orgId; else rec.user_id = userId
    const { error } = await supabase.from('custos_fixos_pagamentos').upsert(rec, { onConflict: 'custo_id,mes' })
    return error ? { error: error.message } : { success: true }
  }
  const { error } = await supabase.from('custos_fixos_pagamentos').delete().eq('custo_id', custoId).eq('mes', mes)
  return error ? { error: error.message } : { success: true }
}

export async function excluirCustoFixo(id) {
  const { error } = await supabase.from('custos_fixos').update({ ativo: false }).eq('id', id)
  return error ? { error: error.message } : { success: true }
}

// ─── Chat financeiro com Claude (via Edge Function) ───────────────────────────

export async function chamarClaudeFinanceiro(mensagens, contexto) {
  const { data: { session } } = await supabase.auth.getSession()
  const token = session?.access_token
  if (!token) return { error: 'Não autenticado' }

  const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
  const res = await fetch(`${supabaseUrl}/functions/v1/claude-financeiro`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ messages: mensagens, contexto }),
  })
  if (!res.ok) {
    const err = await res.text()
    return { error: err }
  }
  const data = await res.json()
  return data.error ? { error: data.error } : { reply: data.reply }
}

export async function fetchChatFinanceiro(userId, orgId, limit = 40) {
  let q = supabase.from('financeiro_chat').select('*')
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data } = await q.order('criado_em', { ascending: true }).limit(limit)
  return data || []
}

export async function salvarMensagemChat(userId, orgId, role, content) {
  const payload = { role, content }
  if (orgId) payload.organizacao_id = orgId; else payload.user_id = userId
  const { error } = await supabase.from('financeiro_chat').insert(payload)
  return error ? { error: error.message } : { success: true }
}

export async function limparChatFinanceiro(userId, orgId) {
  let q = supabase.from('financeiro_chat').delete()
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { error } = await q
  return error ? { error: error.message } : { success: true }
}


// ── UpSeller: pedidos do mês, vínculos de variação e baixa automática ─────────
export async function fetchPedidosMes(mes) {
  const { data, error } = await supabase.rpc('pedidos_mes_resumo', { p_mes: mes })
  return error ? { error: error.message } : (data || { pedidos: 0, unidades: 0, cancelados: 0 })
}

export async function fetchVariacoesUpseller(orgId) {
  const { data, error } = await supabase.rpc('componentes_upseller', { p_org: orgId })
  return error ? { error: error.message } : (data || [])
}

export async function fetchVinculos(orgId) {
  const { data } = await supabase.from('variacao_vinculos').select('*').eq('organizacao_id', orgId)
  return data || []
}

// variacaoId vazio e ignorar=false = remove o vínculo; ignorar=true = parte que não consome material (ex: brinde)
export async function salvarVinculo(orgId, chave, texto, variacaoId, ignorar = false) {
  if (!variacaoId && !ignorar) {
    const { error } = await supabase.from('variacao_vinculos').delete().eq('organizacao_id', orgId).eq('chave', chave)
    return error ? { error: error.message } : { success: true }
  }
  const { error } = await supabase.from('variacao_vinculos').upsert(
    { organizacao_id: orgId, chave, variacao_texto: texto, variacao_id: ignorar ? null : variacaoId, ignorar: !!ignorar },
    { onConflict: 'organizacao_id,chave' })
  return error ? { error: error.message } : { success: true }
}

export async function fetchBaixaResumo(orgId) {
  const { data, error } = await supabase.rpc('baixa_resumo', { p_org: orgId })
  return error ? { error: error.message } : (data || { ativo: false })
}

export async function ativarBaixa(orgId, ativo) {
  const { data: atual } = await supabase.from('baixa_config').select('*').eq('organizacao_id', orgId).maybeSingle()
  const payload = { organizacao_id: orgId, ativo, atualizado_em: new Date().toISOString() }
  // "desde" só é definido na primeira ativação: pedidos anteriores nunca são baixados
  if (!atual) payload.desde = new Date().toISOString()
  const { error } = await supabase.from('baixa_config').upsert(payload, { onConflict: 'organizacao_id' })
  return error ? { error: error.message } : { success: true }
}

export async function processarBaixa(orgId) {
  const { data, error } = await supabase.rpc('processar_baixa_pedidos', { p_org: orgId })
  return error ? { error: error.message } : data
}

// Baixa retroativa: pedidos já enviados desde uma data (exige a função baixa_retroativa no banco)
export async function baixaRetroativa(orgId, desdeISO) {
  const { data, error } = await supabase.rpc('baixa_retroativa', { p_org: orgId, p_desde: desdeISO })
  return error ? { error: error.message } : data
}

// Saídas dos últimos N dias (para estimar quantos dias o estoque dura)
export async function fetchConsumoRecente(userId, orgId, dias = 30) {
  const desde = new Date(Date.now() - dias * 86400000).toISOString()
  let q = supabase.from('estoque_movimentos').select('produto_id, quantidade, tipo, motivo')
    .or('tipo.eq.saida,and(tipo.eq.entrada,motivo.like.Estorno%)').gte('criado_em', desde)
  q = orgId ? q.eq('organizacao_id', orgId) : q.eq('user_id', userId)
  const { data } = await q.limit(5000)
  const mapa = {}
  ;(data || []).filter(m => !String(m.motivo || '').startsWith('Ajuste de contagem')).forEach(m => { mapa[m.produto_id] = (mapa[m.produto_id] || 0) + (m.tipo === 'entrada' ? -1 : 1) * (Number(m.quantidade) || 0) })
  return mapa
}

// Projeção salva por mês (só administradores)
export async function fetchProjecao(orgId, mes) {
  const { data } = await supabase.from('projecoes_mes').select('*').eq('organizacao_id', orgId).eq('mes', mes).maybeSingle()
  return data || null
}
export async function salvarProjecao(orgId, mes, dados, nome) {
  const { error } = await supabase.from('projecoes_mes').upsert(
    { organizacao_id: orgId, mes, dados, atualizado_em: new Date().toISOString(), atualizado_por: nome || null },
    { onConflict: 'organizacao_id,mes' })
  return error ? { error: error.message } : { success: true }
}

export async function fetchPedidosUpseller(orgId, f = {}) {
  const { data, error } = await supabase.rpc('pedidos_upseller_lista', {
    p_org: orgId, p_busca: f.busca || null, p_status: f.status || null, p_loja: f.loja || null,
    p_vinculo: f.vinculo || null, p_limit: f.limite || 50, p_offset: f.offset || 0,
  })
  return error ? { error: error.message } : (data || { total: 0, linhas: [] })
}

// ── Kits de matérias-primas (ex: Tinta 4 cores) ───────────────────────────────
export async function fetchKits(orgId) {
  if (!orgId) return []
  const { data } = await supabase.from('mp_kits').select('*').eq('organizacao_id', orgId).eq('ativo', true).order('nome')
  return data || []
}
// itens: [{ mpId, fator }]  (fator = quanto de cada matéria-prima vai em 1 unidade do kit)
export async function salvarKit(orgId, kit) {
  const payload = { nome: kit.nome, itens: kit.itens, organizacao_id: orgId }
  if (kit.id) {
    const { error } = await supabase.from('mp_kits').update({ nome: kit.nome, itens: kit.itens }).eq('id', kit.id)
    return error ? { error: error.message } : { success: true, id: kit.id }
  }
  const { data, error } = await supabase.from('mp_kits').insert(payload).select('id').single()
  return error ? { error: error.message } : { success: true, id: data.id }
}
export async function excluirKit(id) {
  const { error } = await supabase.from('mp_kits').update({ ativo: false }).eq('id', id)
  return error ? { error: error.message } : { success: true }
}

// Desfaz um lançamento de estoque (entrada ou saída): devolve/retira a quantidade do saldo e apaga o lançamento
export async function desfazerMovimentoEstoque(m, ctx) {
  const op = { tabela: 'estoque_movimentos', acao: 'criar', registro_id: m.id, depois: { id: m.id, produto_id: m.produto_id, tipo: m.tipo, quantidade: m.quantidade } }
  return _reverter(op, ctx)
}


// ─── Robustez: erros, sincronização UpSeller, backup, resumo semanal ─────────

export async function registrarErroApp(ctx, erro, contexto) {
  try {
    if (!supabase || !ctx?.userId) return
    await supabase.from('erros_app').insert({
      organizacao_id: ctx.orgId || null, user_id: ctx.userId, usuario_nome: ctx.nome || null,
      mensagem: String(erro?.message || erro || 'erro').slice(0, 1000),
      stack: String(erro?.stack || '').slice(0, 4000),
      contexto: contexto ? String(contexto).slice(0, 1000) : null,
      url: typeof location !== 'undefined' ? location.href : null,
    })
  } catch (e) { /* não pode falhar ao registrar falha */ }
}

export async function fetchErrosApp(orgId, limite = 30) {
  const { data, error } = await supabase.from('erros_app').select('*').eq('organizacao_id', orgId)
    .order('criado_em', { ascending: false }).limit(limite)
  return error ? { error: error.message } : (data || [])
}

// Última vez que o UpSeller trouxe/atualizou pedidos (a baixa de estoque depende disso)
export async function fetchSyncUpseller() {
  try {
    const [a, b] = await Promise.all([
      supabase.from('upseller_pedidos_reais').select('coletado_em').not('coletado_em', 'is', null).order('coletado_em', { ascending: false }).limit(1),
      supabase.from('upseller_pedidos_reais').select('atualizado_em').not('atualizado_em', 'is', null).order('atualizado_em', { ascending: false }).limit(1),
    ])
    const t = [a.data?.[0]?.coletado_em, b.data?.[0]?.atualizado_em].filter(Boolean).map(x => new Date(x).getTime())
    return t.length ? new Date(Math.max(...t)).toISOString() : null
  } catch (e) { return null }
}

export async function fetchBackups(orgId) {
  const { data, error } = await supabase.from('backup_snapshots').select('id, origem, resumo, criado_em')
    .eq('organizacao_id', orgId).order('criado_em', { ascending: false }).limit(12)
  return error ? { error: error.message } : (data || [])
}
export async function gerarBackupAgora(orgId) {
  const { data, error } = await supabase.rpc('gerar_backup', { p_org: orgId, p_origem: 'manual' })
  return error ? { error: error.message } : { success: true, id: data }
}
export async function fetchBackupDados(id) {
  const { data, error } = await supabase.from('backup_snapshots').select('dados, criado_em').eq('id', id).maybeSingle()
  return error || !data ? { error: error?.message || 'Backup não encontrado' } : data
}

// Números da semana (de `desdeIso` até agora) para o resumo semanal
export async function fetchResumoSemanal(userId, orgId, desdeIso) {
  const q = (extra) => {
    let r = supabase.from('upseller_pedidos_reais').select('id', { count: 'exact', head: true })
      .gte('data_pedido', desdeIso).is('excluido_em', null)
    return extra ? extra(r) : r
  }
  const [tot, canc, movs] = await Promise.all([
    q(r => r.or('status_upseller.is.null,status_upseller.neq.cancelado')),
    q(r => r.eq('status_upseller', 'cancelado')),
    supabase.from('estoque_movimentos').select('produto_id, quantidade, tipo, motivo, estoque_produtos(nome, unidade)')
      .eq('organizacao_id', orgId).gte('criado_em', desdeIso).limit(5000),
  ])
  const uso = {}
  ;(movs.data || []).forEach(m => {
    const mot = String(m.motivo || '')
    if (mot.startsWith('Ajuste de contagem')) return
    const ehSaida = m.tipo === 'saida'
    const ehEstorno = m.tipo === 'entrada' && mot.startsWith('Estorno')
    if (!ehSaida && !ehEstorno) return
    const k = m.produto_id
    uso[k] = uso[k] || { nome: m.estoque_produtos?.nome || '?', unidade: m.estoque_produtos?.unidade || '', qtd: 0 }
    uso[k].qtd += (ehSaida ? 1 : -1) * (Number(m.quantidade) || 0)
  })
  return {
    pedidos: tot.count || 0, cancelados: canc.count || 0,
    uso: Object.values(uso).filter(u => u.qtd > 0).sort((a, b) => b.qtd - a.qtd),
  }
}
