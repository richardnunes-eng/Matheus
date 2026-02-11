// =============================================================================
// CARTÕES DE CRÉDITO — Backend (server-side)
// Integrado ao Sistema Matheus existente.
//
// ATENÇÃO: Este arquivo NÃO define doGet() nem include() — essas funções
// já existem em Code.js. Aqui ficam apenas as funções da tela de Cartões.
//
// LIMITAÇÕES APPS SCRIPT TRATADAS:
// - Sem WebSockets: polling inteligente com backoff (no client)
// - CacheService TTL máx 21600s, ~100KB por chave
// - LockService: tryLock com timeout razoável
// - Sem módulos ES: funções nomeadas no escopo global
// =============================================================================

// =============================================================================
// CONFIG — Cartões
// =============================================================================
var CFG_CART = {
  SHEET_CARTOES:   'CARTOES',
  SHEET_FATURAS:   'FATURAS',
  CACHE_TTL:       30,
  CACHE_KEY_LIST:  'cartoes_list_v1',
  CACHE_KEY_ETAG:  'cartoes_etag_v1',
  LOCK_TIMEOUT_MS: 8000,
  SEED_ENABLED:    true
};

// =============================================================================
// DB — Helpers de planilha (prefixo cart_ para evitar conflito com helpers
//       que possam existir no projeto principal)
// =============================================================================

var HEADERS_CARTOES = [
  'id','nome','emissor','bandeira','ult4','cor','limite_total',
  'limite_usado','dia_fechamento','dia_vencimento','ativo',
  'criado_em','atualizado_em','version'
];

var HEADERS_FATURAS = [
  'id_fatura','id_cartao','mes_ref','ano_ref','dt_inicio','dt_fim',
  'dt_fechamento','dt_vencimento','valor','status','dt_pagamento',
  'atualizado_em','version'
];

/** Cache local do mapa de colunas (por aba). Evita re-leitura de headers. */
var _cartColMapCache = {};
var CARTAO_NOTE_MARKER = 'CARTAO_ID';

function cart_extractCardIdFromNotas_(notas) {
  var txt = String(notas || '');
  var m = txt.match(/\[CARTAO_ID:([^\]]+)\]/i) || txt.match(/CARTAO_ID:([A-Za-z0-9\-_]+)/i);
  return m ? String(m[1]).trim() : '';
}

function cart_isFaturaPaymentNotas_(notas) {
  return /\[PAGAMENTO_FATURA:[^\]]+\]/i.test(String(notas || ''));
}

function cart_isFaturaPaga_(status) {
  var s = String(status || '').trim().toLowerCase();
  return s === 'paga' || s === 'pago' || s === 'paid';
}

function cart_isFaturaParcial_(status) {
  var s = String(status || '').trim().toLowerCase();
  return s === 'parcial';
}

function cart_faturaRefKey_(fatura) {
  return (parseInt(fatura.ano_ref) || 0) * 100 + (parseInt(fatura.mes_ref) || 0);
}

function cart_toYmd_(value) {
  if (!value) return '';
  if (value instanceof Date) return cart_isoDate(value);
  var s = String(value || '').trim();
  if (!s) return '';
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) return s;
  if (s.match(/^\d{2}\/\d{2}\/\d{4}$/)) return s.substring(6, 10) + '-' + s.substring(3, 5) + '-' + s.substring(0, 2);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return cart_isoDate(d);
  return '';
}

function cart_sumCardExpensesInPeriod_(cartaoId, dtInicio, dtFim) {
  var start = cart_toYmd_(dtInicio);
  var end = cart_toYmd_(dtFim);
  if (!start || !end) return 0;
  if (typeof getSheetDataAsObjects !== 'function') return 0;
  var all = getSheetDataAsObjects('LANCAMENTOS') || [];
  return all.reduce(function(sum, tx) {
    if (String(tx.Tipo || '') !== 'DESPESA') return sum;
    if (cart_isFaturaPaymentNotas_(tx.Notas)) return sum;
    var st = String(tx.StatusParcela || '').toUpperCase();
    if (st === 'CANCELADA') return sum;
    var txCardId = cart_extractCardIdFromNotas_(tx.Notas);
    if (String(txCardId) !== String(cartaoId)) return sum;
    var ymd = cart_toYmd_(tx.Data);
    if (!ymd) return sum;
    if (ymd < start || ymd > end) return sum;
    return sum + (parseFloat(tx.Valor) || 0);
  }, 0);
}

function cart_adjustCardUsed_(cartaoId, delta) {
  var id = String(cartaoId || '');
  var d = parseFloat(delta) || 0;
  if (!id || !d) return null;
  try {
    var sheet = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
    var rows = cart_readAll(sheet);
    var idx = -1;
    var existing = null;
    rows.forEach(function(r, i) {
      if (String(r.id) === id) {
        idx = i;
        existing = r;
      }
    });
    if (!existing || idx < 0) return null;

    var atual = parseFloat(existing.limite_usado) || 0;
    var novo = atual + d;
    if (novo < 0) novo = 0;
    novo = Math.round(novo * 100) / 100;

    var updated = Object.assign({}, existing, {
      limite_usado: novo,
      atualizado_em: new Date().toISOString(),
      version: (parseInt(existing.version) || 1) + 1
    });
    cart_updateRow(sheet, idx + 2, updated, HEADERS_CARTOES);
    cart_cacheInvalidateOne(id);
    return novo;
  } catch (e) {
    Logger.log('cart_adjustCardUsed_: ' + e.message);
    return null;
  }
}

function cart_resolvePaymentCategoryId_() {
  if (typeof getCategoriesCached !== 'function') return '';
  var cats = getCategoriesCached(true) || [];
  if (!cats.length) return '';
  var i;
  for (i = 0; i < cats.length; i++) {
    var c = cats[i];
    if (!c || !c.id || !c.active) continue;
    var name = String(c.name || '').toLowerCase();
    var typeOk = (c.applicableType === 'DESPESA' || c.applicableType === 'AMBOS');
    if (typeOk && name === 'cartão de crédito') return c.id;
  }
  for (i = 0; i < cats.length; i++) {
    c = cats[i];
    if (!c || !c.id || !c.active) continue;
    name = String(c.name || '').toLowerCase();
    typeOk = (c.applicableType === 'DESPESA' || c.applicableType === 'AMBOS');
    if (typeOk && name.indexOf('cart') !== -1) return c.id;
  }
  for (i = 0; i < cats.length; i++) {
    c = cats[i];
    if (!c || !c.id || !c.active) continue;
    name = String(c.name || '').toLowerCase();
    typeOk = (c.applicableType === 'DESPESA' || c.applicableType === 'AMBOS');
    if (typeOk && name === 'outros') return c.id;
  }
  for (i = 0; i < cats.length; i++) {
    c = cats[i];
    if (!c || !c.id || !c.active) continue;
    if (c.applicableType === 'DESPESA' || c.applicableType === 'AMBOS') return c.id;
  }
  return '';
}

function cart_createFaturaPaymentTx_(fatura, cartao, accountId, dtPagamento, valor) {
  if (typeof appendRow !== 'function' || typeof nowISO !== 'function') return '';
  var categoryId = cart_resolvePaymentCategoryId_();
  if (!categoryId) throw new Error('Categoria de despesa indisponível para pagamento da fatura.');
  var categories = (typeof getCategoriesCached === 'function') ? (getCategoriesCached(true) || []) : [];
  var catName = '';
  categories.forEach(function(c) { if (String(c.id) === String(categoryId)) catName = String(c.name || ''); });
  var amount = parseFloat(valor) || 0;
  if (amount <= 0) return '';

  var dt = cart_toYmd_(dtPagamento) || cart_isoDate(new Date());
  var mm = String(fatura.mes_ref || '').padStart(2, '0');
  var yy = String(fatura.ano_ref || '');
  var descricao = 'Pagamento fatura - ' + String(cartao.nome || 'Cartão') + ' (' + mm + '/' + yy + ')';
  var notas = '[PAGAMENTO_FATURA:' + String(fatura.id_fatura) + ']\n[CARTAO_ID:' + String(fatura.id_cartao) + ']';

  var row = [
    Utilities.getUuid(), // ID
    dt, // Data
    'DESPESA', // Tipo
    categoryId, // CategoriaId
    catName, // CategoriaNomeSnapshot
    descricao, // Descrição
    amount, // Valor
    'Pagamento Fatura', // FormaPgto
    '', // KM
    nowISO(), // CriadoEm
    nowISO(), // AtualizadoEm
    '', // GrupoId
    '', // Numero
    '', // Total
    'UNICA', // TipoRecorrencia
    'PAGA', // StatusParcela
    dt, // DataVencimento
    amount, // ValorOriginal
    0, // Juros
    0, // Desconto
    dt, // DataPagamento
    'COMPENSADO', // StatusReconciliacao
    dt, // DataCompensacao
    notas, // Notas
    String(accountId || ''), // CarteiraOrigem
    '', // CarteiraDestino
    false // DebitoAutomatico
  ];
  appendRow('LANCAMENTOS', row);
  return row[0];
}

function cart_getFaturaPaymentStats_() {
  var map = {};
  try {
    if (typeof getSheetDataAsObjects !== 'function' || typeof SHEET_LANCAMENTOS === 'undefined') return map;
    var txs = getSheetDataAsObjects(SHEET_LANCAMENTOS) || [];
    txs.forEach(function(tx) {
      var notas = String((tx && tx.Notas) || '');
      var m = notas.match(/\[PAGAMENTO_FATURA:([^\]]+)\]/i);
      if (!m) return;
      var fid = String(m[1] || '').trim();
      if (!fid) return;
      var dt = String((tx && (tx.DataPagamento || tx.Data)) || '');
      var val = parseFloat((tx && tx.Valor) || 0) || 0;
      if (!map[fid]) {
        map[fid] = { dt_pagamento: dt, total_pago: 0 };
      }
      map[fid].total_pago += val;
      if (dt > String(map[fid].dt_pagamento || '')) map[fid].dt_pagamento = dt;
    });
  } catch (e) {
    Logger.log('cart_getFaturaPaymentStats_: ' + e.message);
  }
  return map;
}

function cart_getFaturaPaymentMap_() {
  return cart_getFaturaPaymentStats_();
}

function cart_getDbSpreadsheet_() {
  try {
    if (typeof getDbSpreadsheet === 'function') {
      return getDbSpreadsheet();
    }
  } catch (e) {}
  return SpreadsheetApp.getActiveSpreadsheet();
}

function cart_getOrCreateSheet(name, headers) {
  var ss    = cart_getDbSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1A73E8')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  return sheet;
}

function cart_getColMap(sheet) {
  var name = sheet.getName();
  if (_cartColMapCache[name]) return _cartColMapCache[name];
  var hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map  = {};
  hdrs.forEach(function(h, i) { map[h] = i; });
  _cartColMapCache[name] = map;
  return map;
}

function cart_readAll(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var cols = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, last - 1, cols).getValues();
  var map  = cart_getColMap(sheet);
  return data.map(function(row) {
    var obj = {};
    Object.keys(map).forEach(function(k) { obj[k] = row[map[k]]; });
    return obj;
  });
}

function cart_appendRows(sheet, rows, headers) {
  if (!rows.length) return;
  var last   = sheet.getLastRow();
  var values = rows.map(function(obj) {
    return headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  });
  sheet.getRange(last + 1, 1, values.length, headers.length).setValues(values);
}

function cart_updateRow(sheet, rowIndex, obj, headers) {
  var values = [headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; })];
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues(values);
}

// =============================================================================
// DATAS
// =============================================================================

function cart_safeDate(year, month, day) {
  var maxDay = new Date(year, month + 1, 0).getDate();
  return new Date(year, month, Math.min(day, maxDay));
}

function cart_nextOccurrence(targetDay) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var y = today.getFullYear();
  var m = today.getMonth();
  var candidate = cart_safeDate(y, m, targetDay);
  if (candidate < today) candidate = cart_safeDate(y, m + 1, targetDay);
  return candidate;
}

function cart_fmtDate(d) {
  if (!d) return '';
  if (typeof d === 'string') d = new Date(d);
  if (isNaN(d.getTime())) return '';
  return String(d.getDate()).padStart(2,'0') + '/' +
         String(d.getMonth()+1).padStart(2,'0') + '/' +
         d.getFullYear();
}

function cart_isoDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return d.getFullYear() + '-' +
    String(d.getMonth()+1).padStart(2,'0') + '-' +
    String(d.getDate()).padStart(2,'0');
}

function cart_melhorDiaCompra(diaFechamento) {
  return diaFechamento >= 28 ? 1 : diaFechamento + 1;
}

function cart_periodoFatura(diaFechamento, diaVencimento, mesRef, anoRef) {
  var m      = mesRef - 1; // 0-based
  var dtFim  = cart_safeDate(anoRef, m, diaFechamento);

  var dtIni  = new Date(dtFim);
  dtIni.setDate(dtIni.getDate() + 1);
  dtIni.setMonth(dtIni.getMonth() - 1);

  var dtVenc = cart_safeDate(anoRef, m + 1, diaVencimento);
  if (diaVencimento <= diaFechamento) dtVenc = cart_safeDate(anoRef, m, diaVencimento);

  return {
    dtInicio:     cart_isoDate(dtIni),
    dtFim:        cart_isoDate(dtFim),
    dtFechamento: cart_isoDate(dtFim),
    dtVencimento: cart_isoDate(dtVenc)
  };
}

function cart_refFaturaVigente_(diaFechamento, baseDate) {
  var d = baseDate ? new Date(baseDate) : new Date();
  if (isNaN(d.getTime())) d = new Date();
  var y = d.getFullYear();
  var m = d.getMonth() + 1;
  var dia = d.getDate();
  var refMes = (dia > diaFechamento) ? (m + 1) : m;
  var refAno = y;
  if (refMes > 12) { refMes = 1; refAno += 1; }
  return { ano: refAno, mes: refMes };
}

function cart_enrichCartao(c) {
  var diaF     = parseInt(c.dia_fechamento) || 1;
  var diaV     = parseInt(c.dia_vencimento) || 10;
  var limTotal = parseFloat(c.limite_total) || 0;
  var limUsado = parseFloat(c.limite_usado) || 0;
  var limDisp  = Math.max(0, limTotal - limUsado);
  var pctUso   = limTotal > 0 ? Math.min(100, Math.round((limUsado / limTotal) * 100)) : 0;

  var proxFech = cart_nextOccurrence(diaF);
  var proxVenc = cart_nextOccurrence(diaV);
  var hoje     = new Date(); hoje.setHours(0,0,0,0);
  var diasVenc = Math.round((proxVenc - hoje) / 86400000);

  return {
    id:               String(c.id || ''),
    nome:             String(c.nome || ''),
    emissor:          String(c.emissor || ''),
    bandeira:         String(c.bandeira || ''),
    ult4:             String(c.ult4 || ''),
    cor:              String(c.cor || '#5B5FEF'),
    limite_total:     limTotal,
    limite_usado:     limUsado,
    limite_disponivel: limDisp,
    dia_fechamento:   diaF,
    dia_vencimento:   diaV,
    ativo:            c.ativo === true || c.ativo === 'TRUE' || c.ativo === 1,
    criado_em:        String(c.criado_em  || ''),
    atualizado_em:    String(c.atualizado_em || ''),
    version:          parseInt(c.version) || 1,
    pct_uso:          pctUso,
    prox_fechamento:  cart_fmtDate(proxFech),
    prox_vencimento:  cart_fmtDate(proxVenc),
    melhor_dia_compra: cart_melhorDiaCompra(diaF),
    dias_para_vencimento: diasVenc,
    alerta_uso:       pctUso > 80,
    alerta_vencimento: diasVenc <= 5
  };
}

// =============================================================================
// CACHE
// =============================================================================
var _cartCache = CacheService.getScriptCache();

function cart_cacheScopeUser_() {
  if (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID) {
    return String(AUTH_EXECUTION_USER_ID);
  }
  return 'public';
}

function cart_cacheScopedKey_(key) {
  return String(key || '') + '|u:' + cart_cacheScopeUser_();
}

function cart_cacheGet(key) {
  try {
    var r = _cartCache.get(cart_cacheScopedKey_(key));
    return r ? JSON.parse(r) : null;
  }
  catch(e) { return null; }
}

function cart_cacheSet(key, value, ttl) {
  try {
    var s = JSON.stringify(value);
    if (s.length < 95000) _cartCache.put(cart_cacheScopedKey_(key), s, ttl || CFG_CART.CACHE_TTL);
  } catch(e) { Logger.log('cart_cacheSet: ' + e.message); }
}

function cart_cacheInvalidate() {
  try {
    _cartCache.remove(cart_cacheScopedKey_(CFG_CART.CACHE_KEY_LIST));
    _cartCache.remove(cart_cacheScopedKey_(CFG_CART.CACHE_KEY_ETAG));
  } catch(e) {}
}

function cart_cacheInvalidateOne(id) {
  try { _cartCache.remove(cart_cacheScopedKey_('cartao_det_' + id)); } catch(e) {}
  cart_cacheInvalidate();
}

// =============================================================================
// LOCK
// =============================================================================
function cart_withLock(fn) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CFG_CART.LOCK_TIMEOUT_MS)) {
      throw { code: 'LOCK_TIMEOUT', message: 'Servidor ocupado. Tente novamente.' };
    }
    return fn();
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// =============================================================================
// RESPOSTA PADRÃO
// =============================================================================
function cart_ok(data, extra) {
  var meta = Object.assign({ serverTime: new Date().toISOString(), latencyMs: 0 }, extra || {});
  return JSON.stringify({ ok: true, data: data, meta: meta });
}

function cart_err(code, message, details) {
  return JSON.stringify({
    ok: false,
    error: { code: code, message: message, details: details || null },
    meta: { serverTime: new Date().toISOString() }
  });
}

function cart_timed(fn) {
  var t0     = Date.now();
  var result = fn();
  var parsed = JSON.parse(result);
  if (parsed.meta) parsed.meta.latencyMs = Date.now() - t0;
  return JSON.stringify(parsed);
}

// =============================================================================
// SEED
// =============================================================================
function cart_seed() {
  if (!CFG_CART.SEED_ENABLED) return;
  var sheet = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
  if (sheet.getLastRow() > 1) return;
  var now   = new Date().toISOString();
  var seeds = [
    { id: Utilities.getUuid(), nome: 'Nubank', emissor: 'Nu Pagamentos', bandeira: 'Master',
      ult4: '1234', cor: '#820AD1', limite_total: 5000, limite_usado: 1800,
      dia_fechamento: 10, dia_vencimento: 17, ativo: true,
      criado_em: now, atualizado_em: now, version: 1 },
    { id: Utilities.getUuid(), nome: 'Inter Black', emissor: 'Banco Inter', bandeira: 'Master',
      ult4: '5678', cor: '#FF7A00', limite_total: 10000, limite_usado: 8500,
      dia_fechamento: 20, dia_vencimento: 27, ativo: true,
      criado_em: now, atualizado_em: now, version: 1 }
  ];
  cart_appendRows(sheet, seeds, HEADERS_CARTOES);
}

// =============================================================================
// API — bootstrap
// =============================================================================
function bootstrap() {
  return cart_timed(function() {
    try {
      cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
      cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
      cart_seed();

      var cached = cart_cacheGet(CFG_CART.CACHE_KEY_LIST);
      var etag   = cart_cacheGet(CFG_CART.CACHE_KEY_ETAG);
      if (cached) return cart_ok({ cartoes: cached, etag: etag, fromCache: true }, { etag: etag });

      var sheet   = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
      var rows    = cart_readAll(sheet);
      var cartoes = rows.map(cart_enrichCartao);
      var newEtag = String(Date.now());

      cart_cacheSet(CFG_CART.CACHE_KEY_LIST, cartoes, CFG_CART.CACHE_TTL);
      cart_cacheSet(CFG_CART.CACHE_KEY_ETAG, newEtag,  CFG_CART.CACHE_TTL);

      return cart_ok({
        cartoes: cartoes,
        counts: {
          ativos:     cartoes.filter(function(c) { return c.ativo; }).length,
          arquivados: cartoes.filter(function(c) { return !c.ativo; }).length
        },
        etag: newEtag, fromCache: false
      }, { etag: newEtag });
    } catch(e) {
      Logger.log('bootstrap: ' + e.message);
      return cart_err('UNKNOWN', 'Erro ao carregar dados: ' + e.message);
    }
  });
}

// =============================================================================
// API — listCartoes
// =============================================================================
function listCartoes(clientEtag) {
  return cart_timed(function() {
    try {
      var serverEtag = cart_cacheGet(CFG_CART.CACHE_KEY_ETAG);
      if (clientEtag && serverEtag && clientEtag === serverEtag) {
        return cart_ok({ unchanged: true, etag: serverEtag }, { etag: serverEtag });
      }
      var cached = cart_cacheGet(CFG_CART.CACHE_KEY_LIST);
      if (cached) return cart_ok({ cartoes: cached, etag: serverEtag }, { etag: serverEtag });

      var sheet   = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
      var cartoes = cart_readAll(sheet).map(cart_enrichCartao);
      var newEtag = String(Date.now());
      cart_cacheSet(CFG_CART.CACHE_KEY_LIST, cartoes, CFG_CART.CACHE_TTL);
      cart_cacheSet(CFG_CART.CACHE_KEY_ETAG, newEtag,  CFG_CART.CACHE_TTL);
      return cart_ok({ cartoes: cartoes, etag: newEtag }, { etag: newEtag });
    } catch(e) {
      return cart_err('UNKNOWN', 'Erro ao listar cartões: ' + e.message);
    }
  });
}

// =============================================================================
// API — getCartaoDetalhe
// =============================================================================
function getCartaoDetalhe(cartaoId) {
  return cart_timed(function() {
    if (!cartaoId) return cart_err('VALIDATION', 'cartaoId obrigatório');
    try {
      var cacheKey = 'cartao_det_' + cartaoId;
      var cached   = cart_cacheGet(cacheKey);
      if (cached) return cart_ok(cached);

      var sheetC = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
      var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);

      var cartaoRaw = cart_readAll(sheetC).find(function(c) { return String(c.id) === String(cartaoId); });
      if (!cartaoRaw) return cart_err('NOT_FOUND', 'Cartão não encontrado');

      var cartao  = cart_enrichCartao(cartaoRaw);
      var paymentMap = cart_getFaturaPaymentStats_();
      var faturasBase = cart_readAll(sheetF)
        .filter(function(f) { return String(f.id_cartao) === String(cartaoId); })
        .map(function(f) {
          var faturaId = String(f.id_fatura || '');
          var payInfo = paymentMap[faturaId] || null;
          var statusRaw = String(f.status || 'Aberta');
          var statusEfetivo = statusRaw;
          var valorAtual = cart_sumCardExpensesInPeriod_(cartaoId, f.dt_inicio, f.dt_fim);
          var valorArmazenado = parseFloat(f.valor) || 0;
          var totalPago = payInfo ? (parseFloat(payInfo.total_pago) || 0) : 0;
          var valorTotal = parseFloat(valorAtual) || 0;
          if (valorTotal <= 0) valorTotal = Math.max(valorArmazenado, totalPago);
          var saldoAtual = Math.max(0, valorTotal - totalPago);
          if (saldoAtual <= 0.0001 && (valorTotal > 0 || totalPago > 0 || cart_isFaturaPaga_(statusRaw))) statusEfetivo = 'Paga';
          else if (totalPago > 0 && saldoAtual > 0.0001) statusEfetivo = 'Parcial';
          else if (cart_isFaturaPaga_(statusEfetivo) || cart_isFaturaParcial_(statusEfetivo)) statusEfetivo = 'Aberta';
          var valorPago = Math.max(0, Math.min(totalPago, valorTotal > 0 ? valorTotal : totalPago));
          return {
            id_fatura:    faturaId,    id_cartao:    String(f.id_cartao),
            mes_ref:      parseInt(f.mes_ref)||0,  ano_ref:      parseInt(f.ano_ref)||0,
            dt_inicio:    String(f.dt_inicio||''), dt_fim:       String(f.dt_fim||''),
            dt_fechamento: String(f.dt_fechamento||''), dt_vencimento: String(f.dt_vencimento||''),
            valor_mes:    saldoAtual,
            valor:        saldoAtual,
            valor_total:  valorTotal,
            valor_pago:   valorPago,
            status:       statusEfetivo,
            dt_pagamento: payInfo ? String(payInfo.dt_pagamento || '') : String(f.dt_pagamento || ''),
            atualizado_em: String(f.atualizado_em||''),
            version:      parseInt(f.version)||1
          };
        });

      // Acumulado: soma meses em aberto de forma progressiva.
      faturasBase.sort(function(a,b) { return a.ano_ref !== b.ano_ref ? a.ano_ref-b.ano_ref : a.mes_ref-b.mes_ref; });
      var acumuladoAberto = 0;
      faturasBase.forEach(function(f) {
        var valorMes = parseFloat(f.valor_mes) || 0;
        if (cart_isFaturaPaga_(f.status)) {
          f.valor_acumulado = 0;
          f.valor = valorMes;
          return;
        }
        acumuladoAberto += valorMes;
        f.valor_acumulado = acumuladoAberto;
        f.valor = acumuladoAberto;
      });

      var totalEmAberto = faturasBase
        .filter(function(f) { return !cart_isFaturaPaga_(f.status); })
        .reduce(function(s,f) { return s + (parseFloat(f.valor_mes) || 0); }, 0);
      var refVig = cart_refFaturaVigente_(parseInt(cartaoRaw.dia_fechamento) || 1, new Date());
      var periodoVig = cart_periodoFatura(
        parseInt(cartaoRaw.dia_fechamento) || 1,
        parseInt(cartaoRaw.dia_vencimento) || 10,
        refVig.mes,
        refVig.ano
      );
      var valorVigente = cart_sumCardExpensesInPeriod_(String(cartaoId), periodoVig.dtInicio, periodoVig.dtFim);
      var limUsadoAtual = parseFloat(cartao.limite_usado) || 0;
      var vigFromLimite = false;
      if (valorVigente <= 0 && limUsadoAtual > 0) {
        // Fallback: quando não houver movimentos no período calculado, mas ainda existe saldo usado no cartão.
        valorVigente = limUsadoAtual;
        vigFromLimite = true;
      }
      var faturaVigenteExistente = faturasBase.find(function(f) {
        return parseInt(f.ano_ref) === refVig.ano && parseInt(f.mes_ref) === refVig.mes;
      }) || null;
      var pagamentoVig = faturaVigenteExistente ? (paymentMap[String(faturaVigenteExistente.id_fatura)] || null) : null;
      var totalPagoVig = pagamentoVig ? (parseFloat(pagamentoVig.total_pago) || 0) : 0;
      var saldoVigente = vigFromLimite
        ? Math.max(0, parseFloat(valorVigente) || 0)
        : Math.max(0, (parseFloat(valorVigente) || 0) - totalPagoVig);
      var statusVigente = 'Aberta';
      if (saldoVigente <= 0.0001 && ((parseFloat(valorVigente) || 0) > 0 || totalPagoVig > 0)) statusVigente = 'Paga';
      else if (totalPagoVig > 0 && saldoVigente > 0.0001) statusVigente = 'Parcial';
      var faturaVigente = {
        id_fatura: faturaVigenteExistente ? String(faturaVigenteExistente.id_fatura || '') : '',
        id_cartao: String(cartaoId),
        mes_ref: refVig.mes,
        ano_ref: refVig.ano,
        dt_inicio: periodoVig.dtInicio,
        dt_fim: periodoVig.dtFim,
        dt_fechamento: periodoVig.dtFechamento,
        dt_vencimento: periodoVig.dtVencimento,
        valor_mes: saldoVigente,
        valor: saldoVigente,
        valor_total: vigFromLimite ? (saldoVigente + totalPagoVig) : (parseFloat(valorVigente) || 0),
        valor_pago: totalPagoVig,
        valor_acumulado: saldoVigente,
        status: statusVigente,
        dt_pagamento: pagamentoVig ? String(pagamentoVig.dt_pagamento || '') : ''
      };
      var faturas = faturasBase
        .slice()
        .sort(function(a,b) { return a.ano_ref !== b.ano_ref ? b.ano_ref-a.ano_ref : b.mes_ref-a.mes_ref; })
        .slice(0, 12);
      var kpis = {
        total_faturas:   faturasBase.length,
        valor_em_aberto: totalEmAberto,
        faturas_pagas:   faturasBase.filter(function(f) { return cart_isFaturaPaga_(f.status); }).length
      };

      var detalhe = { cartao: cartao, faturas: faturas, kpis: kpis, fatura_vigente: faturaVigente };
      cart_cacheSet(cacheKey, detalhe, CFG_CART.CACHE_TTL);
      return cart_ok(detalhe);
    } catch(e) {
      Logger.log('getCartaoDetalhe: ' + e.message);
      return cart_err('UNKNOWN', 'Erro ao carregar detalhe: ' + e.message);
    }
  });
}

// =============================================================================
// API — saveCartao (create / update com controle de version)
// =============================================================================
function saveCartao(payload) {
  return cart_timed(function() {
    if (!payload) return cart_err('VALIDATION', 'Payload inválido');
    var p;
    try { p = typeof payload === 'string' ? JSON.parse(payload) : payload; }
    catch(e) { return cart_err('VALIDATION', 'JSON inválido'); }

    if (!p.nome || !p.nome.trim()) return cart_err('VALIDATION', 'Nome obrigatório');
    if (!p.bandeira)               return cart_err('VALIDATION', 'Bandeira obrigatória');
    if (!p.limite_total && p.limite_total !== 0) return cart_err('VALIDATION', 'Limite total obrigatório');
    if (!p.dia_fechamento) return cart_err('VALIDATION', 'Dia de fechamento obrigatório');
    if (!p.dia_vencimento) return cart_err('VALIDATION', 'Dia de vencimento obrigatório');

    try {
      return cart_withLock(function() {
        var sheet = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        var rows  = cart_readAll(sheet);
        var now   = new Date().toISOString();

        if (!p.id) {
          // CREATE
          var novo = {
            id: Utilities.getUuid(), nome: String(p.nome).trim(),
            emissor: String(p.emissor||'').trim(), bandeira: String(p.bandeira),
            ult4: String(p.ult4||'').replace(/\D/g,'').slice(-4),
            cor: String(p.cor||'#5B5FEF'),
            limite_total: parseFloat(p.limite_total)||0, limite_usado: parseFloat(p.limite_usado)||0,
            dia_fechamento: parseInt(p.dia_fechamento)||1, dia_vencimento: parseInt(p.dia_vencimento)||10,
            ativo: true, criado_em: now, atualizado_em: now, version: 1
          };
          cart_appendRows(sheet, [novo], HEADERS_CARTOES);
          cart_cacheInvalidate();
          return cart_ok(cart_enrichCartao(novo), { created: true });
        }

        // UPDATE
        var idx = -1; var existing = null;
        rows.forEach(function(r,i) { if (String(r.id) === String(p.id)) { idx=i; existing=r; } });
        if (!existing) return cart_err('NOT_FOUND', 'Cartão não encontrado');

        var sVer = parseInt(existing.version)||1;
        var cVer = parseInt(p.version)||1;
        if (cVer !== sVer) {
          return cart_err('CONFLICT',
            'Cartão foi atualizado em outro lugar. Recarregue e tente novamente.',
            { serverData: cart_enrichCartao(existing) });
        }

        var limiteTotalParsed = parseFloat(p.limite_total);
        var limiteUsadoParsed = parseFloat(p.limite_usado);
        var updated = {
          id: String(existing.id), nome: String(p.nome).trim(),
          emissor: String(p.emissor||'').trim(), bandeira: String(p.bandeira),
          ult4: String(p.ult4||'').replace(/\D/g,'').slice(-4),
          cor: String(p.cor||existing.cor||'#5B5FEF'),
          limite_total: isFinite(limiteTotalParsed) ? limiteTotalParsed : 0,
          limite_usado: isFinite(limiteUsadoParsed) ? limiteUsadoParsed : (parseFloat(existing.limite_usado) || 0),
          dia_fechamento: parseInt(p.dia_fechamento)||parseInt(existing.dia_fechamento)||1,
          dia_vencimento: parseInt(p.dia_vencimento)||parseInt(existing.dia_vencimento)||10,
          ativo: existing.ativo === true || existing.ativo === 'TRUE' || existing.ativo === 1,
          criado_em: String(existing.criado_em||now), atualizado_em: now, version: sVer+1
        };
        cart_updateRow(sheet, idx+2, updated, HEADERS_CARTOES);
        cart_cacheInvalidateOne(String(p.id));
        return cart_ok(cart_enrichCartao(updated), { updated: true });
      });
    } catch(e) {
      if (e.code) return cart_err(e.code, e.message);
      Logger.log('saveCartao: ' + e.message);
      return cart_err('UNKNOWN', 'Erro ao salvar cartão: ' + e.message);
    }
  });
}

// =============================================================================
// API — archiveCartao
// =============================================================================
function archiveCartao(id, ativo) {
  return cart_timed(function() {
    if (!id) return cart_err('VALIDATION', 'id obrigatório');
    try {
      return cart_withLock(function() {
        var sheet = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        var rows  = cart_readAll(sheet);
        var idx   = -1; var existing = null;
        rows.forEach(function(r,i) { if (String(r.id)===String(id)) { idx=i; existing=r; } });
        if (!existing) return cart_err('NOT_FOUND', 'Cartão não encontrado');

        var updated = Object.assign({}, existing, {
          ativo: ativo === true || ativo === 'true',
          atualizado_em: new Date().toISOString(),
          version: (parseInt(existing.version)||1) + 1
        });
        cart_updateRow(sheet, idx+2, updated, HEADERS_CARTOES);
        cart_cacheInvalidateOne(String(id));
        return cart_ok(cart_enrichCartao(updated));
      });
    } catch(e) {
      if (e.code) return cart_err(e.code, e.message);
      Logger.log('archiveCartao: ' + e.message);
      return cart_err('UNKNOWN', 'Erro ao arquivar/ativar: ' + e.message);
    }
  });
}

// =============================================================================
// API — gerarFaturaMes
// =============================================================================
function gerarFaturaMes(cartaoId, ano, mes) {
  return cart_timed(function() {
    if (!cartaoId||!ano||!mes) return cart_err('VALIDATION','cartaoId, ano e mes obrigatórios');
    ano = parseInt(ano); mes = parseInt(mes);
    if (mes<1||mes>12) return cart_err('VALIDATION','Mês inválido (1-12)');
    try {
      return cart_withLock(function() {
        var sheetC = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);

        var cartaoRaw = cart_readAll(sheetC).find(function(c) { return String(c.id)===String(cartaoId); });
        if (!cartaoRaw) return cart_err('NOT_FOUND','Cartão não encontrado');

        // Idempotência
        var existe = cart_readAll(sheetF).find(function(f) {
          return String(f.id_cartao)===String(cartaoId) &&
                 parseInt(f.mes_ref)===mes && parseInt(f.ano_ref)===ano;
        });
        if (existe) return cart_ok({ fatura:existe, created:false, message:'Fatura já existe' });

        var periodo  = cart_periodoFatura(parseInt(cartaoRaw.dia_fechamento)||1, parseInt(cartaoRaw.dia_vencimento)||10, mes, ano);
        var valorCalc = cart_sumCardExpensesInPeriod_(cartaoId, periodo.dtInicio, periodo.dtFim);
        var novaFat  = {
          id_fatura: Utilities.getUuid(), id_cartao: String(cartaoId),
          mes_ref: mes, ano_ref: ano,
          dt_inicio: periodo.dtInicio, dt_fim: periodo.dtFim,
          dt_fechamento: periodo.dtFechamento, dt_vencimento: periodo.dtVencimento,
          valor: valorCalc, status: 'Aberta', dt_pagamento: '',
          atualizado_em: new Date().toISOString(), version: 1
        };
        cart_appendRows(sheetF, [novaFat], HEADERS_FATURAS);
        cart_cacheInvalidateOne(String(cartaoId));
        return cart_ok({ fatura: novaFat, created: true });
      });
    } catch(e) {
      if (e.code) return cart_err(e.code, e.message);
      Logger.log('gerarFaturaMes: ' + e.message);
      return cart_err('UNKNOWN','Erro ao gerar fatura: ' + e.message);
    }
  });
}

// =============================================================================
// API — pagarFatura (idempotente)
// =============================================================================
function pagarFatura(faturaId, dtPagamento, accountId) {
  return cart_timed(function() {
    if (!faturaId) return cart_err('VALIDATION','faturaId obrigatório');
    if (!accountId) return cart_err('VALIDATION','Conta de pagamento obrigatória');
    dtPagamento = dtPagamento || cart_isoDate(new Date());
    try {
      return cart_withLock(function() {
        var sheetC  = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        var sheetF  = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
        var faturas = cart_readAll(sheetF);
        var fatura = null;
        faturas.forEach(function(f) { if (String(f.id_fatura)===String(faturaId)) fatura=f; });
        if (!fatura) return cart_err('NOT_FOUND','Fatura não encontrada');

        var cartao = cart_readAll(sheetC).find(function(c) { return String(c.id) === String(fatura.id_cartao); });
        if (!cartao) return cart_err('NOT_FOUND','Cartão da fatura não encontrado');
        if (typeof getAccountsCached === 'function') {
          var contas = getAccountsCached(true) || [];
          var conta = contas.find(function(a) { return String(a.id) === String(accountId); });
          if (!conta) return cart_err('VALIDATION', 'Conta de pagamento não encontrada.');
          if (conta.ativo === false) return cart_err('VALIDATION', 'Conta de pagamento está inativa.');
        }

        var alvoRef = cart_faturaRefKey_(fatura);
        var faturasDoCartao = faturas.filter(function(f) {
          return String(f.id_cartao) === String(fatura.id_cartao);
        });
        var faturasParaPagar = faturasDoCartao.filter(function(f) {
          return !cart_isFaturaPaga_(f.status) && cart_faturaRefKey_(f) <= alvoRef;
        });
        if (!faturasParaPagar.length) return cart_ok({ fatura:fatura, message:'Fatura já estava paga' });

        var valoresPorFatura = {};
        var totalPago = 0;
        faturasParaPagar.forEach(function(f) {
          var v = cart_sumCardExpensesInPeriod_(String(f.id_cartao), f.dt_inicio, f.dt_fim);
          valoresPorFatura[String(f.id_fatura)] = v;
          totalPago += (parseFloat(v) || 0);
        });

        var txId = cart_createFaturaPaymentTx_(fatura, cartao, accountId, dtPagamento, totalPago);
        if (!txId) return cart_err('TX_CREATE_FAILED', 'Falha ao registrar a despesa de pagamento da fatura.');
        if (totalPago > 0) cart_adjustCardUsed_(String(fatura.id_cartao), -totalPago);

        var updatedTarget = null;
        var nowIso = new Date().toISOString();
        faturas.forEach(function(f, i) {
          if (String(f.id_cartao) !== String(fatura.id_cartao)) return;
          if (!valoresPorFatura.hasOwnProperty(String(f.id_fatura))) return;
          var updated = Object.assign({}, f, {
            valor: valoresPorFatura[String(f.id_fatura)],
            status: 'Paga', dt_pagamento: String(dtPagamento),
            atualizado_em: nowIso,
            version: (parseInt(f.version)||1)+1
          });
          cart_updateRow(sheetF, i+2, updated, HEADERS_FATURAS);
          if (String(updated.id_fatura) === String(faturaId)) updatedTarget = updated;
        });

        cart_cacheInvalidateOne(String(fatura.id_cartao));
        return cart_ok({
          fatura: updatedTarget || fatura,
          paymentTxId: txId,
          total_pago: totalPago,
          faturas_pagas: faturasParaPagar.length
        });
      });
    } catch(e) {
      if (e.code) return cart_err(e.code, e.message);
      Logger.log('pagarFatura: ' + e.message);
      return cart_err('UNKNOWN','Erro ao pagar fatura: ' + e.message);
    }
  });
}

// =============================================================================
// API — pagarFaturaAtualSaldo
// Paga apenas a fatura vigente (ano/mes informado), com base no valor
// calculado do período da própria fatura (não no limite_usado total).
// =============================================================================
function pagarFaturaAtualSaldo(cartaoId, dtPagamento, accountId, ano, mes, valorPagamento) {
  return cart_timed(function() {
    if (!cartaoId) return cart_err('VALIDATION','cartaoId obrigatório');
    if (!accountId) return cart_err('VALIDATION','Conta de pagamento obrigatória');
    dtPagamento = dtPagamento || cart_isoDate(new Date());
    ano = parseInt(ano) || parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy'));
    mes = parseInt(mes) || parseInt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM'));
    if (mes < 1 || mes > 12) return cart_err('VALIDATION','Mês inválido (1-12)');

    try {
      return cart_withLock(function() {
        var sheetC  = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        var sheetF  = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
        var cartoes = cart_readAll(sheetC);
        var cartao = cartoes.find(function(c) { return String(c.id) === String(cartaoId); });
        if (!cartao) return cart_err('NOT_FOUND','Cartão não encontrado');

        if (typeof getAccountsCached === 'function') {
          var contas = getAccountsCached(true) || [];
          var conta = contas.find(function(a) { return String(a.id) === String(accountId); });
          if (!conta) return cart_err('VALIDATION', 'Conta de pagamento não encontrada.');
          if (conta.ativo === false) return cart_err('VALIDATION', 'Conta de pagamento está inativa.');
        }

        var faturas = cart_readAll(sheetF);
        var alvo = faturas.find(function(f) {
          return String(f.id_cartao) === String(cartaoId) &&
                 parseInt(f.ano_ref) === ano &&
                 parseInt(f.mes_ref) === mes;
        });
        if (!alvo) {
          var periodo = cart_periodoFatura(parseInt(cartao.dia_fechamento)||1, parseInt(cartao.dia_vencimento)||10, mes, ano);
          alvo = {
            id_fatura: Utilities.getUuid(), id_cartao: String(cartaoId),
            mes_ref: mes, ano_ref: ano,
            dt_inicio: periodo.dtInicio, dt_fim: periodo.dtFim,
            dt_fechamento: periodo.dtFechamento, dt_vencimento: periodo.dtVencimento,
            valor: 0, status: 'Aberta', dt_pagamento: '',
            atualizado_em: new Date().toISOString(), version: 1
          };
          cart_appendRows(sheetF, [alvo], HEADERS_FATURAS);
          faturas = cart_readAll(sheetF);
        }

        var valorVigente = cart_sumCardExpensesInPeriod_(String(cartaoId), alvo.dt_inicio, alvo.dt_fim);
        var limUsadoAtual = parseFloat(cartao.limite_usado) || 0;
        var useLimiteAsSaldo = false;
        if (valorVigente <= 0 && limUsadoAtual > 0) {
          // Fallback para permitir abatimento do saldo usado quando o período vigente não retornar lançamentos.
          valorVigente = limUsadoAtual;
          useLimiteAsSaldo = true;
        }
        if (valorVigente <= 0) {
          return cart_ok({
            message: 'Fatura vigente sem valor para pagar.',
            total_pago: 0,
            faturas_pagas: 0,
            referencia: { ano: ano, mes: mes }
          });
        }
        var paymentStats = cart_getFaturaPaymentStats_();
        var hist = paymentStats[String(alvo.id_fatura)] || null;
        var valorJaPago = hist ? (parseFloat(hist.total_pago) || 0) : 0;
        var saldoRestante = useLimiteAsSaldo
          ? Math.max(0, valorVigente)
          : Math.max(0, valorVigente - valorJaPago);
        var valorInformado = parseFloat(valorPagamento);
        var valorAPagar = isFinite(valorInformado) && valorInformado > 0 ? valorInformado : saldoRestante;
        if (valorAPagar > saldoRestante + 0.0001) {
          return cart_err('VALIDATION', 'Valor informado maior que o saldo da fatura atual.');
        }
        if (valorAPagar <= 0) {
          return cart_ok({
            message: 'Fatura vigente já estava paga.',
            total_pago: 0,
            faturas_pagas: 0,
            referencia: { ano: ano, mes: mes }
          });
        }

        var txId = cart_createFaturaPaymentTx_(alvo, cartao, accountId, dtPagamento, valorAPagar);
        if (!txId) return cart_err('TX_CREATE_FAILED', 'Falha ao registrar a despesa de pagamento da fatura.');
        if (valorAPagar > 0) cart_adjustCardUsed_(String(cartaoId), -valorAPagar);

        var countPaid = 0;
        var pagoTotalNovo = valorJaPago + valorAPagar;
        var saldoNovo = useLimiteAsSaldo
          ? Math.max(0, saldoRestante - valorAPagar)
          : Math.max(0, valorVigente - pagoTotalNovo);
        var statusNovo = 'Aberta';
        if (saldoNovo <= 0.0001) statusNovo = 'Paga';
        else if (pagoTotalNovo > 0) statusNovo = 'Parcial';
        faturas.forEach(function(f, i) {
          if (String(f.id_fatura) !== String(alvo.id_fatura)) return;
          var updated = Object.assign({}, f, {
            status: statusNovo,
            dt_pagamento: saldoNovo <= 0.0001 ? String(dtPagamento) : '',
            atualizado_em: new Date().toISOString(),
            version: (parseInt(f.version) || 1) + 1
          });
          updated.valor = saldoNovo;
          cart_updateRow(sheetF, i + 2, updated, HEADERS_FATURAS);
          countPaid++;
        });

        cart_cacheInvalidateOne(String(cartaoId));
        return cart_ok({
          total_pago: valorAPagar,
          valor_em_aberto: saldoNovo,
          valor_total: useLimiteAsSaldo ? (saldoNovo + pagoTotalNovo) : valorVigente,
          valor_pago_acumulado: pagoTotalNovo,
          faturas_pagas: countPaid,
          paymentTxId: txId,
          referencia: { ano: ano, mes: mes }
        });
      });
    } catch (e) {
      if (e.code) return cart_err(e.code, e.message);
      Logger.log('pagarFaturaAtualSaldo: ' + e.message);
      return cart_err('UNKNOWN','Erro ao pagar fatura atual: ' + e.message);
    }
  });
}

// =============================================================================
// API — syncCartoes (polling leve via ETag)
// =============================================================================
function syncCartoes(clientEtag) {
  return listCartoes(clientEtag);
}

// =============================================================================
// API — deleteCartaoCascade
// Exclui cartão + faturas + lançamentos relacionados ao cartão.
// =============================================================================
function deleteCartaoCascade(cartaoId) {
  return cart_timed(function() {
    if (!cartaoId) return cart_err('VALIDATION', 'cartaoId obrigatório');
    try {
      return cart_withLock(function() {
        var id = String(cartaoId);
        var sheetC = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);

        var rowsC = cart_readAll(sheetC) || [];
        var idxCard = -1;
        for (var i = 0; i < rowsC.length; i++) {
          if (String(rowsC[i].id) === id) { idxCard = i; break; }
        }
        if (idxCard < 0) return cart_err('NOT_FOUND', 'Cartão não encontrado');

        var removedFaturas = 0;
        if (sheetF.getLastRow() >= 2) {
          var rawF = sheetF.getRange(2, 1, sheetF.getLastRow() - 1, sheetF.getLastColumn()).getValues();
          var hdrF = sheetF.getRange(1, 1, 1, sheetF.getLastColumn()).getValues()[0];
          var mapF = {};
          hdrF.forEach(function(h, idx) { mapF[h] = idx; });
          for (var rf = rawF.length - 1; rf >= 0; rf--) {
            var idCartaoF = String(rawF[rf][mapF.id_cartao] || '');
            if (idCartaoF !== id) continue;
            sheetF.deleteRow(rf + 2);
            removedFaturas++;
          }
        }

        var removedLancamentos = 0;
        if (typeof SHEET_LANCAMENTOS !== 'undefined') {
          var ss = cart_getDbSpreadsheet_();
          var sheetL = ss.getSheetByName(SHEET_LANCAMENTOS);
          if (sheetL && sheetL.getLastRow() >= 2) {
            var rawL = sheetL.getRange(2, 1, sheetL.getLastRow() - 1, sheetL.getLastColumn()).getValues();
            var hdrL = sheetL.getRange(1, 1, 1, sheetL.getLastColumn()).getValues()[0];
            var mapL = {};
            hdrL.forEach(function(h, idx) { mapL[h] = idx; });
            for (var rl = rawL.length - 1; rl >= 0; rl--) {
              var notas = String(rawL[rl][mapL.Notas] || '');
              var txCardId = cart_extractCardIdFromNotas_(notas);
              if (String(txCardId) !== id) continue;
              sheetL.deleteRow(rl + 2);
              removedLancamentos++;
            }
          }
        }

        sheetC.deleteRow(idxCard + 2);

        cart_cacheInvalidate();
        try { if (typeof invalidateAllCache === 'function') invalidateAllCache(); } catch (e) {}
        try { if (typeof recalculateReserves === 'function') recalculateReserves(); } catch (e2) {}

        return cart_ok({
          deleted: true,
          cartaoId: id,
          removed_faturas: removedFaturas,
          removed_lancamentos: removedLancamentos
        });
      });
    } catch (e) {
      if (e.code) return cart_err(e.code, e.message);
      return cart_err('UNKNOWN', 'Erro ao excluir cartão: ' + e.message);
    }
  });
}

// =============================================================================
// API — deletePaidFaturasHistory
// Exclui somente o histórico de faturas pagas de um cartão.
// =============================================================================
function deletePaidFaturasHistory(cartaoId) {
  return cart_timed(function() {
    if (!cartaoId) return cart_err('VALIDATION', 'cartaoId obrigatório');
    try {
      return cart_withLock(function() {
        var id = String(cartaoId);
        var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
        if (sheetF.getLastRow() < 2) return cart_ok({ removed: 0, cartaoId: id });

        var paymentMap = cart_getFaturaPaymentMap_();
        var rawF = sheetF.getRange(2, 1, sheetF.getLastRow() - 1, sheetF.getLastColumn()).getValues();
        var hdrF = sheetF.getRange(1, 1, 1, sheetF.getLastColumn()).getValues()[0];
        var mapF = {};
        hdrF.forEach(function(h, idx) { mapF[h] = idx; });

        var removed = 0;
        for (var i = rawF.length - 1; i >= 0; i--) {
          var idCartao = String(rawF[i][mapF.id_cartao] || '');
          if (idCartao !== id) continue;
          var faturaId = String(rawF[i][mapF.id_fatura] || '');
          var status = String(rawF[i][mapF.status] || '');
          var isPaid = cart_isFaturaPaga_(status) || !!paymentMap[faturaId];
          if (!isPaid) continue;
          sheetF.deleteRow(i + 2);
          removed++;
        }

        cart_cacheInvalidateOne(id);
        return cart_ok({ removed: removed, cartaoId: id });
      });
    } catch (e) {
      if (e.code) return cart_err(e.code, e.message);
      return cart_err('UNKNOWN', 'Erro ao limpar faturas pagas: ' + e.message);
    }
  });
}

// =============================================================================
// API — deleteFaturaById
// Exclui uma fatura específica pelo ID.
// =============================================================================
function deleteFaturaById(faturaId) {
  return cart_timed(function() {
    if (!faturaId) return cart_err('VALIDATION', 'faturaId obrigatório');
    try {
      return cart_withLock(function() {
        var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
        var rows = cart_readAll(sheetF) || [];
        var idx = -1;
        var alvo = null;
        for (var i = 0; i < rows.length; i++) {
          if (String(rows[i].id_fatura) === String(faturaId)) {
            idx = i;
            alvo = rows[i];
            break;
          }
        }
        if (idx < 0) return cart_err('NOT_FOUND', 'Fatura não encontrada');

        sheetF.deleteRow(idx + 2);
        cart_cacheInvalidateOne(String(alvo.id_cartao || ''));
        return cart_ok({
          deleted: true,
          faturaId: String(faturaId),
          cartaoId: String(alvo.id_cartao || '')
        });
      });
    } catch (e) {
      if (e.code) return cart_err(e.code, e.message);
      return cart_err('UNKNOWN', 'Erro ao excluir fatura: ' + e.message);
    }
  });
}

// =============================================================================
// API — getCartoesFaturaMes
// Retorna valor da fatura por cartão para o mês/ano de referência informado.
// =============================================================================
function getCartoesFaturaMes(ano, mes) {
  return cart_timed(function() {
    ano = parseInt(ano, 10);
    mes = parseInt(mes, 10);
    if (!ano || !mes || mes < 1 || mes > 12) {
      return cart_err('VALIDATION', 'Ano e mês inválidos.');
    }
    try {
      var sheetC = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
      var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
      var cartoes = cart_readAll(sheetC) || [];
      var faturas = cart_readAll(sheetF) || [];
      var paymentMap = cart_getFaturaPaymentStats_();
      var itens = cartoes.map(function(c) {
        var diaF = parseInt(c.dia_fechamento, 10) || 1;
        var diaV = parseInt(c.dia_vencimento, 10) || 10;
        var periodo = cart_periodoFatura(diaF, diaV, mes, ano);
        var valorMes = cart_sumCardExpensesInPeriod_(String(c.id || ''), periodo.dtInicio, periodo.dtFim);
        var valorTotal = parseFloat(valorMes) || 0;
        // Subtrair pagamentos parciais já realizados
        var faturaDoMes = faturas.find(function(f) {
          return String(f.id_cartao) === String(c.id) &&
                 parseInt(f.ano_ref) === ano && parseInt(f.mes_ref) === mes;
        });
        var valorPago = 0;
        var status = 'Aberta';
        if (faturaDoMes) {
          var payInfo = paymentMap[String(faturaDoMes.id_fatura)] || null;
          valorPago = payInfo ? (parseFloat(payInfo.total_pago) || 0) : 0;
          status = String(faturaDoMes.status || 'Aberta');
        }
        var valorPendente = Math.max(0, valorTotal - valorPago);
        if (valorPendente <= 0.0001 && (valorTotal > 0 || valorPago > 0)) status = 'Paga';
        else if (valorPago > 0 && valorPendente > 0.0001) status = 'Parcial';
        return {
          id_cartao: String(c.id || ''),
          ano_ref: ano,
          mes_ref: mes,
          valor_mes: valorPendente,
          valor_total: valorTotal,
          valor_pago: valorPago,
          status: status,
          dt_inicio: periodo.dtInicio,
          dt_fim: periodo.dtFim
        };
      });
      return cart_ok({ ano_ref: ano, mes_ref: mes, itens: itens });
    } catch (e) {
      return cart_err('UNKNOWN', 'Erro ao calcular fatura mensal: ' + e.message);
    }
  });
}

// =============================================================================
// API — getCartoesCSS
// Retorna o conteúdo CSS do arquivo cartoes_styles.html como string pura.
// Chamado pelo client (app_js.html) para injetar estilos dinamicamente com
// adaptação de variáveis CSS ao tema dark do app principal.
// =============================================================================
function getCartoesCSS() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('cartoes_styles').getContent();
    // Remove as tags <style> e </style> — retorna apenas o CSS puro
    return html
      .replace(/^\s*<style[^>]*>/i, '')
      .replace(/<\/style>\s*$/i, '');
  } catch (e) {
    Logger.log('getCartoesCSS error: ' + e.message);
    return ''; // fallback: sem CSS extra, os estilos inline ainda funcionam
  }
}
