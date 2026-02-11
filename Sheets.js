/* ================================================================
   Sheets.gs — Camada de banco de dados (Google Sheets)
   App Financeiro do Motorista
   ================================================================ */

// ── Constantes de abas e cabeçalhos ─────────────────────────────

var SHEET_LANCAMENTOS = 'LANCAMENTOS';
var SHEET_CATEGORIAS = 'CATEGORIAS';
var SHEET_CONFIG = 'CONFIG';
var SHEET_RESERVAS = 'RESERVAS';
var SHEET_PARCELAS_CONFIG = 'PARCELAS_CONFIG';
var SHEET_CONTAS = 'CONTAS';

// ── Legacy wallet IDs (deprecated - usar CONTAS) ─────────────────
var WALLET_PRINCIPAL = 'principal';
var WALLET_RESERVA = 'reserva';

var HEADERS = {};
HEADERS[SHEET_LANCAMENTOS] = [
  'ID', 'Data', 'Tipo', 'CategoriaId', 'CategoriaNomeSnapshot',
  'Descrição', 'Valor', 'FormaPgto', 'KM', 'CriadoEm', 'AtualizadoEm',
  'GrupoId', 'Numero', 'Total',
  'TipoRecorrencia', 'StatusParcela', 'DataVencimento', 'ValorOriginal',
  'Juros', 'Desconto', 'DataPagamento',
  'StatusReconciliacao', 'DataCompensacao', 'Notas',
  'CarteiraOrigem', 'CarteiraDestino',
  'DebitoAutomatico' // NOVO: Flag para processamento automático
];

// ... (existing code) ...

// ── Calcular saldo de conta ──────────────────────────────────────

function calculateAccountBalance(accountId) {
  var account = null;
  var accounts = getSheetDataAsObjects(SHEET_CONTAS);

  for (var i = 0; i < accounts.length; i++) {
    if (String(accounts[i].ID) === String(accountId)) {
      account = accounts[i];
      break;
    }
  }

  if (!account) return 0;

  var saldoInicial = parseFloat(account.SaldoInicial) || 0;
  var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);

  var balance = saldoInicial;

  allTx.forEach(function (tx) {
    // IMPORTANTE: Apenas transações PAGAS afetam o saldo
    // Se StatusParcela não existir (legado), assume PAGA
    var status = tx.StatusParcela || 'PAGA';
    if (status !== 'PAGA' && status !== 'COMPENSADO') return;

    var valor = parseFloat(tx.Valor) || 0;
    var tipo = String(tx.Tipo);
    var contaOrigem = String(tx.CarteiraOrigem || '');
    var contaDestino = String(tx.CarteiraDestino || '');

    // RECEITA: adiciona ao saldo da conta destino
    if (tipo === 'RECEITA' && contaDestino === accountId) {
      balance += valor;
    }

    // DESPESA: subtrai do saldo da conta origem
    if (tipo === 'DESPESA' && contaOrigem === accountId) {
      balance -= valor;
    }

    // TRANSFERENCIA: subtrai da origem, adiciona no destino
    if (tipo === 'TRANSFERENCIA' || tipo === 'TRANSFER') {
      if (contaOrigem === accountId) {
        balance -= valor;
      }
      if (contaDestino === accountId) {
        balance += valor;
      }
    }
  });

  return balance;
}
HEADERS[SHEET_CATEGORIAS] = [
  'ID', 'Nome', 'TipoAplicavel', 'Emoji', 'CorHex', 'Ordem', 'Ativo', 'CriadoEm', 'AtualizadoEm', 'OwnerUserId'
];
HEADERS[SHEET_CONFIG] = ['Chave', 'Valor'];
HEADERS[SHEET_RESERVAS] = [
  'Mes', 'ReservaManutencaoPrevista', 'ReservaManutencaoAcumulada',
  'UsosManutencaoNoMes', 'SaldoReserva'
];
HEADERS[SHEET_PARCELAS_CONFIG] = [
  'GrupoId', 'Tipo', 'Frequencia', 'DiaVencimento',
  'ValorTotal', 'NumParcelas', 'ParcelasPagas', 'ParcelasCanceladas',
  'PermitirEditacaoIndividual', 'CalcularJuros', 'TaxaJuros',
  'Observacoes', 'CriadoEm', 'AtualizadoEm'
];
HEADERS[SHEET_CONTAS] = [
  'ID', 'Nome', 'Tipo', 'SaldoInicial', 'Ativo', 'Ordem', 'CriadoEm', 'AtualizadoEm', 'Instituicao'
];

// ── Obter Spreadsheet (vinculada > salva > criar) ───────────────

function getDbSpreadsheet() {
  if (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID) {
    return getScopedDbSpreadsheet_(String(AUTH_EXECUTION_USER_ID));
  }

  // (a) Tentar planilha vinculada
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) { /* não vinculada */ }

  // (b) Tentar ID salvo
  var props = PropertiesService.getScriptProperties();
  var savedId = props.getProperty('DB_SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      Logger.log('Spreadsheet salva não encontrada, criando nova...');
    }
  }

  // (c) Criar nova
  var newSs = SpreadsheetApp.create('App Financeiro do Motorista - DB');
  props.setProperty('DB_SPREADSHEET_ID', newSs.getId());
  Logger.log('Nova planilha criada: ' + newSs.getId());
  return newSs;
}

function getGlobalDbSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var savedId = props.getProperty('DB_SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      props.deleteProperty('DB_SPREADSHEET_ID');
    }
  }

  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      props.setProperty('DB_SPREADSHEET_ID', active.getId());
      return active;
    }
  } catch (e2) {}

  var created = SpreadsheetApp.create('App Financeiro do Motorista - DB');
  props.setProperty('DB_SPREADSHEET_ID', created.getId());
  return created;
}

function getScopedDbSpreadsheet_(userId) {
  var props = PropertiesService.getScriptProperties();
  var key = 'DB_SPREADSHEET_ID_USER_' + userId;
  var savedId = props.getProperty(key);

  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      props.deleteProperty(key);
    }
  }

  var created = SpreadsheetApp.create('App Financeiro - User ' + userId.substring(0, 8));
  props.setProperty(key, created.getId());
  bootstrapScopedSpreadsheet_(created);
  return created;
}

function bootstrapScopedSpreadsheet_(ss) {
  ensureSheet(ss, SHEET_LANCAMENTOS);
  ensureSheet(ss, SHEET_CATEGORIAS);
  ensureSheet(ss, SHEET_CONFIG);
  ensureSheet(ss, SHEET_RESERVAS);
  ensureSheet(ss, SHEET_PARCELAS_CONFIG);
  ensureSheet(ss, SHEET_CONTAS);

  var catSheet = ss.getSheetByName(SHEET_CATEGORIAS);
  if (catSheet && catSheet.getLastRow() < 2) {
    var cfgSheet = ss.getSheetByName(SHEET_CONFIG);
    var publicCats = getPublicCategoriesRaw_();
    var manutId = '';
    for (var i = 0; i < publicCats.length; i++) {
      if (String(publicCats[i].Nome || '').toLowerCase() === 'manutenção') {
        manutId = String(publicCats[i].ID || '');
        break;
      }
    }
    var defaultCfg = getDefaultConfig(manutId);
    defaultCfg.forEach(function (row) { cfgSheet.appendRow(row); });

    var contasSheet = ss.getSheetByName(SHEET_CONTAS);
    seedScopedAccountsFromGlobal_(ss);
    deactivateGenericBankAccounts_(contasSheet);

    // Garantia final para evitar app sem contas.
    if (contasSheet.getLastRow() < 2) {
      var defaultAccounts = createDefaultAccounts();
      defaultAccounts.forEach(function (row) { contasSheet.appendRow(row); });
    }
  }
}

function seedScopedAccountsFromGlobal_(scopedSs) {
  var scopedSheet = scopedSs.getSheetByName(SHEET_CONTAS);
  if (!scopedSheet) return;

  var existing = [];
  if (scopedSheet.getLastRow() >= 2) {
    existing = scopedSheet.getRange(2, 1, scopedSheet.getLastRow() - 1, HEADERS[SHEET_CONTAS].length).getValues();
  }

  var existingKeys = {};
  existing.forEach(function (row) {
    var key = normalizeAccountSeedKey_(row[1], row[2], row[8]);
    existingKeys[key] = true;
  });

  var rowsToInsert = [];

  // Defaults sempre presentes
  var defaults = createDefaultAccounts();
  defaults.forEach(function (row) {
    var key = normalizeAccountSeedKey_(row[1], row[2], row[8]);
    if (!existingKeys[key]) {
      existingKeys[key] = true;
      rowsToInsert.push(row);
    }
  });

  if (rowsToInsert.length > 0) {
    scopedSheet.getRange(scopedSheet.getLastRow() + 1, 1, rowsToInsert.length, HEADERS[SHEET_CONTAS].length).setValues(rowsToInsert);
  }
}

function normalizeAccountSeedKey_(nome, tipo, instituicao) {
  return String(nome || '').trim().toLowerCase() + '|' +
    String(tipo || '').trim().toLowerCase() + '|' +
    String(instituicao || 'Outro').trim().toLowerCase();
}

function isGenericBankName_(nomeNorm) {
  var n = String(nomeNorm || '').toLowerCase();
  return n === 'banco (generico)' || n === 'banco (genérico)' || n === 'banco generico' || n === 'banco genérico';
}

function deactivateGenericBankAccounts_(contasSheet) {
  if (!contasSheet || contasSheet.getLastRow() < 2) return;
  var data = contasSheet.getRange(2, 1, contasSheet.getLastRow() - 1, HEADERS[SHEET_CONTAS].length).getValues();
  for (var i = 0; i < data.length; i++) {
    var nome = String(data[i][1] || '').toLowerCase();
    if (isGenericBankName_(nome)) {
      contasSheet.getRange(i + 2, 5).setValue(false); // Ativo = false
      contasSheet.getRange(i + 2, 8).setValue(nowISO()); // AtualizadoEm
    }
  }
}

// ── Garantir que aba existe com cabeçalhos ──────────────────────

function ensureSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var h = HEADERS[name];
    if (h && h.length) {
      sheet.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    Logger.log('Aba criada: ' + name);
  }
  return sheet;
}

function ensureAllSheets() {
  var ss = getDbSpreadsheet();
  ensureSheet(ss, SHEET_LANCAMENTOS);
  ensureSheet(ss, SHEET_CATEGORIAS);
  ensureSheet(ss, SHEET_CONFIG);
  ensureSheet(ss, SHEET_RESERVAS);
  ensureSheet(ss, SHEET_PARCELAS_CONFIG);
  ensureSheet(ss, SHEET_CONTAS);

  // ── Migração: garantir que a coluna Instituicao exista em CONTAS ──
  var contasSheet = ss.getSheetByName(SHEET_CONTAS);
  if (contasSheet && contasSheet.getLastColumn() > 0) {
    var currentHeaders = contasSheet.getRange(1, 1, 1, contasSheet.getLastColumn()).getValues()[0];
    var expectedHeaders = HEADERS[SHEET_CONTAS];
    // Se o número de colunas atual é menor que o esperado, adicionar as que faltam
    if (currentHeaders.length < expectedHeaders.length) {
      for (var hi = currentHeaders.length; hi < expectedHeaders.length; hi++) {
        contasSheet.getRange(1, hi + 1).setValue(expectedHeaders[hi]).setFontWeight('bold');
      }
      Logger.log('Migração CONTAS: adicionadas colunas faltantes até coluna ' + expectedHeaders.length);
    }
  }

  // Migração: garantir OwnerUserId em CATEGORIAS
  var catSheet = ss.getSheetByName(SHEET_CATEGORIAS);
  if (catSheet && catSheet.getLastColumn() > 0) {
    var catHeaders = catSheet.getRange(1, 1, 1, catSheet.getLastColumn()).getValues()[0];
    var expectedCatHeaders = HEADERS[SHEET_CATEGORIAS];
    if (catHeaders.length < expectedCatHeaders.length) {
      for (var ci = catHeaders.length; ci < expectedCatHeaders.length; ci++) {
        catSheet.getRange(1, ci + 1).setValue(expectedCatHeaders[ci]).setFontWeight('bold');
      }
      Logger.log('Migração CATEGORIAS: adicionada coluna OwnerUserId');
    }
  }

  // Remover Sheet1 padrão se existir e estiver vazia
  try {
    var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
    if (def && def.getLastRow() <= 1 && ss.getSheets().length > 1) {
      ss.deleteSheet(def);
    }
  } catch (e) { }
}

// ── Helpers genéricos de leitura/escrita ─────────────────────────

function getSheetData(sheetName) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function getSheetDataAsObjects(sheetName) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = HEADERS[sheetName];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return data.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow(sheetName, values) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ensureSheet(ss, sheetName);
  }
  sheet.appendRow(values);
}

function findRowIndex(sheetName, colIndex, value) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return -1;
  var col = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]) === String(value)) return i + 2; // row number (1-indexed)
  }
  return -1;
}

function updateRow(sheetName, rowNum, values) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ensureSheet(ss, sheetName);
  }
  sheet.getRange(rowNum, 1, 1, values.length).setValues([values]);
}

function deleteRow(sheetName, rowNum) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  sheet.deleteRow(rowNum);
}

function clearSheetData(sheetName) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
}

// ── Cache helpers ────────────────────────────────────────────────

var CACHE_TTL = 300; // 5 minutos

function getCached(key) {
  try {
    var cache = CacheService.getScriptCache();
    var val = cache.get(key);
    return val ? JSON.parse(val) : null;
  } catch (e) { return null; }
}

function setCached(key, data) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    if (json.length < 100000) { // limite do CacheService
      cache.put(key, json, CACHE_TTL);
    }
  } catch (e) { Logger.log('Cache set error: ' + e.message); }
}

function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(key);
  } catch (e) { }
}

function invalidateAllCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.removeAll(['categories', 'config', 'categories_all']);
  } catch (e) { }
}

// ── Categorias com cache ─────────────────────────────────────────

function getCategoriesCached(includeInactive) {
  var hasAuthUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID);
  var authKey = hasAuthUser ? String(AUTH_EXECUTION_USER_ID) : 'anon';
  var key = (includeInactive ? 'categories_all' : 'categories') + '_' + authKey;
  if (!hasAuthUser) {
    var cached = getCached(key);
    if (cached) return cached;
  }

  var result = [];
  var seen = {};

  // Categorias públicas: sempre da base global (categorias atuais viram públicas).
  var publicRaw = getPublicCategoriesRaw_();
  publicRaw.forEach(function (c) {
    var mapped = mapCategoryRow_(c);
    mapped.isPublic = true;
    if (!seen[mapped.id]) {
      seen[mapped.id] = true;
      result.push(mapped);
    }
  });

  // Categorias privadas: criadas pelo usuário atual.
  var privateRaw = getPrivateCategoriesRaw_();
  privateRaw.forEach(function (c) {
    var mapped = mapCategoryRow_(c);
    mapped.isPublic = false;
    if (!seen[mapped.id]) {
      seen[mapped.id] = true;
      result.push(mapped);
    }
  });

  result = result.map(function (c) {
    if (!c.id) c.id = generateUUID();
    return c;
  });

  result.sort(function (a, b) { return a.order - b.order; });

  if (!includeInactive) {
    result = result.filter(function (c) { return c.active; });
  }

  if (!hasAuthUser) {
    setCached(key, result);
  }
  return result;
}

function mapCategoryRow_(c) {
  return {
    id: c.ID,
    name: c.Nome,
    applicableType: c.TipoAplicavel,
    emoji: c.Emoji,
    colorHex: c.CorHex,
    order: Number(c.Ordem) || 0,
    active: c.Ativo === true || c.Ativo === 'TRUE' || c.Ativo === 1,
    createdAt: c.CriadoEm,
    updatedAt: c.AtualizadoEm,
    ownerUserId: String(c.OwnerUserId || '')
  };
}

function getPublicCategoriesRaw_() {
  var ss = getGlobalDbSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_CATEGORIAS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return data.map(function (row) {
    return {
      ID: row[headers.indexOf('ID')],
      Nome: row[headers.indexOf('Nome')],
      TipoAplicavel: row[headers.indexOf('TipoAplicavel')],
      Emoji: row[headers.indexOf('Emoji')],
      CorHex: row[headers.indexOf('CorHex')],
      Ordem: row[headers.indexOf('Ordem')],
      Ativo: row[headers.indexOf('Ativo')],
      CriadoEm: row[headers.indexOf('CriadoEm')],
      AtualizadoEm: row[headers.indexOf('AtualizadoEm')],
      OwnerUserId: headers.indexOf('OwnerUserId') >= 0 ? row[headers.indexOf('OwnerUserId')] : ''
    };
  }).filter(function (c) {
    return String(c.OwnerUserId || '').trim() === '';
  });
}

function getPrivateCategoriesRaw_() {
  if (!(typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID)) {
    return [];
  }
  var all = getSheetDataAsObjects(SHEET_CATEGORIAS);
  var ownerId = String(AUTH_EXECUTION_USER_ID);
  return all.filter(function (c) {
    return String(c.OwnerUserId || '') === ownerId;
  });
}

// ── Config com cache ─────────────────────────────────────────────

function getConfigCached() {
  var cached = getCached('config');
  if (cached) return cached;

  var data = getSheetData(SHEET_CONFIG);
  var config = {};
  data.forEach(function (row) {
    if (row[0]) config[String(row[0])] = String(row[1]);
  });

  setCached('config', config);
  return config;
}

function getConfigValue(key) {
  var config = getConfigCached();
  return config[key] || '';
}

function setConfigValue(key, value) {
  var ss = getDbSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  var rowIdx = findRowIndex(SHEET_CONFIG, 1, key);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 2).setValue(value);
  } else {
    sheet.appendRow([key, value]);
  }
  invalidateCache('config');
}

// ── Categorias padrão ────────────────────────────────────────────

function getDefaultCategories() {
  var now = new Date().toISOString();
  var cats = [
    { name: 'Combustível', type: 'DESPESA', emoji: '⛽', color: '#FF6B6B', order: 1 },
    { name: 'Alimentação', type: 'DESPESA', emoji: '🍽️', color: '#FFD93D', order: 2 },
    { name: 'Manutenção', type: 'DESPESA', emoji: '🧰', color: '#6BCB77', order: 3 },
    { name: 'Pedágio', type: 'DESPESA', emoji: '🛣️', color: '#4D96FF', order: 4 },
    { name: 'Estacionamento', type: 'DESPESA', emoji: '🅿️', color: '#845EC2', order: 5 },
    { name: 'Seguro', type: 'DESPESA', emoji: '🛡️', color: '#00C9A7', order: 6 },
    { name: 'Parcela do carro', type: 'DESPESA', emoji: '🚗', color: '#C34A36', order: 7 },
    { name: 'Outros', type: 'DESPESA', emoji: '📌', color: '#A0A0A0', order: 8 },
    { name: 'Ganhos', type: 'RECEITA', emoji: '💰', color: '#2ECC71', order: 9 },
    { name: 'Extras', type: 'RECEITA', emoji: '🎁', color: '#9B59B6', order: 10 },
    { name: 'Despesa Fixa', type: 'DESPESA', emoji: '📅', color: '#FF9F43', order: 11 }
    // Removida categoria 'Reserva' - transferências não usam categoria
  ];

  var manutId = '';
  var rows = cats.map(function (c) {
    var id = Utilities.getUuid();
    if (c.name === 'Manutenção') manutId = id;
    return [id, c.name, c.type, c.emoji, c.color, c.order, true, now, now];
  });

  return { rows: rows, maintenanceCategoryId: manutId };
}

// ── Config padrão ────────────────────────────────────────────────

function getDefaultConfig(maintenanceCategoryId) {
  return [
    ['RESERVA_MANUTENCAO_MENSAL', '500'],
    ['MANUTENCAO_CATEGORY_ID', maintenanceCategoryId || ''],
    // Removido RESERVA_CATEGORY_ID - transferências não usam categoria
    ['META_RESERVA_MANUTENCAO', '500'],
    ['TIMEZONE', 'America/Sao_Paulo'],
    ['MOEDA', 'BRL'],
    ['THEME_DEFAULT', 'dark'],
    ['ALLOWLIST_EMAILS', '']
  ];
}

// ── Recalcular reservas ──────────────────────────────────────────

function normalizeForMaintenance_(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function isMaintenanceExpenseTx_(tx, maintenanceCategoryId) {
  if (!tx || String(tx.Tipo) !== 'DESPESA') return false;
  var catId = String(tx.CategoriaId || '');
  var cfgId = String(maintenanceCategoryId || '');
  if (cfgId && catId === cfgId) return true;
  var snapshot = normalizeForMaintenance_(tx.CategoriaNomeSnapshot);
  return snapshot === 'manutencao' || snapshot.indexOf('manutencao') !== -1;
}

function recalculateReserves() {
  var config = getConfigCached();
  var manutCatId = config['MANUTENCAO_CATEGORY_ID'] || '';
  var tz = config['TIMEZONE'] || 'America/Sao_Paulo';

  // Obter ID da conta Investimentos
  var accounts = getAccountsCached(true);
  var investimentosId = (accounts.find(function (a) { return a.nome === 'Investimentos'; }) || {}).id || 'reserva';
  // Normalizar para comparação
  investimentosId = String(investimentosId).trim().toLowerCase();

  var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);

  // Agrupar transferências para RESERVA e usos (DESPESA Manutenção) por mês
  var depositsByMonth = {};
  var usageByMonth = {};

  allTx.forEach(function (tx) {
    // Usar helper local se não existir, ou direto
    var d = tx.Data;
    var monthKey = '';
    if (d instanceof Date) {
      monthKey = Utilities.formatDate(d, tz, 'yyyy-MM');
    } else {
      var s = String(d);
      if (s.match(/^\d{4}-\d{2}/)) monthKey = s.substring(0, 7);
      else if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) monthKey = s.substring(6, 10) + '-' + s.substring(3, 5);
    }

    if (!monthKey) return;

    var val = parseFloat(tx.Valor) || 0;
    var fromWallet = String(tx.CarteiraOrigem || '').trim().toLowerCase();
    var toWallet = String(tx.CarteiraDestino || '').trim().toLowerCase();

    // Depósitos na reserva = TRANSFER com destino = Investimentos (ou 'reserva' legado)
    if (tx.Tipo === 'TRANSFER' && (toWallet === investimentosId || toWallet.indexOf('reserva') !== -1)) {
      depositsByMonth[monthKey] = (depositsByMonth[monthKey] || 0) + val;
    }

    // Retiradas da reserva = TRANSFER com origem = Investimentos
    if (tx.Tipo === 'TRANSFER' && (fromWallet === investimentosId || fromWallet.indexOf('reserva') !== -1)) {
      depositsByMonth[monthKey] = (depositsByMonth[monthKey] || 0) - val;
    }

    // Usos da reserva = DESPESA com categoria Manutenção (independente da conta, assumimos que sai da reserva conceitualmente)
    // OU se a despesa saiu da conta Investimentos
    var isManutencao = isMaintenanceExpenseTx_(tx, manutCatId);
    var isFromReserva = (fromWallet === investimentosId || fromWallet.indexOf('reserva') !== -1);

    if (tx.Tipo === 'DESPESA' && (isManutencao || isFromReserva)) {
      usageByMonth[monthKey] = (usageByMonth[monthKey] || 0) + val;
    }
  });

  // Determinar range de meses
  var now = new Date();
  var currentMonth = Utilities.formatDate(now, tz, 'yyyy-MM');
  var allMonthKeys = Object.keys(depositsByMonth).concat(Object.keys(usageByMonth));
  allMonthKeys.push(currentMonth);

  var firstTxDate = null;
  allTx.forEach(function (tx) {
    var d = tx.Data;
    if (d instanceof Date && (!firstTxDate || d < firstTxDate)) {
      firstTxDate = d;
    }
  });

  var startMonth = currentMonth;
  if (firstTxDate) {
    startMonth = Utilities.formatDate(firstTxDate, tz, 'yyyy-MM');
  }

  // Gerar lista de meses
  var months = [];
  var cursor = startMonth;
  while (cursor <= currentMonth) {
    months.push(cursor);
    var parts = cursor.split('-');
    var y = parseInt(parts[0]);
    var m = parseInt(parts[1]);
    m++;
    if (m > 12) { m = 1; y++; }
    cursor = y + '-' + (m < 10 ? '0' + m : '' + m);
  }

  // Calcular reservas mês a mês com transações reais
  var prevBalance = 0;
  var reserveRows = months.map(function (month) {
    var deposits = depositsByMonth[month] || 0;
    var accumulated = prevBalance + deposits;
    var usage = usageByMonth[month] || 0;
    var balance = accumulated - usage;
    prevBalance = balance;
    return [month, deposits, accumulated, usage, balance];
  });

  // Reescrever aba RESERVAS
  clearSheetData(SHEET_RESERVAS);
  if (reserveRows.length > 0) {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_RESERVAS);
    sheet.getRange(2, 1, reserveRows.length, 5).setValues(reserveRows);
  }

  return reserveRows;
}

// ── Obter reserva de um mês específico ───────────────────────────

function getReserveForMonth(month) {
  var data = getSheetDataAsObjects(SHEET_RESERVAS);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i].Mes) === month) {
      return {
        month: data[i].Mes,
        predicted: parseFloat(data[i].ReservaManutencaoPrevista) || 0,
        accumulated: parseFloat(data[i].ReservaManutencaoAcumulada) || 0,
        usage: parseFloat(data[i].UsosManutencaoNoMes) || 0,
        balance: parseFloat(data[i].SaldoReserva) || 0
      };
    }
  }
  return { month: month, predicted: 0, accumulated: 0, usage: 0, balance: 0 };
}

// ── Contas (Accounts) com cache ──────────────────────────────────

function getAccountsCached(includeInactive) {
  var hasAuthUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID);
  var scopeUser = hasAuthUser ? String(AUTH_EXECUTION_USER_ID) : 'public';
  var key = (includeInactive ? 'accounts_all' : 'accounts') + '_' + scopeUser;
  var cached = hasAuthUser ? null : getCached(key);
  if (cached) return cached;

  var all = getSheetDataAsObjects(SHEET_CONTAS);
  var result = all.map(function (a) {
    return {
      id: a.ID,
      nome: a.Nome,
      tipo: a.Tipo,
      saldoInicial: parseFloat(a.SaldoInicial) || 0,
      ativo: a.Ativo === true || a.Ativo === 'TRUE' || a.Ativo === 1,
      ordem: Number(a.Ordem) || 0,
      criadoEm: a.CriadoEm,
      atualizadoEm: a.AtualizadoEm
    };
  });

  result.sort(function (a, b) { return a.ordem - b.ordem; });

  if (!includeInactive) {
    result = result.filter(function (a) { return a.ativo; });
  }

  if (!hasAuthUser) {
    setCached(key, result);
  }
  return result;
}

// ── Criar contas padrão ──────────────────────────────────────────

function createDefaultAccounts() {
  var now = new Date().toISOString();
  var accounts = [
    { nome: 'Carteira', tipo: 'Dinheiro', saldoInicial: 0, ordem: 1, instituicao: 'Carteira' },
    { nome: 'Investimentos', tipo: 'Investimentos', saldoInicial: 0, ordem: 2, instituicao: 'Outro' }
  ];

  var rows = accounts.map(function (a) {
    var id = Utilities.getUuid();
    return [id, a.nome, a.tipo, a.saldoInicial, true, a.ordem, now, now, a.instituicao];
  });

  return rows;
}

// ── Validar nome de conta único ──────────────────────────────────

function validateAccountName(name, excludeId) {
  var all = getSheetDataAsObjects(SHEET_CONTAS);
  var nameLower = String(name).toLowerCase().trim();

  for (var i = 0; i < all.length; i++) {
    var existingName = String(all[i].Nome).toLowerCase().trim();
    var existingId = String(all[i].ID);

    if (existingName === nameLower && existingId !== excludeId) {
      return false; // Nome já existe
    }
  }
  return true; // Nome disponível
}

// ── Calcular saldo de conta ──────────────────────────────────────

function calculateAccountBalance(accountId) {
  var account = null;
  var accounts = getSheetDataAsObjects(SHEET_CONTAS);

  for (var i = 0; i < accounts.length; i++) {
    if (String(accounts[i].ID) === String(accountId)) {
      account = accounts[i];
      break;
    }
  }

  if (!account) return 0;

  var saldoInicial = parseFloat(account.SaldoInicial) || 0;
  var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);

  var balance = saldoInicial;

  allTx.forEach(function (tx) {
    // IMPORTANTE: Apenas transações PAGAS afetam o saldo
    // Se StatusParcela não existir (legado), assume PAGA
    var status = tx.StatusParcela || 'PAGA';
    if (status !== 'PAGA' && status !== 'COMPENSADO') return;

    var valor = parseFloat(tx.Valor) || 0;
    var tipo = String(tx.Tipo);
    var contaOrigem = String(tx.CarteiraOrigem || '');
    var contaDestino = String(tx.CarteiraDestino || '');

    // RECEITA: adiciona ao saldo da conta destino
    if (tipo === 'RECEITA' && contaDestino === accountId) {
      balance += valor;
    }

    // DESPESA: subtrai do saldo da conta origem
    if (tipo === 'DESPESA' && contaOrigem === accountId) {
      balance -= valor;
    }

    // TRANSFERENCIA: subtrai da origem, adiciona no destino
    if (tipo === 'TRANSFERENCIA' || tipo === 'TRANSFER') {
      if (contaOrigem === accountId) {
        balance -= valor;
      }
      if (contaDestino === accountId) {
        balance += valor;
      }
    }
  });

  return balance;
}

// ── Verificar se conta tem transações ────────────────────────────

function accountHasTransactions(accountId) {
  var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);

  for (var i = 0; i < allTx.length; i++) {
    var tx = allTx[i];
    var contaOrigem = String(tx.CarteiraOrigem || '');
    var contaDestino = String(tx.CarteiraDestino || '');

    if (contaOrigem === accountId || contaDestino === accountId) {
      return true;
    }
  }

  return false;
}

