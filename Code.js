function doGet(e) {
  try {
    // Servir manifest.json se solicitado
    if (e.parameter && e.parameter.manifest === 'true') {
      return serveManifest();
    }

    // Servir logos dos bancos se solicitado
    if (e.parameter && e.parameter.logo) {
      return serveBankLogo(e.parameter.logo);
    }

    // Tela de Cartões de Crédito: ?v=cartoes
    if (e.parameter && e.parameter.v === 'cartoes') {
      return HtmlService.createTemplateFromFile('cartoes')
        .evaluate()
        .setTitle('Cartões de Crédito — Sistema Matheus')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var isAppView = e.parameter && e.parameter.app === '1';
    var template = HtmlService.createTemplateFromFile(isAppView ? 'app_main' : 'index');
    return template.evaluate()
      .setTitle('App Financeiro do Motorista')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    var msg = err.message || String(err);
    var stack = err.stack || '';
    var needsSetup = msg.indexOf('LANCAMENTOS') !== -1 ||
                     msg.indexOf('CATEGORIAS') !== -1 ||
                     msg.indexOf('CONFIG') !== -1 ||
                     msg.indexOf('Cannot read prop') !== -1;

    var html = '<html><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<style>body{font-family:system-ui,sans-serif;background:#0f1117;color:#e8eaf0;' +
      'max-width:600px;margin:40px auto;padding:20px;line-height:1.6}' +
      'h2{color:#FF6B6B}pre{background:#1a1d27;padding:16px;border-radius:8px;' +
      'overflow-x:auto;font-size:13px;border:1px solid #2a2e3d}' +
      '.hint{background:#222633;padding:16px;border-radius:8px;margin-top:20px;' +
      'border-left:4px solid #6C63FF}</style></head><body>' +
      '<h2>Erro ao carregar o App</h2>' +
      '<p><strong>Mensagem:</strong> ' + msg + '</p>' +
      '<pre>' + stack + '</pre>';

    if (needsSetup) {
      html += '<div class="hint"><strong>Dica:</strong> Execute a função ' +
        '<code>setup()</code> no editor do Apps Script ' +
        '(selecione <em>setup</em> no dropdown e clique em ▶ Executar).</div>';
    }

    html += '</body></html>';

    return HtmlService.createHtmlOutput(html)
      .setTitle('Erro — App Financeiro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// Servir logo de banco diretamente do Google Drive
function serveBankLogo(bankKey) {
  // Validar parâmetro
  if (!bankKey || typeof bankKey !== 'string') {
    Logger.log('Erro: bankKey inválido: ' + bankKey);
    return ContentService.createTextOutput('Parâmetro logo inválido')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var logoIds = {
    'nubank': '1valKbUhGpUjclLRrjEDCwGzcd3V129vu',
    'inter': '1cWiNAbpXYJWiKlnHFTUQ6chMQeAWnKx6',
    'itau': '14ZtTHa15L33stR9weecOYQoJNBUiCKPX',
    'bradesco': '1a93spm-9qVkFcR1v_CSHQlqgnS8KRuhK',
    'santander': '1izWcguwasT4SaHKsjdpzymav51xx30IN',
    'bb': '1BegWGKuJARynZLuqlVXLDL_E5ALzVk_F',
    'caixa': '1t_G0rz0SbSX2NLbuL4omkjVH-N3ML780',
    'c6': '1u2f8IvkbtIrfMJQmY2SAZYOOOdUKxzS2',
    'btg': '11IyuODdg1dYIpIM1qlab49ZGQdoeqvXJ',
    'mercadopago': '17UiiW6GP-Tefg02R64q0rePYyJfIBmWT',
    'picpay': '1MQHKt1tfNaCG4OKJj9pfyyYC4pgVVxHw'
  };

  var key = bankKey.toLowerCase().trim();
  var fileId = logoIds[key];

  Logger.log('Requisição logo: ' + key + ' | ID: ' + fileId);

  if (!fileId) {
    return ContentService.createTextOutput('Logo não encontrada para: ' + key)
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var contentType = blob.getContentType();

    Logger.log('Logo servida com sucesso: ' + key + ' | Tipo: ' + contentType);

    // Para SVG, retornar como texto
    if (contentType.indexOf('svg') !== -1) {
      return ContentService.createTextOutput(blob.getDataAsString())
        .setMimeType('image/svg+xml');
    }

    // Para outros formatos (PNG, JPG), converter para base64 e embutir
    var base64 = Utilities.base64Encode(blob.getBytes());
    var dataUrl = 'data:' + contentType + ';base64,' + base64;

    // Redirecionar para a imagem base64 via HTML
    var html = '<html><body><img src="' + dataUrl + '" style="max-width:100%;height:auto;" /></body></html>';

    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (e) {
    Logger.log('Erro ao servir logo ' + key + ': ' + e.toString());
    return ContentService.createTextOutput('Erro: ' + e.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function serveManifest() {
  var manifest = {
    "name": "App Financeiro do Motorista",
    "short_name": "Financeiro",
    "description": "Controle financeiro profissional para motoristas de aplicativo",
    "start_url": ScriptApp.getService().getUrl(),
    "display": "standalone",
    "background_color": "#0f1117",
    "theme_color": "#6C63FF",
    "orientation": "portrait-primary",
    "icons": [
      {
        "src": "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect fill='%236C63FF' width='100' height='100' rx='20'/><text x='50' y='70' font-size='60' text-anchor='middle' fill='white'>💰</text></svg>",
        "sizes": "192x192",
        "type": "image/svg+xml",
        "purpose": "any maskable"
      },
      {
        "src": "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect fill='%236C63FF' width='100' height='100' rx='20'/><text x='50' y='70' font-size='60' text-anchor='middle' fill='white'>💰</text></svg>",
        "sizes": "512x512",
        "type": "image/svg+xml",
        "purpose": "any maskable"
      }
    ]
  };
  
  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Include para HTML partials ───────────────────────────────────

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Auth e setup ─────────────────────────────────────────────────

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function checkAuth() {
  var email = getCurrentUserEmail();
  var allowlist = getConfigValue('ALLOWLIST_EMAILS');
  if (!allowlist) return true; // sem allowlist = todos podem
  var emails = allowlist.split(',').map(function(e) { return e.trim().toLowerCase(); });
  return emails.indexOf(email.toLowerCase()) !== -1;
}

function setup() {
  ensureAllSheets();
  
  var ss = getDbSpreadsheet();
  var catSheet = ss.getSheetByName(SHEET_CATEGORIAS);
  if (catSheet.getLastRow() < 2) {
    var defaults = getDefaultCategories();
    defaults.rows.forEach(function(row) {
      catSheet.appendRow(row);
    });
    
    var cfgSheet = ss.getSheetByName(SHEET_CONFIG);
    var defaultCfg = getDefaultConfig(defaults.maintenanceCategoryId);
    defaultCfg.forEach(function(row) {
      cfgSheet.appendRow(row);
    });
    
    Logger.log('Setup completo! Categorias e config criadas.');
  }
  
  return 'Setup concluído! Recarregue o app.';
}

// ── Responses ────────────────────────────────────────────────────

function okResponse(data) {
  return { ok: true, data: data };
}

function errorResponse(message, code) {
  return { ok: false, error: { message: message, code: code } };
}

// ── Helpers de data/hora ─────────────────────────────────────────

function nowISO() {
  return new Date().toISOString();
}

function dateToYMD(d) {
  if (!d) return '';
  if (!(d instanceof Date)) d = new Date(d);
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + dd;
}

function parseDateInput(str) {
  if (!str) return new Date();
  if (str instanceof Date) return str;
  var parts = String(str).split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  return new Date(str);
}

function generateUUID() {
  return Utilities.getUuid();
}

// ── Tema ─────────────────────────────────────────────────────────

function getTheme() {
  var userProps = PropertiesService.getUserProperties();
  var theme = userProps.getProperty('THEME') || 'dark';
  return theme;
}

function setTheme(theme) {
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty('THEME', theme);
  return okResponse({ theme: theme });
}

// ── Timezone ─────────────────────────────────────────────────────

function getTZ() {
  return getConfigValue('TIMEZONE') || 'America/Sao_Paulo';
}

// Normaliza id de carteira (tolerante a maiúsculas e espaços)
function normalizeWalletId(id) {
  return String(id || '').trim().toLowerCase();
}

// Extrai chave YYYY-MM a partir de Date ou string em formatos comuns
function extractMonthKey(value, tz) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz || getTZ(), 'yyyy-MM');
  }
  var s = String(value);
  if (s.match(/^\d{4}-\d{2}/)) return s.substring(0, 7);
  if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) return s.substring(6, 10) + '-' + s.substring(3, 5);
  return '';
}

// ── Bank Logos Management ────────────────────────────────────────────

/**
 * Baixa logos de bancos do GitHub e salva no Google Drive com acesso público
 * Execute esta função UMA VEZ para fazer o upload das logos
 */
function uploadBankLogosToGoogleDrive() {
  // URLs alternativas de logos usando diferentes fontes confiáveis
  var bankLogos = {
    'nubank': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Nu%20Pagamentos%20S.A%20(Nubank)/nubank-3.svg',
    'inter': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Banco%20Inter%20S.A/logo%20banco%20inter.svg',
    'itau': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Ita%C3%BA%20Unibanco%20S.A/itau-5.svg',
    'bradesco': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Banco%20Bradesco%20S.A/bradesco-2.svg',
    'santander': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Banco%20Santander%20(Brasil)%20S.A/santander-5.svg',
    'banco-do-brasil': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Banco%20do%20Brasil%20S.A/banco-do-brasil-4.svg',
    'caixa': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Caixa%20Econ%C3%B4mica%20Federal/caixa-economica-federal-logo.svg',
    'c6-bank': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Banco%20C6%20S.A/c6-bank-seeklogo.com.svg',
    'btg-pactual': 'https://raw.githubusercontent.com/Tgentil/Bancos-em-SVG/main/Banco%20BTG%20Pactual%20S.A/btg-pactual-logo-1.svg',
    'mercado-pago': 'https://logodownload.org/wp-content/uploads/2020/02/mercado-pago-logo-3.svg',
    'picpay': 'https://seeklogo.com/images/P/picpay-logo-EE83996357-seeklogo.com.png'
  };

  // Criar ou obter pasta "BankLogos" na raiz do Drive
  var folders = DriveApp.getFoldersByName('BankLogos');
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
    Logger.log('Pasta BankLogos já existe');
  } else {
    folder = DriveApp.createFolder('BankLogos');
    Logger.log('Pasta BankLogos criada');
  }

  var results = {};
  var errors = [];

  // Baixar e salvar cada logo
  for (var bankKey in bankLogos) {
    try {
      var url = bankLogos[bankKey];
      Logger.log('Baixando: ' + bankKey + ' de ' + url);

      // Baixar imagem com timeout maior
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true
      });

      var responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        throw new Error('HTTP ' + responseCode + ' - Arquivo não encontrado ou inacessível');
      }

      var blob = response.getBlob();

      // Determinar extensão correta
      var contentType = blob.getContentType();
      var extension = '.svg';
      if (contentType && contentType.indexOf('png') !== -1) {
        extension = '.png';
      }
      var fileName = bankKey + extension;

      // Verificar se arquivo já existe
      var existingFiles = folder.getFilesByName(fileName);
      var file;

      if (existingFiles.hasNext()) {
        // Atualizar arquivo existente
        file = existingFiles.next();
        file.setContent(blob);
        Logger.log('Atualizado: ' + fileName);
      } else {
        // Criar novo arquivo
        file = folder.createFile(blob.setName(fileName));
        Logger.log('Criado: ' + fileName);
      }

      // Tornar arquivo público
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // Obter URL pública
      var publicUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      results[bankKey] = publicUrl;

      Logger.log('✓ ' + bankKey + ': ' + publicUrl);

    } catch (e) {
      var errorMsg = bankKey + ': ' + e.toString();
      errors.push(errorMsg);
      Logger.log('✗ Erro em ' + errorMsg);
    }
  }

  // Resumo
  Logger.log('\n=== RESUMO ===');
  Logger.log('Sucessos: ' + Object.keys(results).length);
  Logger.log('Erros: ' + errors.length);

  if (errors.length > 0) {
    Logger.log('\nErros encontrados:');
    errors.forEach(function(err) {
      Logger.log('  - ' + err);
    });
  }

  Logger.log('\n=== URLs PÚBLICAS ===');
  Logger.log(JSON.stringify(results, null, 2));

  return okResponse({
    success: Object.keys(results).length,
    failed: errors.length,
    urls: results,
    errors: errors
  });
}

/**
 * Retorna as URLs das logos já salvas no Drive
 * Usa o ID da pasta compartilhada: 1LozJoqsPgZjmJGqWB3t1Qc1J_TEru3Gq
 */
function getBankLogosUrls() {
  try {
    // ID da pasta compartilhada
    var folderId = '1LozJoqsPgZjmJGqWB3t1Qc1J_TEru3Gq';
    var folder = DriveApp.getFolderById(folderId);

    var files = folder.getFiles();
    var urls = {};
    var mapping = {};

    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName().toLowerCase();
      var fileId = file.getId();

      // Tornar arquivo público se ainda não for
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        Logger.log('Aviso: não foi possível alterar permissão de ' + fileName);
      }

      var publicUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

      // Mapear nomes de arquivos para chaves do sistema
      if (fileName.indexOf('nubank') !== -1) {
        urls['nubank'] = publicUrl;
        mapping['nubank'] = fileName;
      } else if (fileName.indexOf('inter') !== -1) {
        urls['inter'] = publicUrl;
        mapping['inter'] = fileName;
      } else if (fileName.indexOf('itau') !== -1 || fileName.indexOf('itaú') !== -1) {
        urls['itau'] = publicUrl;
        mapping['itau'] = fileName;
      } else if (fileName.indexOf('bradesco') !== -1) {
        urls['bradesco'] = publicUrl;
        mapping['bradesco'] = fileName;
      } else if (fileName.indexOf('santander') !== -1) {
        urls['santander'] = publicUrl;
        mapping['santander'] = fileName;
      } else if (fileName.indexOf('brasil') !== -1 || fileName.indexOf('bb') !== -1) {
        urls['banco-do-brasil'] = publicUrl;
        mapping['banco-do-brasil'] = fileName;
      } else if (fileName.indexOf('caixa') !== -1) {
        urls['caixa'] = publicUrl;
        mapping['caixa'] = fileName;
      } else if (fileName.indexOf('c6') !== -1) {
        urls['c6-bank'] = publicUrl;
        mapping['c6-bank'] = fileName;
      } else if (fileName.indexOf('btg') !== -1) {
        urls['btg-pactual'] = publicUrl;
        mapping['btg-pactual'] = fileName;
      } else if (fileName.indexOf('mercado') !== -1 || fileName.indexOf('pago') !== -1) {
        urls['mercado-pago'] = publicUrl;
        mapping['mercado-pago'] = fileName;
      } else if (fileName.indexOf('picpay') !== -1 || fileName.indexOf('pic') !== -1) {
        urls['picpay'] = publicUrl;
        mapping['picpay'] = fileName;
      }
    }

    Logger.log('=== LOGOS ENCONTRADAS ===');
    Logger.log(JSON.stringify(mapping, null, 2));
    Logger.log('\n=== URLs PÚBLICAS ===');
    Logger.log(JSON.stringify(urls, null, 2));

    return okResponse({
      urls: urls,
      mapping: mapping,
      total: Object.keys(urls).length
    });

  } catch (e) {
    return errorResponse('Erro ao acessar pasta do Drive: ' + e.toString(), 'DRIVE_ERROR');
  }
}

/**
 * Função auxiliar para listar todos os arquivos da pasta
 * Use para ver os nomes exatos dos arquivos
 */
function listBankLogosFiles() {
  var folderId = '1LozJoqsPgZjmJGqWB3t1Qc1J_TEru3Gq';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var fileList = [];

  while (files.hasNext()) {
    var file = files.next();
    fileList.push({
      name: file.getName(),
      id: file.getId(),
      url: 'https://drive.google.com/uc?export=view&id=' + file.getId()
    });
  }

  Logger.log('=== ARQUIVOS NA PASTA ===');
  fileList.forEach(function(f) {
    Logger.log(f.name + ' → ' + f.url);
  });

  return okResponse(fileList);
}

/**
 * SOLUÇÃO DEFINITIVA: Converte todas as logos para base64
 * Execute esta função UMA VEZ e cole o resultado no BANK_REPOSITORY
 */
function convertLogosToBase64() {
  var logoIds = {
    'Nubank': '1valKbUhGpUjclLRrjEDCwGzcd3V129vu',
    'Inter': '1cWiNAbpXYJWiKlnHFTUQ6chMQeAWnKx6',
    'Itau': '14ZtTHa15L33stR9weecOYQoJNBUiCKPX',
    'Bradesco': '1a93spm-9qVkFcR1v_CSHQlqgnS8KRuhK',
    'Santander': '1izWcguwasT4SaHKsjdpzymav51xx30IN',
    'BB': '1BegWGKuJARynZLuqlVXLDL_E5ALzVk_F',
    'Caixa': '1t_G0rz0SbSX2NLbuL4omkjVH-N3ML780',
    'C6': '1u2f8IvkbtIrfMJQmY2SAZYOOOdUKxzS2',
    'BTG': '11IyuODdg1dYIpIM1qlab49ZGQdoeqvXJ',
    'MercadoPago': '17UiiW6GP-Tefg02R64q0rePYyJfIBmWT',
    'PicPay': '1MQHKt1tfNaCG4OKJj9pfyyYC4pgVVxHw'
  };

  Logger.log('=== CONVERTENDO LOGOS PARA BASE64 ===\n');
  Logger.log('COLE ESTE CÓDIGO NO BANK_REPOSITORY:\n');

  for (var bank in logoIds) {
    try {
      var fileId = logoIds[bank];
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      var mimeType = blob.getContentType();
      var dataUrl = 'data:' + mimeType + ';base64,' + base64;

      Logger.log("BANK_REPOSITORY['" + bank + "'].logo = '" + dataUrl.substring(0, 100) + "...' // " + (dataUrl.length) + " caracteres");

    } catch (e) {
      Logger.log("// ERRO em " + bank + ": " + e.toString());
    }
  }

  Logger.log('\n=== GERANDO ARQUIVO COMPLETO ===');
  Logger.log('Por questões de tamanho, use a função getBankLogosBase64() para obter as URLs');

  return okResponse({ message: 'Veja os logs' });
}

/**
 * Retorna as logos em base64 para usar via RPC
 */
function getBankLogosBase64() {
  var logoIds = {
    'Nubank': '1valKbUhGpUjclLRrjEDCwGzcd3V129vu',
    'Inter': '1cWiNAbpXYJWiKlnHFTUQ6chMQeAWnKx6',
    'Itau': '14ZtTHa15L33stR9weecOYQoJNBUiCKPX',
    'Bradesco': '1a93spm-9qVkFcR1v_CSHQlqgnS8KRuhK',
    'Santander': '1izWcguwasT4SaHKsjdpzymav51xx30IN',
    'BB': '1BegWGKuJARynZLuqlVXLDL_E5ALzVk_F',
    'Caixa': '1t_G0rz0SbSX2NLbuL4omkjVH-N3ML780',
    'C6': '1u2f8IvkbtIrfMJQmY2SAZYOOOdUKxzS2',
    'BTG': '11IyuODdg1dYIpIM1qlab49ZGQdoeqvXJ',
    'MercadoPago': '17UiiW6GP-Tefg02R64q0rePYyJfIBmWT',
    'PicPay': '1MQHKt1tfNaCG4OKJj9pfyyYC4pgVVxHw'
  };

  var result = {};

  for (var bank in logoIds) {
    try {
      var fileId = logoIds[bank];
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      var mimeType = blob.getContentType();
      result[bank] = 'data:' + mimeType + ';base64,' + base64;
    } catch (e) {
      Logger.log('Erro em ' + bank + ': ' + e);
      result[bank] = null;
    }
  }

  return okResponse(result);
}

/**
 * EXECUTE ESTA FUNÇÃO UMA VEZ para tornar todos os arquivos públicos
 * e gerar o código JavaScript com as URLs corretas
 */
function makeAllLogosPublicAndGenerateCode() {
  var folderId = '1LozJoqsPgZjmJGqWB3t1Qc1J_TEru3Gq';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var mapping = {};

  Logger.log('=== TORNANDO ARQUIVOS PÚBLICOS ===\n');

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();

    // Tornar público
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log('✓ ' + fileName + ' → PÚBLICO');
    } catch (e) {
      Logger.log('✗ ' + fileName + ' → ERRO: ' + e);
    }

    var url = 'https://drive.google.com/uc?export=view&id=' + fileId;
    var key = fileName.toLowerCase();

    // Detectar banco pelo nome
    if (key.indexOf('nubank') !== -1) mapping['Nubank'] = url;
    else if (key.indexOf('inter') !== -1) mapping['Inter'] = url;
    else if (key.indexOf('itau') !== -1 || key.indexOf('itaú') !== -1) mapping['Itau'] = url;
    else if (key.indexOf('bradesco') !== -1) mapping['Bradesco'] = url;
    else if (key.indexOf('santander') !== -1) mapping['Santander'] = url;
    else if (key.indexOf('brasil') !== -1 || key.indexOf('bb') !== -1) mapping['BB'] = url;
    else if (key.indexOf('caixa') !== -1) mapping['Caixa'] = url;
    else if (key.indexOf('c6') !== -1) mapping['C6'] = url;
    else if (key.indexOf('btg') !== -1) mapping['BTG'] = url;
    else if (key.indexOf('mercado') !== -1 || key.indexOf('pago') !== -1) mapping['MercadoPago'] = url;
    else if (key.indexOf('picpay') !== -1 || key.indexOf('pic') !== -1) mapping['PicPay'] = url;
  }

  Logger.log('\n=== CÓDIGO PARA COLAR NO app_js.html ===\n');
  Logger.log('// Cole este código no BANK_REPOSITORY:\n');

  for (var bank in mapping) {
    Logger.log("  BANK_REPOSITORY['" + bank + "'].logo = '" + mapping[bank] + "';");
  }

  Logger.log('\n=== OU COLE ESTE OBJETO COMPLETO ===\n');
  Logger.log(JSON.stringify(mapping, null, 2));

  return okResponse(mapping);
}
