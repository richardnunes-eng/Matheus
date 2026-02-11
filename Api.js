/* ================================================================
   Api.gs — Todos os endpoints RPC (google.script.run)
   App Financeiro do Motorista
   ================================================================ */

// ── getInitData ──────────────────────────────────────────────────

function getInitData() {
  try {
    var categories = getCategoriesCached(false);
    var config = getConfigCached();
    var now = new Date();
    var tz = config['TIMEZONE'] || 'America/Sao_Paulo';
    var month = parseInt(Utilities.formatDate(now, tz, 'MM'));
    var year = parseInt(Utilities.formatDate(now, tz, 'yyyy'));
    var email = '';
    try {
      email = getCurrentUserEmail() || '';
    } catch (authErr) {
      email = '';
    }
    if (!email) {
      email = 'Usuario autenticado';
    }

    // Tema efetivo
    var userTheme = '';
    try {
      userTheme = PropertiesService.getUserProperties().getProperty('THEME') || '';
    } catch (e) { }
    var theme = userTheme || config['THEME_DEFAULT'] || 'dark';

    return okResponse({
      categories: categories,
      config: config,
      month: month,
      year: year,
      email: email,
      theme: theme
    });
  } catch (e) {
    return errorResponse(e.message, 'INIT_ERROR');
  }
}

// ── getDashboard ─────────────────────────────────────────────────

function getDashboard(params) {
  try {
    var month = params.month;
    var year = params.year;
    var monthKey = year + '-' + (month < 10 ? '0' + month : '' + month);
    var tz = getTZ();

    // Obter IDs das contas principais para lógica de negócio
    var accounts = getAccountsCached(true);
    var investimentosId = (accounts.find(function (a) { return a.nome === 'Investimentos'; }) || {}).id || 'reserva';
    var carteiraId = (accounts.find(function (a) { return a.nome === 'Carteira'; }) || {}).id || 'principal';

    // Obter transações do mês
    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var monthTx = allTx.filter(function (tx) {
      var d = tx.Data;
      var mk = '';
      if (d instanceof Date) {
        mk = Utilities.formatDate(d, tz, 'yyyy-MM');
      } else {
        var s = String(d);
        if (s.match(/^\d{4}-\d{2}/)) mk = s.substring(0, 7);
        else if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) mk = s.substring(6, 10) + '-' + s.substring(3, 5);
      }
      return mk === monthKey;
    });

    var totalReceita = 0;
    var totalDespesa = 0;
    var byCategory = {};
    var config = getConfigCached();
    var manutCatId = config['MANUTENCAO_CATEGORY_ID'] || '';
    var gastoManutencao = 0;
    var transferenciasParaReserva = 0; // NOVO: rastrear transferências para reserva

    monthTx.forEach(function (tx) {
      var val = parseFloat(tx.Valor) || 0;
      // IMPORTANTE: Transferências não contam como receita/despesa
      if (tx.Tipo === 'RECEITA') {
        totalReceita += val;
      } else if (tx.Tipo === 'DESPESA') {
        totalDespesa += val;
        var catId = String(tx.CategoriaId);
        if (!byCategory[catId]) byCategory[catId] = 0;
        byCategory[catId] += val;
        if (isMaintenanceExpenseTx_(tx, manutCatId)) {
          gastoManutencao += val;
        }
      } else if (tx.Tipo === 'TRANSFER') {
        // NOVO: Se for transferência da principal para reserva, subtrair do disponível
        // Usar IDs dinâmicos
        var origem = normalizeWalletId(tx.CarteiraOrigem);
        var destino = normalizeWalletId(tx.CarteiraDestino);

        // Verifica se é transferencia de COMPOSIÇÃO DE RESERVA (Carteira -> Investimentos)
        if (origem === carteiraId && destino === investimentosId) {
          transferenciasParaReserva += val;
        }
      }
    });

    // Categorias para join
    var categories = getCategoriesCached(true);
    var catMap = {};
    categories.forEach(function (c) { catMap[c.id] = c; });

    // Despesas por categoria (detalhado)
    var despesasPorCategoria = [];
    Object.keys(byCategory).forEach(function (catId) {
      var cat = catMap[catId] || { name: 'Desconhecida', emoji: '❓', colorHex: '#999' };
      despesasPorCategoria.push({
        categoryId: catId,
        name: cat.name,
        emoji: cat.emoji,
        colorHex: cat.colorHex,
        total: byCategory[catId]
      });
    });
    despesasPorCategoria.sort(function (a, b) { return b.total - a.total; });

    // Combustível
    var gastoCombustivel = 0;
    categories.forEach(function (c) {
      if (c.name.toLowerCase() === 'combustível' && byCategory[c.id]) {
        gastoCombustivel = byCategory[c.id];
      }
    });

    // Alimentação
    var gastoAlimentacao = 0;
    categories.forEach(function (c) {
      if (c.name.toLowerCase() === 'alimentação' && byCategory[c.id]) {
        gastoAlimentacao = byCategory[c.id];
      }
    });

    // Custos Fixos (Despesa Fixa) - Nova Lógica Baseada em Categoria
    var custosFixos = 0;
    categories.forEach(function (c) {
      if ((c.name.toLowerCase() === 'despesa fixa' || c.name.toLowerCase() === 'despesas fixas') && byCategory[c.id]) {
        custosFixos = byCategory[c.id];
      }
    });
    // Legacy support (send 0)
    var custoFixoParcela = 0;
    var custoFixoSeguro = 0;

    // Reserva manutenção (recalcular para garantir dados atualizados)
    recalculateReserves();
    // Calcular reserva direto das transações (saldo acumulado até o mês atual)
    var reserva = (function () {
      var depositsByMonth = {};
      var usageByMonth = {};

      allTx.forEach(function (tx) {
        var mk = extractMonthKey(tx.Data, tz);
        if (!mk) return;

        var val = parseFloat(tx.Valor) || 0;
        var fromWallet = normalizeWalletId(tx.CarteiraOrigem);
        var toWallet = normalizeWalletId(tx.CarteiraDestino);

        if (tx.Tipo === 'TRANSFER' && toWallet === investimentosId) {
          depositsByMonth[mk] = (depositsByMonth[mk] || 0) + val;
        }
        if (tx.Tipo === 'TRANSFER' && fromWallet === investimentosId) {
          depositsByMonth[mk] = (depositsByMonth[mk] || 0) - val;
        }
        if (isMaintenanceExpenseTx_(tx, manutCatId)) {
          usageByMonth[mk] = (usageByMonth[mk] || 0) + val;
        }
      });

      var months = Object.keys(depositsByMonth).concat(Object.keys(usageByMonth));
      if (months.indexOf(monthKey) === -1) months.push(monthKey);
      months.sort();

      var prevBalance = 0;
      var summary = { month: monthKey, predicted: 0, accumulated: 0, usage: 0, balance: 0 };
      months.forEach(function (mk) {
        var deposits = depositsByMonth[mk] || 0;
        var accumulated = prevBalance + deposits;
        var usage = usageByMonth[mk] || 0;
        var balance = accumulated - usage;
        if (mk === monthKey) {
          summary = {
            month: mk,
            predicted: deposits,
            accumulated: accumulated,
            usage: usage,
            balance: balance
          };
        }
        prevBalance = balance;
      });

      return summary;
    })();

    // NOVO: Subtrair transferências para reserva do lucro disponível
    var lucro = totalReceita - totalDespesa - transferenciasParaReserva;

    var cardKpis = { totalLimite: 0, totalUsado: 0, totalDisponivel: 0, cartoesAtivos: 0, faturasAbertas: 0 };
    try {
      var activeCards = [];
      if (typeof listCartoes === 'function') {
        var cardsRaw = listCartoes();
        var cardsParsed = (typeof cardsRaw === 'string') ? JSON.parse(cardsRaw) : cardsRaw;
        var cards = (cardsParsed && cardsParsed.ok && cardsParsed.data && cardsParsed.data.cartoes) ? cardsParsed.data.cartoes : [];
        activeCards = cards.filter(function (c) { return !!c.ativo; });
        cardKpis.cartoesAtivos = activeCards.length;
        cardKpis.totalLimite = activeCards.reduce(function (s, c) { return s + (parseFloat(c.limite_total) || 0); }, 0);
        // KPI "Cartões (Usado)" deve refletir o saldo usado atual dos cartões.
        cardKpis.totalUsado = activeCards.reduce(function (s, c) { return s + (parseFloat(c.limite_usado) || 0); }, 0);
        cardKpis.totalDisponivel = activeCards.reduce(function (s, c) { return s + (parseFloat(c.limite_disponivel) || 0); }, 0);
      }
      if (typeof cart_getOrCreateSheet === 'function' && typeof cart_readAll === 'function' &&
        typeof CFG_CART !== 'undefined' && typeof HEADERS_FATURAS !== 'undefined') {
        var sheetF = cart_getOrCreateSheet(CFG_CART.SHEET_FATURAS, HEADERS_FATURAS);
        var faturas = cart_readAll(sheetF) || [];
        var paymentMap = (typeof cart_getFaturaPaymentStats_ === 'function') ? cart_getFaturaPaymentStats_() : {};
        var openTotal = 0;
        faturas
          .filter(function (f) {
            var st = String(f.status || '').trim().toLowerCase();
            return st !== 'paga' && st !== 'pago' && st !== 'paid';
          })
          .forEach(function (f) {
            var cid = String(f.id_cartao || '');
            if (!cid) return;
            var valorAtual = 0;
            if (typeof cart_sumCardExpensesInPeriod_ === 'function') {
              valorAtual = cart_sumCardExpensesInPeriod_(cid, f.dt_inicio, f.dt_fim);
            } else {
              valorAtual = parseFloat(f.valor) || 0;
            }
            // Subtrair pagamentos parciais já realizados
            var faturaId = String(f.id_fatura || '');
            var payInfo = paymentMap[faturaId] || null;
            var totalJaPago = payInfo ? (parseFloat(payInfo.total_pago) || 0) : 0;
            var saldoPendente = Math.max(0, valorAtual - totalJaPago);
            openTotal += saldoPendente;
          });
        cardKpis.faturasAbertas = openTotal;
      }
    } catch (cardErr) {
      Logger.log('getDashboard cardKpis: ' + cardErr.message);
    }

    return okResponse({
      monthKey: monthKey,
      totalReceita: totalReceita,
      totalDespesa: totalDespesa,
      lucro: lucro,
      gastoCombustivel: gastoCombustivel,
      gastoAlimentacao: gastoAlimentacao,
      gastoManutencao: gastoManutencao,
      reservaManutencao: reserva,
      custosFixos: custosFixos,
      custoFixoParcela: custoFixoParcela,
      custoFixoSeguro: custoFixoSeguro,
      despesasPorCategoria: despesasPorCategoria,
      totalTransacoes: monthTx.length,
      cardKpis: cardKpis,
      transferenciasParaReserva: transferenciasParaReserva // NOVO: adicionar ao retorno
    });
  } catch (e) {
    return errorResponse(e.message, 'DASHBOARD_ERROR');
  }
}

function extractCardIdFromNotes_(notes) {
  var txt = String(notes || '');
  var m = txt.match(/\[CARTAO_ID:([^\]]+)\]/i) || txt.match(/CARTAO_ID:([A-Za-z0-9\-_]+)/i);
  return m ? String(m[1]).trim() : '';
}

function isCardBillPaymentNotes_(notes) {
  return /\[PAGAMENTO_FATURA:[^\]]+\]/i.test(String(notes || ''));
}

function resolveCardNameById_(cardId, cardNameMap) {
  if (!cardId) return '';
  if (cardNameMap && cardNameMap[cardId]) return cardNameMap[cardId];
  try {
    if (typeof cart_getOrCreateSheet === 'function' && typeof cart_readAll === 'function' &&
      typeof CFG_CART !== 'undefined' && typeof HEADERS_CARTOES !== 'undefined') {
      var sheet = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
      var raw = cart_readAll(sheet) || [];
      for (var i = 0; i < raw.length; i++) {
        if (String(raw[i].id) === String(cardId)) return String(raw[i].nome || '');
      }
    }
  } catch (e) { }
  return '';
}

function invalidateCardsCacheOnTxMutation_() {
  try {
    if (typeof cart_cacheInvalidate === 'function') {
      cart_cacheInvalidate();
    }
  } catch (e) {
    Logger.log('invalidateCardsCacheOnTxMutation_: ' + e.message);
  }
}

function getCardUsageFromFields_(type, notes, statusParcela, amount) {
  var cardId = extractCardIdFromNotes_(notes);
  if (!cardId) return { cardId: '', amount: 0 };
  if (isCardBillPaymentNotes_(notes)) return { cardId: cardId, amount: 0 };
  if (String(type || '').toUpperCase() !== 'DESPESA') return { cardId: cardId, amount: 0 };
  var st = String(statusParcela || '').toUpperCase();
  if (st === 'CANCELADA') return { cardId: cardId, amount: 0 };
  return { cardId: cardId, amount: parseFloat(amount) || 0 };
}

function isCardExpenseTx_(type, notes) {
  return String(type || '').toUpperCase() === 'DESPESA' &&
    !!extractCardIdFromNotes_(notes || '') &&
    !isCardBillPaymentNotes_(notes || '');
}

function applyCardUsageDelta_(cardId, delta) {
  var id = String(cardId || '');
  var d = parseFloat(delta) || 0;
  if (!id || !d) return;
  try {
    if (typeof cart_adjustCardUsed_ === 'function') {
      cart_adjustCardUsed_(id, d);
    }
  } catch (e) {
    Logger.log('applyCardUsageDelta_: ' + e.message);
  }
}

// ── listTransactions ─────────────────────────────────────────────

function listTransactions(params) {
  try {
    var month = params.month;
    var year = params.year;
    var filterType = params.type || '';
    var filterCatId = params.categoryId || '';
    var search = (params.search || '').toLowerCase();
    var page = params.page || 1;
    var pageSize = params.pageSize || 20;
    var tz = getTZ();
    var monthKey = year + '-' + (month < 10 ? '0' + month : '' + month);

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);

    // Filtrar por mês
    var filtered = allTx.filter(function (tx) {
      var d = tx.Data;
      var mk = '';
      if (d instanceof Date) {
        mk = Utilities.formatDate(d, tz, 'yyyy-MM');
      } else {
        var s = String(d);
        if (s.match(/^\d{4}-\d{2}/)) mk = s.substring(0, 7);
        else if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) mk = s.substring(6, 10) + '-' + s.substring(3, 5);
      }
      return mk === monthKey;
    });

    // Filtros adicionais
    if (filterType) {
      filtered = filtered.filter(function (tx) { return tx.Tipo === filterType; });
    }
    if (filterCatId) {
      filtered = filtered.filter(function (tx) { return String(tx.CategoriaId) === filterCatId; });
    }
    if (search) {
      filtered = filtered.filter(function (tx) {
        return (String(tx['Descrição']).toLowerCase().indexOf(search) !== -1) ||
          (String(tx.CategoriaNomeSnapshot).toLowerCase().indexOf(search) !== -1);
      });
    }

    // Ordenar por data decrescente
    filtered.sort(function (a, b) {
      var da = a.Data instanceof Date ? a.Data : new Date(a.Data);
      var db = b.Data instanceof Date ? b.Data : new Date(b.Data);
      return db - da;
    });

    var total = filtered.length;
    var totalPages = Math.ceil(total / pageSize) || 1;
    var start = (page - 1) * pageSize;
    var paged = filtered.slice(start, start + pageSize);

    // Mapear para objetos de saída
    var categories = getCategoriesCached(true);
    var catMap = {};
    categories.forEach(function (c) { catMap[c.id] = c; });

    var cardNameMap = {};
    if (typeof cart_getOrCreateSheet === 'function' && typeof cart_readAll === 'function' &&
      typeof CFG_CART !== 'undefined' && typeof HEADERS_CARTOES !== 'undefined') {
      try {
        var sheetCards = cart_getOrCreateSheet(CFG_CART.SHEET_CARTOES, HEADERS_CARTOES);
        (cart_readAll(sheetCards) || []).forEach(function (c) {
          cardNameMap[String(c.id)] = String(c.nome || '');
        });
      } catch (e) { }
    }

    var items = paged.map(function (tx) {
      var cat = catMap[String(tx.CategoriaId)] || null;
      var d = tx.Data;
      var dateStr = '';
      if (d instanceof Date) {
        dateStr = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
      } else {
        dateStr = String(d);
      }
      var isPaid = (tx.StatusParcela === 'PAGA' || tx.StatusParcela === 'COMPENSADO');
      var notes = tx.Notas || '';
      var cardId = extractCardIdFromNotes_(notes);
      var cardName = resolveCardNameById_(cardId, cardNameMap);
      return {
        id: tx.ID,
        date: dateStr,
        dateRaw: d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d),
        type: tx.Tipo,
        categoryId: tx.CategoriaId,
        categoryName: cat ? cat.name : tx.CategoriaNomeSnapshot,
        categoryEmoji: cat ? cat.emoji : '',
        categoryColor: cat ? cat.colorHex : '#999',
        description: tx['Descrição'],
        amount: parseFloat(tx.Valor) || 0,
        paymentMethod: tx.FormaPgto,
        km: tx.KM,
        statusParcela: tx.StatusParcela || '',
        grupoId: tx.GrupoId || '',
        numero: tx.Numero || '',
        total: tx.Total || '',
        notas: notes,
        cardId: cardId,
        cardName: cardName,
        fromWalletId: tx.CarteiraOrigem || '',
        toWalletId: tx.CarteiraDestino || '',
        debitoAutomatico: (tx.DebitoAutomatico === true || String(tx.DebitoAutomatico).toUpperCase() === 'TRUE'),
        isPaid: isPaid
      };
    });

    return okResponse({
      items: items,
      page: page,
      pageSize: pageSize,
      total: total,
      totalPages: totalPages
    });
  } catch (e) {
    return errorResponse(e.message, 'LIST_TX_ERROR');
  }
}

// ── createTransaction ────────────────────────────────────────────

function createTransaction(payload) {
  try {
    // Validações
    if (!payload.date) return errorResponse('Data é obrigatória.', 'VALIDATION');
    if (!payload.type || ['RECEITA', 'DESPESA', 'TRANSFER'].indexOf(payload.type) === -1)
      return errorResponse('Tipo inválido.', 'VALIDATION');
    var amount = parseFloat(payload.amount);
    var cardIdFromNotes = extractCardIdFromNotes_(payload.notas || '');
    var isCardExpense = (payload.type === 'DESPESA' && !!cardIdFromNotes && !isCardBillPaymentNotes_(payload.notas || ''));
    if (!isFinite(amount) || amount < 0 || (amount === 0 && !isCardExpense)) {
      return errorResponse(isCardExpense ? 'Valor deve ser maior ou igual a zero.' : 'Valor deve ser maior que zero.', 'VALIDATION');
    }

    // NOVO: Validação para transferências
    if (payload.type === 'TRANSFER') {
      if (!payload.fromWalletId || !payload.toWalletId)
        return errorResponse('Transferência requer carteira de origem e destino.', 'VALIDATION');
      if (payload.fromWalletId === payload.toWalletId)
        return errorResponse('Carteira de origem e destino não podem ser iguais.', 'VALIDATION');
      // Transferências não usam categoria
    } else {
      // RECEITA e DESPESA requerem categoria
      if (!payload.categoryId) return errorResponse('Categoria é obrigatória.', 'VALIDATION');
    }

    // Se for parcelamento ou recorrência
    if (payload.numParcelas && (payload.numParcelas > 1 || payload.numParcelas === -1)) {
      return createParceledTransaction(payload);
    }

    // Obter nome da categoria para snapshot (se não for transferência)
    var catName = '';
    if (payload.type !== 'TRANSFER') {
      var categories = getCategoriesCached(true);
      var cat = null;
      categories.forEach(function (c) { if (c.id === payload.categoryId) cat = c; });
      if (!cat) return errorResponse('Categoria não encontrada.', 'VALIDATION');
      catName = cat.name;
    }

    var id = generateUUID();
    var now = nowISO();
    var date = parseDateInput(payload.date);
    var dateStr = dateToYMD(date);

    cardIdFromNotes = extractCardIdFromNotes_(payload.notas || '');
    isCardExpense = (payload.type === 'DESPESA' && !!cardIdFromNotes && !isCardBillPaymentNotes_(payload.notas || ''));

    // Lógica de Status
    var isPaid = payload.paid !== undefined ? payload.paid : (date <= new Date());
    var statusParcela = payload.type === 'TRANSFER' ? '' : (isPaid ? 'PAGA' : 'PENDENTE');
    var dataPagamento = isPaid ? dateStr : '';
    if (isCardExpense) {
      // Despesa de cartão só pode ser baixada via pagamento de fatura.
      isPaid = false;
      statusParcela = 'PENDENTE';
      dataPagamento = '';
      payload.fromWalletId = '';
      if (!payload.paymentMethod) payload.paymentMethod = 'Crédito';
    }
    var debitoAutomatico = payload.autoDebit || false;
    if (isCardExpense) debitoAutomatico = false;

    var row = [
      id,
      dateStr,
      payload.type,
      payload.categoryId || '',
      catName,
      payload.description || '',
      amount,
      payload.paymentMethod || '',
      payload.km || '',
      now,
      now,
      '', // GrupoId
      '', // Numero
      '', // Total
      'UNICA', // TipoRecorrencia
      statusParcela,
      dateStr, // DataVencimento
      amount, // ValorOriginal
      0, // Juros
      0, // Desconto
      dataPagamento,
      (isPaid ? 'COMPENSADO' : 'PENDENTE'), // StatusReconciliacao
      (isPaid ? dateStr : ''), // DataCompensacao
      payload.notas || '', // Notas
      payload.fromWalletId || '', // CarteiraOrigem (NOVO)
      payload.toWalletId || '',    // CarteiraDestino (NOVO)
      debitoAutomatico // DebitoAutomatico
    ];

    appendRow(SHEET_LANCAMENTOS, row);
    var cardUsage = getCardUsageFromFields_(payload.type, payload.notas || '', statusParcela, amount);
    applyCardUsageDelta_(cardUsage.cardId, cardUsage.amount);

    // Recalcular reservas se for transferência ou manutenção
    var config = getConfigCached();
    if (payload.type === 'TRANSFER' || isMaintenanceExpenseTx_({
      Tipo: payload.type,
      CategoriaId: payload.categoryId,
      CategoriaNomeSnapshot: catName
    }, config['MANUTENCAO_CATEGORY_ID'])) {
      recalculateReserves();
    }
    invalidateCardsCacheOnTxMutation_();

    return okResponse({ id: id });
  } catch (e) {
    return errorResponse(e.message, 'CREATE_TX_ERROR');
  }
}

// ── createParceledTransaction ────────────────────────────────────

function createParceledTransaction(payload) {
  try {
    var numParcelas = parseInt(payload.numParcelas);
    var isInfinite = (numParcelas === -1);
    var loopLimit = isInfinite ? 12 : numParcelas; // Cria 12 meses se for infinito

    var totalAmount = parseFloat(payload.amount);
    // Se infinito, o valor já é o da parcela. Se parcelado, divide.
    var valorParcela = isInfinite ? totalAmount : (totalAmount / numParcelas);

    var baseDate = parseDateInput(payload.date);
    var grupoId = generateUUID();

    // Obter nome da categoria
    var categories = getCategoriesCached(true);
    var cat = null;
    categories.forEach(function (c) { if (c.id === payload.categoryId) cat = c; });
    if (!cat) return errorResponse('Categoria não encontrada.', 'VALIDATION');

    var config = getConfigCached();
    var isManutencao = isMaintenanceExpenseTx_({
      Tipo: payload.type,
      CategoriaId: payload.categoryId,
      CategoriaNomeSnapshot: cat.name
    }, config['MANUTENCAO_CATEGORY_ID']);
    var createdIds = [];
    var now = nowISO();
    var cardIdFromNotes = extractCardIdFromNotes_(payload.notas || '');
    var cardUsageDeltaTotal = 0;
    var isCardExpense = (payload.type === 'DESPESA' && !!cardIdFromNotes && !isCardBillPaymentNotes_(payload.notas || ''));
    if (isCardExpense && isInfinite) {
      return errorResponse('Cartão não permite recorrência infinita. Use parcelamento com número de parcelas.', 'CARD_INFINITE_BLOCKED');
    }

    // Configurações de parcelamento
    var tipoRecorrencia = isInfinite ? 'FIXA' : (payload.tipoRecorrencia || 'PARCELADA');
    var frequencia = payload.frequencia || 'MENSAL';
    var diaVencimento = payload.diaVencimento || baseDate.getDate();
    var calcularJuros = payload.calcularJuros || false;
    var taxaJuros = payload.taxaJuros || 0;

    for (var i = 0; i < loopLimit; i++) {
      // Calcular data da parcela (mês + i)
      var parcelaDate = new Date(baseDate);
      parcelaDate.setMonth(parcelaDate.getMonth() + i);

      // Ajustar dia de vencimento se especificado
      if (diaVencimento) {
        var lastDay = new Date(parcelaDate.getFullYear(), parcelaDate.getMonth() + 1, 0).getDate();
        var dia = Math.min(diaVencimento, lastDay);
        parcelaDate.setDate(dia);
      }

      var id = generateUUID();
      var desc = isInfinite
        ? (payload.description || '')
        : (payload.description || '') + ' (' + (i + 1) + '/' + numParcelas + ')';

      var dataVencimento = dateToYMD(parcelaDate);

      // Calcular juros se necessário (apenas para parcelado normal)
      var juros = 0;
      var valorFinal = valorParcela;
      if (!isInfinite && calcularJuros && taxaJuros > 0 && i > 0) {
        juros = valorParcela * (taxaJuros / 100) * i;
        valorFinal = valorParcela + juros;
      }

      // Status: primeira parcela obedece flag, demais pendentes
      var isFirst = (i === 0);
      var isPaid = false;

      if (isFirst) {
        isPaid = payload.paid !== undefined ? payload.paid : true;
      }
      // Se não for a primeira, sempre pendente

      var status = isPaid ? 'PAGA' : 'PENDENTE';
      var dataPagamento = isPaid ? dateToYMD(baseDate) : '';
      var debitoAutomatico = payload.autoDebit || false;
      var carteiraOrigem = payload.fromWalletId || '';
      var formaPgto = payload.paymentMethod || '';
      if (isCardExpense) {
        status = 'PENDENTE';
        dataPagamento = '';
        debitoAutomatico = false;
        carteiraOrigem = '';
        if (!formaPgto || formaPgto === '-') formaPgto = 'Crédito';
      }

      var row = [
        id,
        dateToYMD(parcelaDate),
        payload.type,
        payload.categoryId,
        cat.name,
        desc,
        valorFinal,
        formaPgto,
        payload.km || '',
        now,
        now,
        grupoId,
        isInfinite ? '' : (i + 1), // Numero da parcela (vazio se fixo)
        isInfinite ? '' : numParcelas, // Total parcelas (vazio se fixo)
        tipoRecorrencia,
        status,
        dataVencimento,
        valorParcela, // ValorOriginal
        juros,
        0, // Desconto
        dataPagamento,
        (status === 'PAGA') ? 'COMPENSADO' : 'PENDENTE',
        (status === 'PAGA') ? dateToYMD(baseDate) : '',
        payload.notas || '',
        carteiraOrigem,
        payload.toWalletId || '',
        debitoAutomatico
      ];

      appendRow(SHEET_LANCAMENTOS, row);
      createdIds.push(id);
      if (payload.type === 'DESPESA' && cardIdFromNotes && status !== 'CANCELADA' && !isCardBillPaymentNotes_(payload.notas || '')) {
        cardUsageDeltaTotal += (parseFloat(valorFinal) || 0);
      }
    }

    if (isCardExpense) {
      // Regra de cartão parcelado: consome o limite total no momento da compra.
      cardUsageDeltaTotal = parseFloat(totalAmount) || 0;
    }
    applyCardUsageDelta_(cardIdFromNotes, cardUsageDeltaTotal);

    // Criar registro de configuração do grupo
    var grupoConfigRow = [
      grupoId,
      payload.type,
      frequencia,
      diaVencimento,
      totalAmount,
      isInfinite ? -1 : numParcelas, // -1 indica infinito
      1, // ParcelasPagas (primeira já paga)
      0, // ParcelasCanceladas
      true, // PermitirEditacaoIndividual
      calcularJuros,
      taxaJuros,
      payload.observacoes || '',
      now,
      now
    ];
    appendRow(SHEET_PARCELAS_CONFIG, grupoConfigRow);

    // Recalcular reservas se alguma parcela for de manutenção
    if (isManutencao) {
      recalculateReserves();
    }
    invalidateCardsCacheOnTxMutation_();

    return okResponse({
      ids: createdIds,
      parcelas: loopLimit,
      grupoId: grupoId,
      valorParcela: valorParcela
    });
  } catch (e) {
    return errorResponse(e.message, 'CREATE_PARCELA_ERROR');
  }
}


// ── updateTransaction ────────────────────────────────────────────

function updateTransaction(id, payload) {
  var scope = payload.scope;
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);

    // Encontrar transação alvo
    var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);
    if (rowNum < 0) return errorResponse('Transação não encontrada.', 'NOT_FOUND');

    var existing = sheet.getRange(rowNum, 1, 1, HEADERS[SHEET_LANCAMENTOS].length).getValues()[0];

    // Se não for edição em lote ou não tiver recorrência
    var grupoId = existing[11];
    var isRecurrent = !!grupoId;

    if (!scope || scope === 'THIS' || !isRecurrent) {
      var res = _updateSingleTransaction(sheet, rowNum, id, payload, existing);
      recalculateReserves();
      invalidateCardsCacheOnTxMutation_();
      return res;
    }

    // --- Edição em Lote (ALL ou FOLLOWING) ---
    var numeroRef = parseInt(existing[12]) || 0;
    var dateRefOld = existing[1]; // Data original (string YYYY-MM-DD ou Date)
    if (dateRefOld instanceof Date) dateRefOld = dateToYMD(dateRefOld);

    var dateRefNew = payload.date ? String(payload.date) : dateRefOld;

    // Calcular Delta de dias se a data mudou
    var daysDelta = 0;
    if (payload.date && dateRefNew !== dateRefOld) {
      var dOld = parseDateInput(dateRefOld);
      var dNew = parseDateInput(dateRefNew);
      var diffTime = dNew.getTime() - dOld.getTime();
      daysDelta = Math.round(diffTime / (1000 * 3600 * 24));
    }

    // Buscar todas do grupo para update
    var allData = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var updatesCount = 0;

    for (var i = 0; i < allData.length; i++) {
      var t = allData[i];
      if (String(t.GrupoId) === String(grupoId)) {
        var tNum = parseInt(t.Numero) || 0;
        var currentRow = i + 2;

        var shouldUpdate = (scope === 'ALL') || (scope === 'FOLLOWING' && tNum >= numeroRef);

        if (shouldUpdate) {
          // Preparar payload específico
          var specificPayload = {};

          // Copiar campos editáveis
          if (payload.categoryId !== undefined) specificPayload.categoryId = payload.categoryId;
          if (payload.description !== undefined) specificPayload.description = payload.description;
          if (payload.amount !== undefined) specificPayload.amount = payload.amount;
          if (payload.paymentMethod !== undefined) specificPayload.paymentMethod = payload.paymentMethod;
          if (payload.km !== undefined) specificPayload.km = payload.km;
          if (payload.notas !== undefined) specificPayload.notas = payload.notas;
          if (payload.type !== undefined) specificPayload.type = payload.type;
          if (payload.fromWalletId !== undefined) specificPayload.fromWalletId = payload.fromWalletId;
          if (payload.toWalletId !== undefined) specificPayload.toWalletId = payload.toWalletId;

          // Data: aplicar delta
          if (daysDelta !== 0) {
            var tDateOld = t.Data;
            var tDateObj = (tDateOld instanceof Date) ? tDateOld : parseDateInput(tDateOld);
            if (tDateObj) {
              tDateObj.setDate(tDateObj.getDate() + daysDelta);
              specificPayload.date = dateToYMD(tDateObj);
            }
          }

          var currentRowValues = sheet.getRange(currentRow, 1, 1, HEADERS[SHEET_LANCAMENTOS].length).getValues()[0];
          _updateSingleTransaction(sheet, currentRow, t.ID, specificPayload, currentRowValues);
          updatesCount++;
        }
      }
    }

    recalculateReserves();
    invalidateCardsCacheOnTxMutation_();
    return okResponse({ count: updatesCount });

  } catch (e) {
    return errorResponse(e.message, 'UPDATE_ERROR');
  }
}

// Helper interno para atualizar UMA transação
// Helper interno para atualizar UMA transação
function _updateSingleTransaction(sheet, rowNum, id, payload, existing) {
  var oldType = existing[2];
  var oldNotes = existing[23];
  var oldUsage = getCardUsageFromFields_(oldType, oldNotes, existing[15], existing[6]);
  // Validações
  var date = payload.date ? parseDateInput(payload.date) : existing[1];
  var type = payload.type || existing[2];
  var categoryId = payload.categoryId || existing[3];
  var amount = payload.amount !== undefined ? parseFloat(payload.amount) : existing[6];
  var nextNotes = payload.notas !== undefined ? payload.notas : existing[23];
  var oldIsCardExpense = isCardExpenseTx_(oldType, oldNotes);
  var nextIsCardExpense = isCardExpenseTx_(type, nextNotes);

  if (!isFinite(amount) || amount < 0 || (amount === 0 && !nextIsCardExpense)) {
    throw new Error(nextIsCardExpense ? 'Valor deve ser maior ou igual a zero.' : 'Valor deve ser maior que zero.');
  }
  if (payload.paid !== undefined && (oldIsCardExpense || nextIsCardExpense)) {
    throw new Error('Lançamentos de cartão devem ser pagos pela fatura, não por compensação individual.');
  }

  // Snapshot da categoria
  var catName = existing[4];
  if (payload.categoryId && type !== 'TRANSFER') {
    var categories = getCategoriesCached(true);
    var cat = null;
    categories.forEach(function (c) { if (c.id === payload.categoryId) cat = c; });
    if (cat) catName = cat.name;
  }

  // Gerenciar Status e Débito Automático
  var statusParcela = existing[15];
  var dataPagamento = existing[20];
  var statusReconciliacao = existing[21];
  var dataCompensacao = existing[22];
  var debitoAutomatico = existing[26];

  // Se 'paid' for fornecido, atualiza status
  if (payload.paid !== undefined) {
    var isPaid = payload.paid;
    statusParcela = isPaid ? 'PAGA' : 'PENDENTE';

    if (isPaid) {
      // Se pagou, define datas. Se já tinha data, mantem, senão usa data da tx
      var txDateStr = date instanceof Date ? dateToYMD(date) : String(date);
      if (!dataPagamento) dataPagamento = txDateStr;
      statusReconciliacao = 'COMPENSADO';
      if (!dataCompensacao) dataCompensacao = txDateStr;
    } else {
      dataPagamento = '';
      statusReconciliacao = 'PENDENTE';
      dataCompensacao = '';
    }
  }

  if (payload.autoDebit !== undefined) {
    debitoAutomatico = payload.autoDebit;
  }
  var nextFromWallet = payload.fromWalletId !== undefined ? payload.fromWalletId : existing[24];
  var nextPaymentMethod = payload.paymentMethod !== undefined ? payload.paymentMethod : existing[7];
  if (nextIsCardExpense) {
    statusParcela = 'PENDENTE';
    dataPagamento = '';
    statusReconciliacao = 'PENDENTE';
    dataCompensacao = '';
    debitoAutomatico = false;
    nextFromWallet = '';
    if (!nextPaymentMethod || nextPaymentMethod === '-') nextPaymentMethod = 'Crédito';
  }

  var updatedRow = [
    id,
    date instanceof Date ? dateToYMD(date) : String(date),
    type,
    categoryId,
    catName,
    payload.description !== undefined ? payload.description : existing[5],
    amount,
    nextPaymentMethod,
    payload.km !== undefined ? payload.km : existing[8],
    existing[9], // CriadoEm
    nowISO(),    // AtualizadoEm
    existing[11], // GrupoId
    existing[12], // Numero
    existing[13], // Total
    existing[14], // TipoRecorrencia
    statusParcela,
    existing[16], // DataVencimento
    existing[17], // ValorOriginal
    existing[18], // Juros
    existing[19], // Desconto
    dataPagamento,
    statusReconciliacao,
    dataCompensacao,
    nextNotes, // Notas
    nextFromWallet, // CarteiraOrigem
    payload.toWalletId !== undefined ? payload.toWalletId : existing[25],      // CarteiraDestino
    debitoAutomatico // Coluna 27 (index 26)
  ];

  sheet.getRange(rowNum, 1, 1, updatedRow.length).setValues([updatedRow]);
  var newUsage = getCardUsageFromFields_(updatedRow[2], updatedRow[23], updatedRow[15], updatedRow[6]);
  if (oldUsage.cardId && oldUsage.cardId === newUsage.cardId) {
    applyCardUsageDelta_(oldUsage.cardId, (newUsage.amount || 0) - (oldUsage.amount || 0));
  } else {
    applyCardUsageDelta_(oldUsage.cardId, -(oldUsage.amount || 0));
    applyCardUsageDelta_(newUsage.cardId, (newUsage.amount || 0));
  }
  return okResponse({ id: id });
}

// ── Processamento de Débito Automático (Trigger Diário) ──────────

function processAutoDebits() {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var data = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var todayYMD = dateToYMD(new Date());
    var count = 0;

    for (var i = 0; i < data.length; i++) {
      var t = data[i];

      // Verificar elegibilidade: 
      // - Status != PAGA
      // - Tem flag DebitoAutomatico
      // - DataVencimento <= Hoje

      // Nota: DebitoAutomatico pode ser boolean ou string "TRUE"
      var isAuto = (t.DebitoAutomatico === true || String(t.DebitoAutomatico).toUpperCase() === 'TRUE');
      var isPending = (t.StatusParcela !== 'PAGA' && t.StatusParcela !== 'COMPENSADO');
      var vencimento = t.DataVencimento || t.Data;
      // vencimento é string YYYY-MM-DD

      if (isAuto && isPending && vencimento <= todayYMD) {
        var rowNum = i + 2;
        var rowValues = sheet.getRange(rowNum, 1, 1, HEADERS[SHEET_LANCAMENTOS].length).getValues()[0];

        // Atualizar para PAGO
        var payload = { paid: true };
        _updateSingleTransaction(sheet, rowNum, t.ID, payload, rowValues);
        count++;
      }
    }

    if (count > 0) {
      recalculateReserves();
      Logger.log('Processados ' + count + ' débitos automáticos.');
    }
    return count;
  } catch (e) {
    Logger.log('Erro no processAutoDebits: ' + e.message);
    return 0;
  }
}

function setupTrigger() {
  // Remove triggers existentes para evitar duplicidade
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processAutoDebits') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Cria novo trigger diário (toda dia entre 6am e 7am)
  ScriptApp.newTrigger('processAutoDebits')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}

// ── deleteTransaction ────────────────────────────────────────────

function deleteTransaction(id, scope) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);

    if (rowNum < 0) return errorResponse('Transação não encontrada.', 'NOT_FOUND');

    var existing = sheet.getRange(rowNum, 1, 1, HEADERS[SHEET_LANCAMENTOS].length).getValues()[0];
    var grupoId = existing[11];

    if (!scope || scope === 'THIS' || !grupoId) {
      var usageSingle = getCardUsageFromFields_(existing[2], existing[23], existing[15], existing[6]);
      deleteRow(SHEET_LANCAMENTOS, rowNum);
      applyCardUsageDelta_(usageSingle.cardId, -(usageSingle.amount || 0));
      recalculateReserves();
      invalidateCardsCacheOnTxMutation_();
      return okResponse({ deleted: true });
    }

    // --- Exclusão em Lote ---
    var numeroRef = parseInt(existing[12]) || 0;

    // Buscar todas do grupo
    var allData = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var rowsToDelete = [];
    var cardUsageDeltaByCard = {};

    for (var i = 0; i < allData.length; i++) {
      var t = allData[i];
      if (String(t.GrupoId) === String(grupoId)) {
        var tNum = parseInt(t.Numero) || 0;
        var currentRow = i + 2;

        if (scope === 'ALL') {
          rowsToDelete.push(currentRow);
          var uAll = getCardUsageFromFields_(t.Tipo, t.Notas, t.StatusParcela, t.Valor);
          if (uAll.cardId && uAll.amount) {
            cardUsageDeltaByCard[uAll.cardId] = (cardUsageDeltaByCard[uAll.cardId] || 0) - uAll.amount;
          }
        } else if (scope === 'FOLLOWING' && tNum >= numeroRef) {
          rowsToDelete.push(currentRow);
          var uFol = getCardUsageFromFields_(t.Tipo, t.Notas, t.StatusParcela, t.Valor);
          if (uFol.cardId && uFol.amount) {
            cardUsageDeltaByCard[uFol.cardId] = (cardUsageDeltaByCard[uFol.cardId] || 0) - uFol.amount;
          }
        }
      }
    }

    // Ordenar decrescente para deletar do fim para o começo
    rowsToDelete.sort(function (a, b) { return b - a; });

    rowsToDelete.forEach(function (r) {
      sheet.deleteRow(r);
    });
    Object.keys(cardUsageDeltaByCard).forEach(function (cardId) {
      applyCardUsageDelta_(cardId, cardUsageDeltaByCard[cardId]);
    });

    recalculateReserves();
    invalidateCardsCacheOnTxMutation_();
    return okResponse({ count: rowsToDelete.length });

  } catch (e) {
    return errorResponse(e.message, 'DELETE_TX_ERROR');
  }
}

// ── listCategories ───────────────────────────────────────────────

function listCategories(params) {
  try {
    var includeInactive = params && params.includeInactive;
    var cats = getCategoriesCached(!!includeInactive);
    return okResponse(cats);
  } catch (e) {
    return errorResponse(e.message, 'LIST_CAT_ERROR');
  }
}

// ── createCategory ───────────────────────────────────────────────

function createCategory(payload) {
  try {
    if (!payload.name || !payload.name.trim())
      return errorResponse('Nome e obrigatorio.', 'VALIDATION');
    if (payload.emoji && payload.emoji.length > 4)
      return errorResponse('Emoji: maximo 2 caracteres visuais.', 'VALIDATION');
    if (payload.colorHex && !payload.colorHex.match(/^#[0-9A-Fa-f]{6}$/))
      return errorResponse('Cor invalida (use #RRGGBB).', 'VALIDATION');
    if (!payload.applicableType || ['DESPESA', 'RECEITA', 'AMBOS'].indexOf(payload.applicableType) === -1)
      return errorResponse('TipoAplicavel invalido.', 'VALIDATION');

    var existing = getCategoriesCached(true);
    var maxOrder = 0;
    existing.forEach(function (cat) { if ((Number(cat.order) || 0) > maxOrder) maxOrder = Number(cat.order) || 0; });

    var id = generateUUID();
    var now = nowISO();
    var ownerId = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID) ? String(AUTH_EXECUTION_USER_ID) : '';

    var row = [
      id,
      payload.name.trim(),
      payload.applicableType,
      payload.emoji || '??',
      payload.colorHex || '#A0A0A0',
      maxOrder + 1,
      true,
      now,
      now,
      ownerId
    ];

    appendRow(SHEET_CATEGORIAS, row);
    invalidateCache('categories');
    invalidateCache('categories_all');

    return okResponse({ id: id });
  } catch (e) {
    return errorResponse(e.message, 'CREATE_CAT_ERROR');
  }
}
// -- updateCategory ───────────────────────────────────────────────

function updateCategory(payload) {
  try {
    if (!payload.id) return errorResponse('ID e obrigatorio.', 'VALIDATION');

    var rowNum = findRowIndex(SHEET_CATEGORIAS, 1, payload.id);
    if (rowNum < 0) {
      var allCats = getCategoriesCached(true);
      var isPublic = allCats.some(function (c) { return String(c.id) === String(payload.id) && c.isPublic; });
      if (isPublic) return errorResponse('Categoria publica nao pode ser editada.', 'PUBLIC_CATEGORY_READONLY');
      return errorResponse('Categoria nao encontrada.', 'NOT_FOUND');
    }

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CATEGORIAS);
    var existing = sheet.getRange(rowNum, 1, 1, HEADERS[SHEET_CATEGORIAS].length).getValues()[0];

    if (payload.colorHex && !payload.colorHex.match(/^#[0-9A-Fa-f]{6}$/))
      return errorResponse('Cor invalida.', 'VALIDATION');
    if (payload.applicableType && ['DESPESA', 'RECEITA', 'AMBOS'].indexOf(payload.applicableType) === -1)
      return errorResponse('TipoAplicavel invalido.', 'VALIDATION');

    var ownerUserId = existing.length >= 10 ? existing[9] : '';
    var currentUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID) ? String(AUTH_EXECUTION_USER_ID) : '';
    if (String(ownerUserId || '') !== currentUser) {
      return errorResponse('Voce nao pode editar categoria de outro usuario.', 'FORBIDDEN');
    }

    var updatedRow = [
      payload.id,
      payload.name !== undefined ? payload.name.trim() : existing[1],
      payload.applicableType !== undefined ? payload.applicableType : existing[2],
      payload.emoji !== undefined ? payload.emoji : existing[3],
      payload.colorHex !== undefined ? payload.colorHex : existing[4],
      existing[5],
      payload.active !== undefined ? payload.active : existing[6],
      existing[7],
      nowISO(),
      ownerUserId
    ];

    updateRow(SHEET_CATEGORIAS, rowNum, updatedRow);
    invalidateCache('categories');
    invalidateCache('categories_all');

    return okResponse({ id: payload.id });
  } catch (e) {
    return errorResponse(e.message, 'UPDATE_CAT_ERROR');
  }
}
// -- reorderCategories ────────────────────────────────────────────

function reorderCategories(items) {
  try {
    if (!Array.isArray(items)) return errorResponse('Envie array de {id, order}.', 'VALIDATION');

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CATEGORIAS);
    var currentUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID) ? String(AUTH_EXECUTION_USER_ID) : '';
    var all = getSheetDataAsObjects(SHEET_CATEGORIAS);
    var ownerById = {};
    all.forEach(function (c) { ownerById[String(c.ID)] = String(c.OwnerUserId || ''); });

    var affected = 0;
    items.forEach(function (item) {
      if (ownerById[String(item.id)] !== currentUser) return;
      var rowNum = findRowIndex(SHEET_CATEGORIAS, 1, item.id);
      if (rowNum > 0) {
        sheet.getRange(rowNum, 6).setValue(item.order);
        sheet.getRange(rowNum, 9).setValue(nowISO());
        affected++;
      }
    });

    invalidateCache('categories');
    invalidateCache('categories_all');

    return okResponse({ reordered: affected });
  } catch (e) {
    return errorResponse(e.message, 'REORDER_ERROR');
  }
}
// -- deleteCategory ───────────────────────────────────────────────

function deleteCategory(id) {
  try {
    if (!id) return errorResponse('ID e obrigatorio.', 'VALIDATION');

    var rowNum = findRowIndex(SHEET_CATEGORIAS, 1, id);
    if (rowNum < 0) {
      var allCats = getCategoriesCached(true);
      var isPublic = allCats.some(function (c) { return String(c.id) === String(id) && c.isPublic; });
      if (isPublic) return errorResponse('Categoria publica nao pode ser removida.', 'PUBLIC_CATEGORY_READONLY');
      return errorResponse('Categoria nao encontrada.', 'NOT_FOUND');
    }

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CATEGORIAS);
    var existing = sheet.getRange(rowNum, 1, 1, HEADERS[SHEET_CATEGORIAS].length).getValues()[0];
    var ownerUserId = existing.length >= 10 ? existing[9] : '';
    var currentUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID) ? String(AUTH_EXECUTION_USER_ID) : '';
    if (String(ownerUserId || '') !== currentUser) {
      return errorResponse('Voce nao pode remover categoria de outro usuario.', 'FORBIDDEN');
    }

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var hasTransactions = allTx.some(function (tx) {
      return String(tx.CategoriaId) === id;
    });

    if (hasTransactions) {
      var updatedRow = [
        id,
        existing[1],
        existing[2],
        existing[3],
        existing[4],
        existing[5],
        false,
        existing[7],
        nowISO(),
        ownerUserId
      ];

      updateRow(SHEET_CATEGORIAS, rowNum, updatedRow);
      invalidateCache('categories');
      invalidateCache('categories_all');

      return errorResponse('Categoria possui transacoes associadas. Foi desativada ao inves de deletada.', 'HAS_TRANSACTIONS');
    }

    deleteRow(SHEET_CATEGORIAS, rowNum);
    invalidateCache('categories');
    invalidateCache('categories_all');

    return okResponse({ deleted: true });
  } catch (e) {
    return errorResponse(e.message, 'DELETE_CAT_ERROR');
  }
}

// -- getConfig ────────────────────────────────────────────────────

function getConfig() {
  try {
    var config = getConfigCached();
    return okResponse(config);
  } catch (e) {
    return errorResponse(e.message, 'GET_CONFIG_ERROR');
  }
}

// ── Helpers ──────────────────────────────────────────────────────

function normalizeWalletId(id) {
  return String(id || '').trim().toLowerCase();
}

function extractMonthKey(date, tz) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, tz, 'yyyy-MM');
  }
  var s = String(date);
  if (s.match(/^\d{4}-\d{2}/)) return s.substring(0, 7);
  else if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) return s.substring(6, 10) + '-' + s.substring(3, 5);
  return '';
}

// ── saveConfig ───────────────────────────────────────────────────

function saveConfig(config) {
  try {
    var allowedKeys = [
      'RESERVA_MANUTENCAO_MENSAL', 'MANUTENCAO_CATEGORY_ID',
      'RESERVA_CATEGORY_ID', 'META_RESERVA_MANUTENCAO',
      'TIMEZONE', 'MOEDA', 'THEME_DEFAULT', 'ALLOWLIST_EMAILS'
    ];

    var needRecalc = false;
    Object.keys(config).forEach(function (key) {
      if (allowedKeys.indexOf(key) !== -1) {
        setConfigValue(key, config[key]);
        if (key === 'MANUTENCAO_CATEGORY_ID' || key === 'RESERVA_CATEGORY_ID') {
          needRecalc = true;
        }
      }
    });

    invalidateCache('config');

    if (needRecalc) {
      recalculateReserves();
    }

    return okResponse({ saved: true });
  } catch (e) {
    return errorResponse(e.message, 'SAVE_CONFIG_ERROR');
  }
}

// ── setMaintenanceCategoryId ─────────────────────────────────────

function setMaintenanceCategoryId(categoryId) {
  try {
    setConfigValue('MANUTENCAO_CATEGORY_ID', categoryId);
    invalidateCache('config');
    recalculateReserves();
    return okResponse({ saved: true });
  } catch (e) {
    return errorResponse(e.message, 'SET_MANUT_ERROR');
  }
}

// ── Validar nome de conta único ──────────────────────────────────

function validateAccountName(name, excludeId) {
  var all = getSheetDataAsObjects(SHEET_CONTAS);
  var nameLower = String(name).toLowerCase().trim();

  for (var i = 0; i < all.length; i++) {
    var existingName = String(all[i].Nome).toLowerCase().trim();
    var existingId = String(all[i].ID);
    var isActive = all[i].Ativo === true || all[i].Ativo === 'TRUE' || all[i].Ativo === 1;

    // Apenas conflita se a conta existente estiver ATIVA
    if (isActive && existingName === nameLower && existingId !== excludeId) {
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

  // Se a conta não for encontrada ou estiver inativa, podemos retornar 0,
  // mas aqui estamos apenas calculando saldo.
  // Note: calculateAccountBalance usa SHEET_CONTAS direto para pegar SaldoInicial.

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

// ── setTheme (por usuário) ───────────────────────────────────────

function setTheme(theme) {
  try {
    PropertiesService.getUserProperties().setProperty('THEME', theme);
    return okResponse({ theme: theme });
  } catch (e) {
    return errorResponse(e.message, 'SET_THEME_ERROR');
  }
}

// ── exportCsv ────────────────────────────────────────────────────

function exportCsv(params) {
  try {
    var month = params.month;
    var year = params.year;
    var monthKey = year + '-' + (month < 10 ? '0' + month : '' + month);
    var tz = getTZ();

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var filtered = allTx.filter(function (tx) {
      var d = tx.Data;
      var mk = '';
      if (d instanceof Date) mk = Utilities.formatDate(d, tz, 'yyyy-MM');
      else {
        var s = String(d);
        if (s.match(/^\d{4}-\d{2}/)) mk = s.substring(0, 7);
      }
      return mk === monthKey;
    });

    var categories = getCategoriesCached(true);
    var catMap = {};
    categories.forEach(function (c) { catMap[c.id] = c; });

    var csv = 'Data;Tipo;Categoria;Descrição;Valor;Forma Pgto;KM\n';
    filtered.forEach(function (tx) {
      var d = tx.Data instanceof Date ? Utilities.formatDate(tx.Data, tz, 'dd/MM/yyyy') : String(tx.Data);
      var cat = catMap[String(tx.CategoriaId)];
      var catName = cat ? cat.emoji + ' ' + cat.name : tx.CategoriaNomeSnapshot;
      var val = (parseFloat(tx.Valor) || 0).toFixed(2).replace('.', ',');
      csv += [d, tx.Tipo, catName, tx['Descrição'], val, tx.FormaPgto, tx.KM].join(';') + '\n';
    });


    return okResponse({ csv: csv, filename: 'financeiro_' + monthKey + '.csv' });
  } catch (e) {
    return errorResponse(e.message, 'EXPORT_ERROR');
  }
}

// ── Exportação Avançada ──────────────────────────────────────────

function exportSpreadsheet(params) {
  try {
    var result = generateExportSheet(params);
    return okResponse({ url: result.url, filename: result.filename });
  } catch (e) {
    return errorResponse(e.message, 'EXPORT_SHEET_ERROR');
  }
}

function exportPdf(params) {
  try {
    var result = generateExportSheet(params);
    var ss = SpreadsheetApp.openByUrl(result.url);
    var sheetId = ss.getSheets()[0].getSheetId();
    var pdfBlob = DriveApp.getFileById(ss.getId()).getBlob().getAs('application/pdf');
    var base64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Opcional: Deletar a planilha temporária se for PDF
    // DriveApp.getFileById(ss.getId()).setTrashed(true);

    return okResponse({
      base64: base64,
      filename: result.filename.replace('.xlsx', '.pdf').replace('.xls', '.pdf'),
      url: result.url // Retorna também a URL caso queira manter
    });
  } catch (e) {
    return errorResponse(e.message, 'EXPORT_PDF_ERROR');
  }
}

function generateExportSheet(params) {
  var month = params.month;
  var year = params.year;
  var monthKey = year + '-' + (month < 10 ? '0' + month : '' + month);
  var tz = getTZ();

  // Filtrar dados
  var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
  var filtered = allTx.filter(function (tx) {
    var mk = extractMonthKey(tx.Data, tz);
    return mk === monthKey;
  });

  // Criar planilha
  var filename = 'Extrato Financeiro - ' + month + '/' + year;
  var ss = SpreadsheetApp.create(filename);
  var sheet = ss.getActiveSheet();
  sheet.setName('Extrato');

  // Cabeçalho
  var headers = ['Data', 'Tipo', 'Categoria', 'Descrição', 'Valor', 'Forma Pgto', 'Situação'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4a4a4a')
    .setFontColor('#ffffff');

  if (filtered.length > 0) {
    var categories = getCategoriesCached(true);
    var catMap = {};
    categories.forEach(function (c) { catMap[c.id] = c; });

    var rows = filtered.map(function (tx) {
      var d = tx.Data instanceof Date ? tx.Data : new Date(tx.Data);
      var cat = catMap[String(tx.CategoriaId)];
      var catName = cat ? cat.name : (tx.CategoriaNomeSnapshot || 'Outros');
      // Adicionar emoji se disponível
      if (cat && cat.emoji) catName = cat.emoji + ' ' + catName;

      var val = parseFloat(tx.Valor) || 0;
      var situacao = '';
      if (tx.Tipo === 'DESPESA') situacao = 'Pago';
      // Se tiver lógica de "Pago/Pendente" real, usar aqui. 
      // O modelo atual não tem coluna explícita 'Pago' na sheet Lancamentos (apenas StatusParcela ou checkbox UI).
      // Vou assumir 'Confirmado' por enquanto.

      return [d, tx.Tipo, catName, tx['Descrição'], val, tx.FormaPgto, situacao];
    });

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // Formatação
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy'); // Data
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('R$ #,##0.00'); // Valor
  }

  // Auto-resize
  try {
    sheet.autoResizeColumns(1, headers.length);
  } catch (e) { }

  return { url: ss.getUrl(), filename: filename };
}

// ── Relatórios — últimos 12 meses ────────────────────────────────

function getReport12Months() {
  try {
    var tz = getTZ();
    var now = new Date();
    var config = getConfigCached();
    var manutCatId = config['MANUTENCAO_CATEGORY_ID'] || '';

    // Gerar últimos 12 meses
    var months = [];
    for (var i = 11; i >= 0; i--) {
      var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push(Utilities.formatDate(d, tz, 'yyyy-MM'));
    }

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);

    // Categorias para buscar combustível
    var categories = getCategoriesCached(true);
    var combustivelId = '';
    categories.forEach(function (c) {
      if (c.name.toLowerCase() === 'combustível') combustivelId = c.id;
    });

    var accountsAll = getAccountsCached(true);
    var investmentsAccount = accountsAll.find(function (a) {
      return String(a.nome || '').toLowerCase() === 'investimentos';
    }) || null;
    var investmentsId = investmentsAccount ? String(investmentsAccount.id) : '';
    var investmentsInitial = investmentsAccount ? (parseFloat(investmentsAccount.saldoInicial) || 0) : 0;

    var report = months.map(function (mk) {
      var txMonth = allTx.filter(function (tx) {
        var d = tx.Data;
        var m = '';
        if (d instanceof Date) m = Utilities.formatDate(d, tz, 'yyyy-MM');
        else {
          var s = String(d);
          if (s.match(/^\d{4}-\d{2}/)) m = s.substring(0, 7);
        }
        return m === mk;
      });

      var receita = 0, despesa = 0, combustivel = 0, manutencao = 0;
      txMonth.forEach(function (tx) {
        var val = parseFloat(tx.Valor) || 0;
        if (tx.Tipo === 'RECEITA') receita += val;
        else {
          despesa += val;
          if (String(tx.CategoriaId) === combustivelId) combustivel += val;
          if (isMaintenanceExpenseTx_(tx, manutCatId)) manutencao += val;
        }
      });

      var reservaBalance = getInvestmentBalanceByMonth_(mk, tz, allTx, investmentsId, investmentsInitial);

      return {
        month: mk,
        receita: receita,
        despesa: despesa,
        lucro: receita - despesa,
        combustivel: combustivel,
        manutencao: manutencao,
        reservaBalance: reservaBalance
      };
    });

    return okResponse(report);
  } catch (e) {
    return errorResponse(e.message, 'REPORT_ERROR');
  }
}

function getInvestmentBalanceByMonth_(monthKey, tz, allTx, investmentsId, initialBalance) {
  if (!investmentsId) return 0;
  var balance = parseFloat(initialBalance) || 0;

  allTx.forEach(function (tx) {
    var d = tx.Data;
    var mk = '';
    if (d instanceof Date) mk = Utilities.formatDate(d, tz, 'yyyy-MM');
    else {
      var s = String(d || '');
      if (s.match(/^\d{4}-\d{2}/)) mk = s.substring(0, 7);
      else if (s.match(/^\d{2}\/\d{2}\/\d{4}/)) mk = s.substring(6, 10) + '-' + s.substring(3, 5);
    }
    if (!mk || mk > monthKey) return;

    var status = tx.StatusParcela || 'PAGA';
    if (status !== 'PAGA' && status !== 'COMPENSADO') return;

    var val = parseFloat(tx.Valor) || 0;
    var type = String(tx.Tipo || '');
    var fromWallet = String(tx.CarteiraOrigem || '');
    var toWallet = String(tx.CarteiraDestino || '');

    if (type === 'RECEITA' && toWallet === investmentsId) balance += val;
    if (type === 'DESPESA' && fromWallet === investmentsId) balance -= val;
    if (type === 'TRANSFER' || type === 'TRANSFERENCIA') {
      if (fromWallet === investmentsId) balance -= val;
      if (toWallet === investmentsId) balance += val;
    }
  });

  return balance;
}

// ── Contas (Accounts) com cache ──────────────────────────────────

function getAccountsCached(includeInactive) {
  var hasAuthUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID);
  var scopeUser = hasAuthUser ? String(AUTH_EXECUTION_USER_ID) : 'public';
  var key = (includeInactive ? 'accounts_all_v3' : 'accounts_v3') + '_' + scopeUser;
  var cached = hasAuthUser ? null : getCached(key);

  if (cached) {
    Logger.log('GET_ACCOUNTS: Retornando do cache: ' + key + ' (count: ' + cached.length + ')');
    return cached;
  }

  var all = getSheetDataAsObjects(SHEET_CONTAS);
  Logger.log('GET_ACCOUNTS: Lendo da planilha. Total linhas (objetos): ' + all.length);

  var result = all.map(function (a, idx) {
    if (idx === 0) Logger.log('GET_ACCOUNTS: Amostra do primeiro objeto: ' + JSON.stringify(a));
    return {
      id: a.ID,
      nome: a.Nome,
      tipo: a.Tipo,
      saldoInicial: parseFloat(a.SaldoInicial) || 0,
      ativo: a.Ativo === true || a.Ativo === 'TRUE' || a.Ativo === 1,
      ordem: Number(a.Ordem) || 0,
      criadoEm: a.CriadoEm,
      atualizadoEm: a.AtualizadoEm,
      instituicao: a.Instituicao || 'Outro'
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
    { nome: 'Investimentos', tipo: 'Investimentos', saldoInicial: 0, ordem: 2, instituicao: 'Outro' },
      ];

  var rows = accounts.map(function (a) {
    var id = Utilities.getUuid();
    return [id, a.nome, a.tipo, a.saldoInicial, true, a.ordem, now, now, a.instituicao];
  });

  return rows;
}

// ── Migração legada ──────────────────────────────────────────────

function migrateLegacyCategories() {
  try {
    var categories = getCategoriesCached(true);
    var catByName = {};
    categories.forEach(function (c) {
      catByName[c.name.toLowerCase()] = c.id;
    });

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    if (sheet.getLastRow() < 2) return okResponse({ migrated: 0 });

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS[SHEET_LANCAMENTOS].length).getValues();
    var migrated = 0;

    data.forEach(function (row, i) {
      var catId = String(row[3]).trim();
      var catSnapshot = String(row[4]).trim().toLowerCase();

      // Se não tem CategoriaId mas tem snapshot
      if ((!catId || catId === '' || catId === 'undefined') && catSnapshot) {
        var matchId = catByName[catSnapshot];
        if (matchId) {
          sheet.getRange(i + 2, 4).setValue(matchId); // CategoriaId
          sheet.getRange(i + 2, 11).setValue(nowISO()); // AtualizadoEm
          migrated++;
        }
      }
    });

    if (migrated > 0) {
      invalidateAllCache();
      recalculateReserves();
    }

    return okResponse({ migrated: migrated });
  } catch (e) {
    return errorResponse(e.message, 'MIGRATE_ERROR');
  }
}

/* ================================================================
   Gestão de Parcelas — Novos Endpoints
   ================================================================ */

// ── markInstallmentAsPaid ────────────────────────────────────────

function markInstallmentAsPaid(id, dataPagamento, valorPago) {
  try {
    var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);
    if (rowNum < 0) return errorResponse('Parcela não encontrada.', 'NOT_FOUND');

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var headers = HEADERS[SHEET_LANCAMENTOS];
    var existing = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

    // Índices das colunas
    var idxStatusParcela = headers.indexOf('StatusParcela');
    var idxDataPagamento = headers.indexOf('DataPagamento');
    var idxValor = headers.indexOf('Valor');
    var idxGrupoId = headers.indexOf('GrupoId');
    var idxStatusRec = headers.indexOf('StatusReconciliacao');
    var idxDataComp = headers.indexOf('DataCompensacao');

    if (isCardExpenseTx_(existing[2], existing[23])) {
      return errorResponse('Lançamentos de cartão devem ser pagos pela fatura.', 'CARD_TX_PAY_BLOCKED');
    }

    if (existing[idxStatusParcela] === 'PAGA') {
      return errorResponse('Parcela já está marcada como paga.', 'ALREADY_PAID');
    }

    var dtPg = dataPagamento || dateToYMD(new Date());
    sheet.getRange(rowNum, idxStatusParcela + 1).setValue('PAGA');
    sheet.getRange(rowNum, idxDataPagamento + 1).setValue(dtPg);
    if (idxStatusRec >= 0) sheet.getRange(rowNum, idxStatusRec + 1).setValue('COMPENSADO');
    if (idxDataComp >= 0) sheet.getRange(rowNum, idxDataComp + 1).setValue(dtPg);

    if (valorPago && parseFloat(valorPago) !== parseFloat(existing[idxValor])) {
      sheet.getRange(rowNum, idxValor + 1).setValue(parseFloat(valorPago));
    }

    sheet.getRange(rowNum, headers.indexOf('AtualizadoEm') + 1).setValue(nowISO());

    var grupoId = existing[idxGrupoId];
    if (grupoId) updateGrupoParcelasPagas(grupoId, 1);

    var catId = existing[headers.indexOf('CategoriaId')];
    var catSnapshot = existing[headers.indexOf('CategoriaNomeSnapshot')];
    var config = getConfigCached();
    if (isMaintenanceExpenseTx_({
      Tipo: 'DESPESA',
      CategoriaId: catId,
      CategoriaNomeSnapshot: catSnapshot
    }, config['MANUTENCAO_CATEGORY_ID'])) {
      recalculateReserves();
    }
    invalidateCardsCacheOnTxMutation_();

    return okResponse({ id: id, status: 'PAGA' });
  } catch (e) {
    return errorResponse(e.message, 'MARK_PAID_ERROR');
  }
}

// ── markMultipleAsPaid ───────────────────────────────────────────

function markMultipleAsPaid(payload) {
  try {
    var ids = payload.ids || [];
    var dataPagamento = payload.dataPagamento || nowISO().substring(0, 10); // Data de hoje se não informado

    if (!ids || ids.length === 0) return errorResponse('Nenhum ID fornecido.', 'VALIDATION');

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var headers = HEADERS[SHEET_LANCAMENTOS];

    var idxStatus = headers.indexOf('StatusParcela') + 1;
    var idxData = headers.indexOf('DataPagamento') + 1;
    var idxAtualizado = headers.indexOf('AtualizadoEm') + 1;
    var idxGrupo = headers.indexOf('GrupoId') + 1;
    var idxTipo = headers.indexOf('Tipo') + 1;
    var idxNotas = headers.indexOf('Notas') + 1;
    var idxStatusRec = headers.indexOf('StatusReconciliacao') + 1;
    var idxDataComp = headers.indexOf('DataCompensacao') + 1;

    var gruposAfetados = {};
    var processed = 0;
    var skippedCard = [];

    ids.forEach(function (id) {
      var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);
      if (rowNum > 0) {
        var tipo = idxTipo > 0 ? sheet.getRange(rowNum, idxTipo).getValue() : '';
        var notas = idxNotas > 0 ? sheet.getRange(rowNum, idxNotas).getValue() : '';
        if (isCardExpenseTx_(tipo, notas)) {
          skippedCard.push(id);
          return;
        }
        // Verificar se já não estava paga para não incrementar contador duplicado
        var currentStatus = sheet.getRange(rowNum, idxStatus).getValue();
        if (currentStatus !== 'PAGA') {
          sheet.getRange(rowNum, idxStatus).setValue('PAGA');
          sheet.getRange(rowNum, idxData).setValue(dataPagamento);
          if (idxStatusRec > 0) sheet.getRange(rowNum, idxStatusRec).setValue('COMPENSADO');
          if (idxDataComp > 0) sheet.getRange(rowNum, idxDataComp).setValue(dataPagamento);
          sheet.getRange(rowNum, idxAtualizado).setValue(nowISO());
          processed++;

          var gId = sheet.getRange(rowNum, idxGrupo).getValue();
          if (gId) gruposAfetados[gId] = (gruposAfetados[gId] || 0) + 1;
        }
      }
    });

    // Atualizar contadores de parcelas pagas nos grupos
    Object.keys(gruposAfetados).forEach(function (gId) {
      updateGrupoParcelasPagas(gId, gruposAfetados[gId]);
    });

    // Recalcular reservas se necessário (simplificado: recalcula sempre que houver pagamento em massa para garantir)
    recalculateReserves();
    invalidateCardsCacheOnTxMutation_();

    clearCache(CACHE_KEY_TRANSACTIONS);
    return okResponse({ count: processed, skippedCard: skippedCard });
  } catch (e) {
    return errorResponse(e.message, 'BULK_PAY_ERROR');
  }
}

// ── cancelInstallment ────────────────────────────────────────────

function cancelInstallment(id, cancelType) {
  try {
    var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);
    if (rowNum < 0) return errorResponse('Parcela não encontrada.', 'NOT_FOUND');

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var headers = HEADERS[SHEET_LANCAMENTOS];
    var existing = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

    var grupoId = existing[headers.indexOf('GrupoId')];
    var numero = existing[headers.indexOf('Numero')];
    var idxStatusParcela = headers.indexOf('StatusParcela');
    var canceledCount = 0;

    if (cancelType === 'single') {
      var singleUsage = getCardUsageFromFields_(existing[2], existing[23], existing[idxStatusParcela], existing[6]);
      sheet.getRange(rowNum, idxStatusParcela + 1).setValue('CANCELADA');
      sheet.getRange(rowNum, headers.indexOf('AtualizadoEm') + 1).setValue(nowISO());
      canceledCount = 1;
      applyCardUsageDelta_(singleUsage.cardId, -(singleUsage.amount || 0));
    } else if (cancelType === 'future' || cancelType === 'all') {
      var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
      var toCancelIds = [];
      var cardUsageDeltaByCard = {};

      allTx.forEach(function (tx) {
        if (String(tx.GrupoId) === String(grupoId)) {
          var txNum = tx.Numero;
          if (cancelType === 'all' || txNum >= numero) {
            if (tx.StatusParcela !== 'PAGA' && tx.StatusParcela !== 'CANCELADA') {
              toCancelIds.push(tx.ID);
              var bulkUsage = getCardUsageFromFields_(tx.Tipo, tx.Notas, tx.StatusParcela, tx.Valor);
              if (bulkUsage.cardId && bulkUsage.amount) {
                cardUsageDeltaByCard[bulkUsage.cardId] = (cardUsageDeltaByCard[bulkUsage.cardId] || 0) - bulkUsage.amount;
              }
            }
          }
        }
      });

      toCancelIds.forEach(function (txId) {
        var row = findRowIndex(SHEET_LANCAMENTOS, 1, txId);
        if (row > 0) {
          sheet.getRange(row, idxStatusParcela + 1).setValue('CANCELADA');
          sheet.getRange(row, headers.indexOf('AtualizadoEm') + 1).setValue(nowISO());
          canceledCount++;
        }
      });
      Object.keys(cardUsageDeltaByCard).forEach(function (cardId) {
        applyCardUsageDelta_(cardId, cardUsageDeltaByCard[cardId]);
      });
    }

    if (grupoId && canceledCount > 0) {
      updateGrupoParcelasCanceladas(grupoId, canceledCount);
    }
    invalidateCardsCacheOnTxMutation_();

    return okResponse({ canceledCount: canceledCount });
  } catch (e) {
    return errorResponse(e.message, 'CANCEL_ERROR');
  }
}

// ── getInstallmentDetails ────────────────────────────────────────

function getInstallmentDetails(grupoId) {
  try {
    var grupoConfig = null;
    var configData = getSheetDataAsObjects(SHEET_PARCELAS_CONFIG);
    configData.forEach(function (cfg) {
      if (String(cfg.GrupoId) === String(grupoId)) grupoConfig = cfg;
    });

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var parcelas = allTx.filter(function (tx) {
      return String(tx.GrupoId) === String(grupoId);
    });

    if (parcelas.length === 0) {
      return errorResponse('Nenhuma parcela encontrada para este grupo.', 'NOT_FOUND');
    }

    // Fallback se não achar config (reconstruir baseando nas parcelas)
    if (!grupoConfig) {
      var first = parcelas[0];
      grupoConfig = {
        Tipo: first.Tipo,
        Frequencia: 'MENSAL', // Assume mensal
        ValorTotal: parcelas.reduce(function (acc, p) { return acc + (parseFloat(p.Valor) || 0); }, 0),
        NumParcelas: first.Total || parcelas.length,
        DiaVencimento: new Date(first.DataVencimento).getDate()
      };
    }

    var categories = getCategoriesCached(true);
    var catMap = {};
    categories.forEach(function (c) { catMap[c.id] = c; });

    var tz = getTZ();
    var items = parcelas.map(function (tx) {
      var cat = catMap[String(tx.CategoriaId)] || null;
      return {
        id: tx.ID,
        numero: tx.Numero,
        total: tx.Total,
        date: tx.Data instanceof Date ? Utilities.formatDate(tx.Data, tz, 'dd/MM/yyyy') : String(tx.Data),
        vencimento: tx.DataVencimento,
        valor: parseFloat(tx.Valor) || 0,
        valorOriginal: parseFloat(tx.ValorOriginal) || 0,
        juros: parseFloat(tx.Juros) || 0,
        status: tx.StatusParcela,
        dataPagamento: tx.DataPagamento,
        categoryName: cat ? cat.name : tx.CategoriaNomeSnapshot,
        categoryEmoji: cat ? cat.emoji : '',
        description: tx['Descrição']
      };
    });

    items.sort(function (a, b) { return a.numero - b.numero; });

    var pagas = items.filter(function (i) { return i.status === 'PAGA'; }).length;
    var pendentes = items.filter(function (i) { return i.status === 'PENDENTE'; }).length;
    var canceladas = items.filter(function (i) { return i.status === 'CANCELADA'; }).length;

    return okResponse({
      grupoId: grupoId,
      config: {
        tipo: grupoConfig.Tipo,
        frequencia: grupoConfig.Frequencia,
        valorTotal: parseFloat(grupoConfig.ValorTotal) || 0,
        numParcelas: parseInt(grupoConfig.NumParcelas) || 0
      },
      parcelas: items,
      stats: {
        total: items.length,
        pagas: pagas,
        pendentes: pendentes,
        canceladas: canceladas,
        progresso: items.length > 0 ? Math.round((pagas / items.length) * 100) : 0
      }
    });
  } catch (e) {
    return errorResponse(e.message, 'GET_DETAILS_ERROR');
  }
}

// ── Helper Functions ─────────────────────────────────────────────

function updateGrupoParcelasPagas(grupoId, increment) {
  try {
    var rowNum = findRowIndex(SHEET_PARCELAS_CONFIG, 1, grupoId);
    if (rowNum < 0) return;

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PARCELAS_CONFIG);
    var headers = HEADERS[SHEET_PARCELAS_CONFIG];
    var idxPagas = headers.indexOf('ParcelasPagas') + 1;
    var idxAtualizado = headers.indexOf('AtualizadoEm') + 1;

    var atual = sheet.getRange(rowNum, idxPagas).getValue();
    sheet.getRange(rowNum, idxPagas).setValue(parseInt(atual) + increment);
    sheet.getRange(rowNum, idxAtualizado).setValue(nowISO());
  } catch (e) {
    Logger.log('Erro ao atualizar grupo pagas: ' + e.message);
  }
}

function updateGrupoParcelasCanceladas(grupoId, increment) {
  try {
    var rowNum = findRowIndex(SHEET_PARCELAS_CONFIG, 1, grupoId);
    if (rowNum < 0) return;

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_PARCELAS_CONFIG);
    var headers = HEADERS[SHEET_PARCELAS_CONFIG];
    var idxCanceladas = headers.indexOf('ParcelasCanceladas') + 1;
    var idxAtualizado = headers.indexOf('AtualizadoEm') + 1;

    var atual = sheet.getRange(rowNum, idxCanceladas).getValue();
    sheet.getRange(rowNum, idxCanceladas).setValue(parseInt(atual) + increment);
    sheet.getRange(rowNum, idxAtualizado).setValue(nowISO());
  } catch (e) {
    Logger.log('Erro ao atualizar grupo canceladas: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// RECONCILIATION (Compensado/Pendente)
// ═══════════════════════════════════════════════════════════════════

/**
 * Marca uma transação como compensada (reconciliada com banco)
 */
function reconcileTransaction(payload) {
  try {
    var id = payload.id;
    var dataCompensacao = payload.dataCompensacao || nowISO().substring(0, 10);

    if (!id) {
      return errorResponse('ID da transação é obrigatório', 'MISSING_ID');
    }

    var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);
    if (rowNum < 0) {
      return errorResponse('Transação não encontrada', 'NOT_FOUND');
    }

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var headers = HEADERS[SHEET_LANCAMENTOS];

    var idxStatus = headers.indexOf('StatusReconciliacao') + 1;
    var idxData = headers.indexOf('DataCompensacao') + 1;
    var idxAtualizado = headers.indexOf('AtualizadoEm') + 1;
    var idxTipo = headers.indexOf('Tipo') + 1;
    var idxNotas = headers.indexOf('Notas') + 1;
    var tipo = idxTipo > 0 ? sheet.getRange(rowNum, idxTipo).getValue() : '';
    var notas = idxNotas > 0 ? sheet.getRange(rowNum, idxNotas).getValue() : '';
    if (isCardExpenseTx_(tipo, notas)) {
      return errorResponse('Lançamento de cartão não pode ser compensado aqui. Use o pagamento da fatura.', 'CARD_TX_RECONCILE_BLOCKED');
    }

    sheet.getRange(rowNum, idxStatus).setValue('COMPENSADO');
    sheet.getRange(rowNum, idxData).setValue(dataCompensacao);
    sheet.getRange(rowNum, idxAtualizado).setValue(nowISO());

    clearCache(CACHE_KEY_TRANSACTIONS);

    return okResponse({
      id: id,
      status: 'COMPENSADO',
      dataCompensacao: dataCompensacao
    });
  } catch (e) {
    return errorResponse(e.message, 'RECONCILE_ERROR');
  }
}

/**
 * Reconcilia múltiplas transações de uma vez
 */
function bulkReconcile(payload) {
  try {
    var ids = payload.ids || [];
    var dataCompensacao = payload.dataCompensacao || nowISO().substring(0, 10);

    if (!ids || ids.length === 0) {
      return errorResponse('Lista de IDs é obrigatória', 'MISSING_IDS');
    }

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LANCAMENTOS);
    var headers = HEADERS[SHEET_LANCAMENTOS];
    var idxStatus = headers.indexOf('StatusReconciliacao') + 1;
    var idxData = headers.indexOf('DataCompensacao') + 1;
    var idxAtualizado = headers.indexOf('AtualizadoEm') + 1;
    var idxTipo = headers.indexOf('Tipo') + 1;
    var idxNotas = headers.indexOf('Notas') + 1;

    var reconciled = 0;
    var skippedCard = [];
    ids.forEach(function (id) {
      var rowNum = findRowIndex(SHEET_LANCAMENTOS, 1, id);
      if (rowNum > 0) {
        var tipo = idxTipo > 0 ? sheet.getRange(rowNum, idxTipo).getValue() : '';
        var notas = idxNotas > 0 ? sheet.getRange(rowNum, idxNotas).getValue() : '';
        if (isCardExpenseTx_(tipo, notas)) {
          skippedCard.push(id);
          return;
        }
        sheet.getRange(rowNum, idxStatus).setValue('COMPENSADO');
        sheet.getRange(rowNum, idxData).setValue(dataCompensacao);
        sheet.getRange(rowNum, idxAtualizado).setValue(nowISO());
        reconciled++;
      }
    });

    clearCache(CACHE_KEY_TRANSACTIONS);

    return okResponse({
      reconciled: reconciled,
      total: ids.length,
      skippedCard: skippedCard
    });
  } catch (e) {
    return errorResponse(e.message, 'BULK_RECONCILE_ERROR');
  }
}

/**
 * Retorna resumo de reconciliação (saldos compensados vs pendentes)
 */
function getReconciliationSummary(params) {
  try {
    var month = params.month || new Date().getMonth();
    var year = params.year || new Date().getFullYear();

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var hoje = new Date().toISOString().substring(0, 10);

    var receitasCompensadas = 0;
    var receitasPendentes = 0;
    var despesasCompensadas = 0;
    var despesasPendentes = 0;
    var naoReconciliadasCount = 0;
    var pendentesList = [];

    allTx.forEach(function (tx) {
      var txDate = new Date(tx.Data + 'T00:00:00');

      // Filtrar por mês (opcional, pode ser removido para total geral)
      if (month !== null && year !== null) {
        if (txDate.getMonth() !== month || txDate.getFullYear() !== year) {
          return;
        }
      }

      var valor = parseFloat(tx.Valor) || 0;
      var isCompensado = tx.StatusReconciliacao === 'COMPENSADO';

      if (tx.Tipo === 'RECEITA') {
        if (isCompensado) {
          receitasCompensadas += valor;
        } else {
          receitasPendentes += valor;
        }
      } else if (tx.Tipo === 'DESPESA') {
        if (isCompensado) {
          despesasCompensadas += valor;
        } else {
          despesasPendentes += valor;
        }
      }

      // Contar não reconciliadas antigas
      if (!isCompensado && tx.Data < hoje) {
        naoReconciliadasCount++;
        pendentesList.push({
          id: tx.Id,
          date: tx.Data,
          description: tx.Descricao,
          amount: valor,
          type: tx.Tipo,
          category: tx.CategoriaNome
        });
      }
    });

    var saldoCompensado = receitasCompensadas - despesasCompensadas;
    var saldoPendente = receitasPendentes - despesasPendentes;
    var saldoTotal = saldoCompensado + saldoPendente;

    return okResponse({
      receitasCompensadas: receitasCompensadas,
      receitasPendentes: receitasPendentes,
      despesasCompensadas: despesasCompensadas,
      despesasPendentes: despesasPendentes,
      saldoCompensado: saldoCompensado,
      saldoPendente: saldoPendente,
      saldoTotal: saldoTotal,
      naoReconciliadasCount: naoReconciliadasCount,
      pendentesList: pendentesList.slice(0, 20) // Limitar a 20 para performance
    });
  } catch (e) {
    return errorResponse(e.message, 'RECONCILIATION_SUMMARY_ERROR');
  }
}
// ── Agrupar rotas de Contas ──────────────────────────────────────

/**
 * Retorna lista de contas com saldo atualizado
 */
function invalidateAccountsCaches_() {
  try {
    var cache = CacheService.getScriptCache();
    var keys = [
      'accounts', 'accounts_all', 'accounts_v2', 'accounts_all_v2', 'accounts_v3', 'accounts_all_v3'
    ];
    var scopeUser = (typeof AUTH_EXECUTION_USER_ID !== 'undefined' && AUTH_EXECUTION_USER_ID)
      ? String(AUTH_EXECUTION_USER_ID)
      : 'public';
    ['public', 'anon', scopeUser].forEach(function (suffix) {
      keys.push('accounts_' + suffix);
      keys.push('accounts_all_' + suffix);
      keys.push('accounts_v2_' + suffix);
      keys.push('accounts_all_v2_' + suffix);
      keys.push('accounts_v3_' + suffix);
      keys.push('accounts_all_v3_' + suffix);
    });
    cache.removeAll(keys);
  } catch (e) {
    // fallback best effort
    try {
      invalidateCache('accounts');
      invalidateCache('accounts_all');
      invalidateCache('accounts_v3');
      invalidateCache('accounts_all_v3');
    } catch (err) {}
  }
}

function listAccounts(params) {
  try {
    var props = PropertiesService.getScriptProperties();
    // Migração: garantir que a coluna Instituicao exista pelo NOME
    var ss = getDbSpreadsheet();
    var contasSheet = ss.getSheetByName(SHEET_CONTAS);
    if (contasSheet && contasSheet.getLastColumn() > 0) {
      var headerRange = contasSheet.getRange(1, 1, 1, contasSheet.getLastColumn());
      var currentHeaders = headerRange.getValues()[0];
      var idxInstituicao = currentHeaders.indexOf('Instituicao');

      if (idxInstituicao === -1) {
        // Adicionar no final
        var nextCol = currentHeaders.length + 1;
        contasSheet.getRange(1, nextCol).setValue('Instituicao').setFontWeight('bold');
        Logger.log('Migração CONTAS: Coluna Instituicao adicionada na coluna ' + nextCol);
        invalidateAccountsCaches_();
      }
    }
    props.setProperty('CONTAS_MIGRATED_V2', 'true');

    var includeInactive = params && params.includeInactive;
    var accounts = getAccountsCached(!!includeInactive);

    // Calcular saldo atual de cada conta
    // Idealmente, cachear isso ou calcular em batch se ficar lento
    var accountsWithBalance = accounts.map(function (acc) {
      acc.saldoAtual = calculateAccountBalance(acc.id);
      return acc;
    });

    return okResponse(accountsWithBalance);
  } catch (e) {
    return errorResponse(e.message, 'LIST_ACCOUNTS_ERROR');
  }
}

function createAccount(payload) {
  try {
    if (!payload.nome || !payload.nome.trim())
      return errorResponse('Nome da conta é obrigatório.', 'VALIDATION');

    var tipo = payload.tipo || 'Outro';
    var saldoInicial = parseFloat(payload.saldoInicial) || 0;
    var instituicao = payload.instituicao || 'Outro';

    // Encontrar próxima ordem
    var existing = getAccountsCached(true);
    var maxOrder = 0;
    existing.forEach(function (a) { if (a.ordem > maxOrder) maxOrder = a.ordem; });

    var id = Utilities.getUuid();
    var now = nowISO();

    var row = [
      id,
      payload.nome.trim(),
      tipo,
      saldoInicial,
      true, // Ativo
      maxOrder + 1,
      now,
      now,
      instituicao
    ];

    appendRow(SHEET_CONTAS, row);
    invalidateAccountsCaches_();

    return okResponse({ id: id });
  } catch (e) {
    return errorResponse(e.message, 'CREATE_ACCOUNT_ERROR');
  }
}

function updateAccount(payload) {
  try {
    if (!payload.id) return errorResponse('ID é obrigatório.', 'VALIDATION');

    var rowNum = findRowIndex(SHEET_CONTAS, 1, payload.id);
    if (rowNum < 0) return errorResponse('Conta não encontrada.', 'NOT_FOUND');

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CONTAS);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idxInstituicao = headers.indexOf('Instituicao');

    // Se a coluna não existir, usamos o fallback do HEADERS original (coluna 9)
    var colInstituicao = (idxInstituicao !== -1) ? (idxInstituicao + 1) : 9;

    var existingRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    var existingInstituicao = (idxInstituicao !== -1) ? existingRow[idxInstituicao] : 'Outro';

    Logger.log('UPDATE_ACCOUNT: id=' + payload.id + ', instituicao_recebida=' + payload.instituicao + ', existente=' + existingInstituicao);

    var updatedRow = [];
    var expectedHeaders = HEADERS[SHEET_CONTAS];

    // Mapear cada header esperado para o valor apropriado
    expectedHeaders.forEach(function (h, i) {
      if (h === 'ID') updatedRow.push(payload.id);
      else if (h === 'Nome') updatedRow.push(payload.nome !== undefined ? payload.nome.trim() : existingRow[headers.indexOf('Nome')]);
      else if (h === 'Tipo') updatedRow.push(payload.tipo !== undefined ? payload.tipo : existingRow[headers.indexOf('Tipo')]);
      else if (h === 'SaldoInicial') updatedRow.push(payload.saldoInicial !== undefined ? parseFloat(payload.saldoInicial) : existingRow[headers.indexOf('SaldoInicial')]);
      else if (h === 'Ativo') updatedRow.push(payload.ativo !== undefined ? payload.ativo : existingRow[headers.indexOf('Ativo')]);
      else if (h === 'Ordem') updatedRow.push(existingRow[headers.indexOf('Ordem')]);
      else if (h === 'CriadoEm') updatedRow.push(existingRow[headers.indexOf('CriadoEm')]);
      else if (h === 'AtualizadoEm') updatedRow.push(nowISO());
      else if (h === 'Instituicao') updatedRow.push(payload.instituicao !== undefined ? payload.instituicao : (existingInstituicao || 'Outro'));
      else updatedRow.push(existingRow[i] || '');
    });

    updateRow(SHEET_CONTAS, rowNum, updatedRow);

    // Invalida ambos os caches (v2 e v3 para garantir)
    invalidateAccountsCaches_();

    Logger.log('UPDATE_ACCOUNT: Sucesso. Caches invalidados.');

    return okResponse({ id: payload.id });
  } catch (e) {
    return errorResponse(e.message, 'UPDATE_ACCOUNT_ERROR');
  }
}

function deactivateAccount(id) {
  try {
    if (!id) return errorResponse('ID é obrigatório.', 'VALIDATION');

    var rowNum = findRowIndex(SHEET_CONTAS, 1, id);
    if (rowNum < 0) return errorResponse('Conta não encontrada.', 'NOT_FOUND');

    // Verificar se tem transações
    var hasTx = false;
    try {
      hasTx = accountHasTransactions(id);
    } catch (e) { hasTx = true; } // Fallback to safe soft-delete if check fails

    if (!hasTx) {
      // Hard Delete: remove a linha fisicamente se nunca foi usada
      deleteRow(SHEET_CONTAS, rowNum);
      invalidateAccountsCaches_();
      return okResponse({ deactivated: true, method: 'hard_delete' });
    }

    // Soft Delete: marca como inativa
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_CONTAS);

    // Ler a linha toda para preservar dados
    var existing = sheet.getRange(rowNum, 1, 1, HEADERS[SHEET_CONTAS].length).getValues()[0];

    var updatedRow = existing.slice(); // Clone
    updatedRow[4] = false; // Ativo = false
    updatedRow[7] = nowISO(); // AtualizadoEm

    updateRow(SHEET_CONTAS, rowNum, updatedRow);
    invalidateAccountsCaches_();

    return okResponse({ deactivated: true, method: 'soft_delete' });
  } catch (e) {
    return errorResponse(e.message, 'DEACTIVATE_ACCOUNT_ERROR');
  }
}

function clearAllServerCache() {
  try {
    var cache = CacheService.getScriptCache();
    // Lista exaustiva de chaves conhecidas
    var keys = [
      'accounts', 'accounts_all', 'accounts_v2', 'accounts_all_v2', 'accounts_v3', 'accounts_all_v3',
      'categories', 'categories_all', 'config', 'dashboard', 'transactions', 'installmentsDash'
    ];
    cache.removeAll(keys);
    return okResponse({ message: 'All caches cleared' });
  } catch (e) {
    return errorResponse(e.message, 'CLEAR_CACHE_ERROR');
  }
}

function getAccountBalance(id) {
  try {
    var balance = calculateAccountBalance(id);
    return okResponse({ balance: balance });
  } catch (e) {
    return errorResponse(e.message, 'GET_BALANCE_ERROR');
  }
}


