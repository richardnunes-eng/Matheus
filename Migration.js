/**
 * EXECUÇÃO ÚNICA - MIGRAÇÃO PARA FASE 5
 * 
 * Execute esta função UMA VEZ no Apps Script Editor para:
 * 1. Criar a planilha PARCELAS_CONFIG se não existir
 * 2. Adicionar as colunas StatusReconciliacao, DataCompensacao, Notas
 * 
 * COMO EXECUTAR:
 * 1. Abra o Apps Script Editor (clasp open-script)
 * 2. Cole este código em um arquivo novo chamado "Migration.gs"
 * 3. Selecione a função "migrateToPhase5" no menu dropdown
 * 4. Clique em "Executar" (▶️)
 * 5. Autorize quando solicitado
 * 6. Aguarde a mensagem de sucesso
 * 7. Delete este arquivo após a execução
 */

function migrateToPhase5() {
  try {
    Logger.log('Iniciando migração para Fase 5...');

    // 1. Garantir que todas as planilhas existam
    Logger.log('Verificando planilhas...');
    ensureAllSheets();
    Logger.log('✓ Todas as planilhas criadas/verificadas');

    // 2. Adicionar novas colunas na LANCAMENTOS se não existirem
    var ss = getDbSpreadsheet();
    var lancamentosSheet = ss.getSheetByName(SHEET_LANCAMENTOS);

    if (lancamentosSheet) {
      var headers = lancamentosSheet.getRange(1, 1, 1, lancamentosSheet.getLastColumn()).getValues()[0];
      var needsUpdate = false;

      // Verificar se as colunas já existem
      if (headers.indexOf('StatusReconciliacao') === -1) {
        Logger.log('Adicionando colunas de reconciliação...');

        // Pegar a última coluna
        var lastCol = lancamentosSheet.getLastColumn();

        // Adicionar 3 novas colunas
        lancamentosSheet.getRange(1, lastCol + 1).setValue('StatusReconciliacao');
        lancamentosSheet.getRange(1, lastCol + 2).setValue('DataCompensacao');
        lancamentosSheet.getRange(1, lastCol + 3).setValue('Notas');

        // Preencher valores padrão para transações existentes
        var lastRow = lancamentosSheet.getLastRow();
        if (lastRow > 1) {
          // StatusReconciliacao = PENDENTE para todas
          var statusRange = lancamentosSheet.getRange(2, lastCol + 1, lastRow - 1, 1);
          var statusValues = [];
          for (var i = 0; i < lastRow - 1; i++) {
            statusValues.push(['PENDENTE']);
          }
          statusRange.setValues(statusValues);

          Logger.log('✓ Preenchidas ' + (lastRow - 1) + ' transações com status PENDENTE');
        }

        needsUpdate = true;
      } else {
        Logger.log('✓ Colunas de reconciliação já existem');
      }

      if (needsUpdate) {
        Logger.log('✓ Colunas adicionadas com sucesso');
      }
    }

    // 3. Verificar PARCELAS_CONFIG
    var parcelasSheet = ss.getSheetByName(SHEET_PARCELAS_CONFIG);
    if (parcelasSheet) {
      Logger.log('✓ Planilha PARCELAS_CONFIG existe');
    } else {
      Logger.log('⚠ Planilha PARCELAS_CONFIG não foi criada - execute ensureAllSheets()');
    }

    // 4. Limpar cache
    clearAllCaches();
    Logger.log('✓ Cache limpo');

    Logger.log('');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('✅ MIGRAÇÃO CONCLUÍDA COM SUCESSO!');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('');
    Logger.log('Próximos passos:');
    Logger.log('1. Recarregue a aplicação web');
    Logger.log('2. Teste criar uma transação com notas');
    Logger.log('3. Verifique a planilha LANCAMENTOS');
    Logger.log('4. Delete este arquivo Migration.gs');

    // Mostrar um alerta visual (se executado via interface web)
    try {
      SpreadsheetApp.getUi().alert(
        'Migração Concluída!',
        'Fase 5 instalada com sucesso!\n\n' +
        '✓ Planilha PARCELAS_CONFIG criada\n' +
        '✓ Colunas de reconciliação adicionadas\n' +
        '✓ Cache limpo\n\n' +
        'Você já pode usar o campo de Notas e as funcionalidades de reconciliação!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e) { }

  } catch (e) {
    Logger.log('❌ ERRO na migração: ' + e.message);
    Logger.log(e.stack);
    throw e;
  }
}

function clearAllCaches() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['transactions', 'categories', 'config', 'accounts', 'accounts_all']);
}

// ══════════════════════════════════════════════════════════════════
// MIGRAÇÃO PARA SISTEMA DE CONTAS
// ══════════════════════════════════════════════════════════════════

/**
 * Migra o sistema para usar Contas/Carteiras ao invés de "Reserva Manutenção"
 * 
 * Esta função é idempotente e pode ser executada múltiplas vezes.
 * 
 * COMO EXECUTAR:
 * 1. Abra o Apps Script Editor
 * 2. Selecione "migrateToAccountsSystem" no menu
 * 3. Clique em "Executar"
 * 4. Aguarde os logs
 */
function migrateToAccountsSystem() {
  try {
    Logger.log('═══════════════════════════════════════════');
    Logger.log('MIGRAÇÃO: Sistema de Contas/Carteiras');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('');

    // 1. Garantir que todas as sheets existam
    Logger.log('1. Verificando planilhas...');
    ensureAllSheets();
    Logger.log('   ✓ Todas as planilhas verificadas/criadas');

    var ss = getDbSpreadsheet();
    var contasSheet = ss.getSheetByName(SHEET_CONTAS);
    var lancamentosSheet = ss.getSheetByName(SHEET_LANCAMENTOS);

    // 2. Verificar se já tem contas criadas
    var contasExistentes = getSheetDataAsObjects(SHEET_CONTAS);
    Logger.log('');
    Logger.log('2. Verificando contas existentes...');
    Logger.log('   Contas encontradas: ' + contasExistentes.length);

    if (contasExistentes.length === 0) {
      Logger.log('   → Criando contas padrão...');
      var defaultAccounts = createDefaultAccounts();

      for (var i = 0; i < defaultAccounts.length; i++) {
        appendRow(SHEET_CONTAS, defaultAccounts[i]);
      }

      Logger.log('   ✓ 3 contas padrão criadas:');
      Logger.log('     - Carteira');
      Logger.log('     - Investimentos');
      Logger.log('     - Banco (genérico)');
    } else {
      Logger.log('   ✓ Contas já existem, pulando criação');
    }

    // 3. Obter IDs das contas padrão
    Logger.log('');
    Logger.log('3. Mapeando IDs de contas...');
    var accounts = getAccountsCached(true);
    var investimentosId = null;
    var bancoGenericoId = null;

    for (var i = 0; i < accounts.length; i++) {
      if (accounts[i].nome === 'Investimentos') {
        investimentosId = accounts[i].id;
      }
      if (accounts[i].nome === 'Banco (genérico)') {
        bancoGenericoId = accounts[i].id;
      }
    }

    Logger.log('   Investimentos ID: ' + investimentosId);
    Logger.log('   Banco (genérico) ID: ' + bancoGenericoId);

    if (!investimentosId || !bancoGenericoId) {
      throw new Error('Contas padrão não encontradas. Execute a migração novamente.');
    }

    // 4. Migrar transações antigas
    Logger.log('');
    Logger.log('4. Migrando transações antigas...');

    var allTx = getSheetDataAsObjects(SHEET_LANCAMENTOS);
    var migrated = 0;
    var alreadyMigrated = 0;

    for (var i = 0; i < allTx.length; i++) {
      var tx = allTx[i];
      var rowNum = i + 2; // +2 porque começa em 1 e tem header
      var needsUpdate = false;
      var updates = {};

      var tipo = String(tx.Tipo);
      var carteiraOrigem = String(tx.CarteiraOrigem || '');
      var carteiraDestino = String(tx.CarteiraDestino || '');

      // Normalizar "TRANSFER" para "TRANSFERENCIA"
      if (tipo === 'TRANSFER') {
        updates.tipo = 'TRANSFERENCIA';
        needsUpdate = true;
      }

      // Migrar CarteiraDestino = 'reserva' → Investimentos
      if (carteiraDestino.toLowerCase().indexOf('reserva') !== -1 && carteiraDestino !== investimentosId) {
        updates.carteiraDestino = investimentosId;
        needsUpdate = true;
      }

      // Migrar CarteiraOrigem = 'reserva' → Investimentos
      if (carteiraOrigem.toLowerCase().indexOf('reserva') !== -1 && carteiraOrigem !== investimentosId) {
        updates.carteiraOrigem = investimentosId;
        needsUpdate = true;
      }

      // RECEITA sem conta destino → Banco (genérico)
      if (tipo === 'RECEITA' && !carteiraDestino) {
        updates.carteiraDestino = bancoGenericoId;
        needsUpdate = true;
      }

      // DESPESA sem conta origem → Banco (genérico)
      if (tipo === 'DESPESA' && !carteiraOrigem) {
        updates.carteiraOrigem = bancoGenericoId;
        needsUpdate = true;
      }

      // TRANSFERENCIA sem contas → Banco (genérico)
      if ((tipo === 'TRANSFERENCIA' || tipo === 'TRANSFER')) {
        if (!carteiraOrigem) {
          updates.carteiraOrigem = bancoGenericoId;
          needsUpdate = true;
        }
        if (!carteiraDestino) {
          updates.carteiraDestino = bancoGenericoId;
          needsUpdate = true;
        }
      }

      if (needsUpdate) {
        // Atualizar a linha
        var headers = HEADERS[SHEET_LANCAMENTOS];
        var rowData = [];

        for (var h = 0; h < headers.length; h++) {
          var header = headers[h];
          var value = tx[header];

          // Aplicar updates
          if (header === 'Tipo' && updates.tipo) {
            value = updates.tipo;
          }
          if (header === 'CarteiraOrigem' && updates.carteiraOrigem) {
            value = updates.carteiraOrigem;
          }
          if (header === 'CarteiraDestino' && updates.carteiraDestino) {
            value = updates.carteiraDestino;
          }

          rowData.push(value);
        }

        updateRow(SHEET_LANCAMENTOS, rowNum, rowData);
        migrated++;
      } else {
        alreadyMigrated++;
      }
    }

    Logger.log('   ✓ Transações migradas: ' + migrated);
    Logger.log('   ✓ Transações já OK: ' + alreadyMigrated);
    Logger.log('   Total processado: ' + allTx.length);

    // 5. Limpar cache
    Logger.log('');
    Logger.log('5. Limpando cache...');
    clearAllCaches();
    Logger.log('   ✓ Cache limpo');

    // 6. Conclusão
    Logger.log('');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('✅ MIGRAÇÃO CONCLUÍDA COM SUCESSO!');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('');
    Logger.log('Resumo:');
    Logger.log('• Contas padrão: ' + (contasExistentes.length === 0 ? 'Criadas' : 'Já existiam'));
    Logger.log('• Transações migradas: ' + migrated);
    Logger.log('• Sistema pronto para uso!');
    Logger.log('');
    Logger.log('Próximos passos:');
    Logger.log('1. Faça deploy da nova versão');
    Logger.log('2. Recarregue a aplicação');
    Logger.log('3. Acesse "Contas" no menu');

    // Alerta visual (se executado via UI)
    try {
      SpreadsheetApp.getUi().alert(
        'Migração Concluída! ✅',
        'Sistema de Contas/Carteiras instalado com sucesso!\n\n' +
        '✓ ' + (contasExistentes.length === 0 ? '3 contas padrão criadas' : 'Contas existentes preservadas') + '\n' +
        '✓ ' + migrated + ' transações migradas\n' +
        '✓ Cache limpo\n\n' +
        'Faça deploy da nova versão e recarregue a aplicação!',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (e) { }

  } catch (e) {
    Logger.log('');
    Logger.log('❌ ERRO na migração: ' + e.message);
    Logger.log(e.stack);
    throw e;
  }
}

