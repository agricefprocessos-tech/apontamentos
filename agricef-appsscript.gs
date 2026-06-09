// ================================================================
// AGRICEF — Web App Apps Script v4.2 (deploy 14/05/2026)
//
// Colunas da planilha (inalteradas):
//   A  Carimbo de data/hora
//   B  NOME DO OPERADOR
//   C  TIPO DE APONTAMENTO
//   D  TIPO DE OPERAÇÃO 1
//   E  CÓDIGO DO ITEM
//   F  Nº SERIE | IMPLEMENTO | CLIENTE | INTERNO
//   G  QUANTIDADE
//   H  OBSERVAÇÃO 1
//   I  TIPOS DE RETRABALHOS
//   J  Nº DA RNC
//   K  TIPOS DE PARADA
//   L  TIPO DE OPERAÇÃO 2
//   M  TIPO DE OPERAÇÃO DE SET-UP
//   N  OBSERVAÇÃO 2
//   O  OBSERVAÇÃO  ← usado para QTD PLANEJADA
// ================================================================

const SPREADSHEET_ID  = '15vtJ2eOw3Zd9f5MmwqEj18nsGAvVkFYFpsUsRbZM6Ik';
const ABA_RESPOSTAS   = 'Respostas do Formulário 1';
const ABA_ABERTOS     = 'Abertos';
const ABA_OPERADORES  = 'Cadastro_Operadores';
const ABA_SALDO       = 'Saldo_Parcial';
const ABA_SERIES      = 'Cadastro_Series';
const ABA_OPERACOES   = 'Cadastro_Operacoes';

const TIPOS_APONTAMENTO = {
  'ABERTURA':           'ABERTURA',
  'FECHAMENTO':         'FECHAMENTO',
  'INICIO_RETRABALHO':  'INÍCIO DE RETRABALHO',
  'TERMINO_RETRABALHO': 'TÉRMINO DE RETRABALHO',
  'INICIO_PARADA':      'INÍCIO DE PARADA',
  'TERMINO_PARADA':     'TÉRMINO DE PARADA',
};

const OPERACOES = {
  '0010': '0010 - CORTAR',
  '0020': '0020 - FURAR',
  '0030': '0030 - MONTAR CALDEIRARIA',
  '0040': '0040 - SOLDAR',
  '0050': '0050 - LIXAR',
  '0060': '0060 - PINTAR',
  '0070': '0070 - MONTAR MECÂNICA | ELÉTRICA',
  '0080': '0080 - INSPECIONAR',
  '0090': '0090 - CALAFETAR',
  '0100': '0100 - ADESIVAR',
  '0101': '0101 - TESTES',
  '0102': '0102 - TREINAMENTO',
};

// ================================================================
// ENTRY POINTS
// ================================================================

function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'verificarAberto')  return verificarAberto(e.parameter.operador, e.parameter.implemento);
  if (action === 'verificarSaldo')    return verificarSaldoParcialAction(e.parameter.nrSerie, e.parameter.codItem || '', e.parameter.operacao);
  if (action === 'getCadastros')     return getCadastros();
  if (action === 'getAbertos')       return getAbertosAction();
  if (action === 'getData')          return getDadosRespostas();
  if (action === 'triggerRelatorio' && e.parameter.key === 'AGF2026') {
    try {
      enviarRelatorioSemanal();
      return jsonResponse({ success: true, message: 'Relatório semanal enviado para ' + EMAIL_RELATORIO });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'triggerDiario' && e.parameter.key === 'AGF2026') {
    try {
      enviarRelatorioDiario();
      return jsonResponse({ success: true, message: 'Relatório diário enviado para ' + EMAIL_RELATORIO });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'ativarTrigger' && e.parameter.key === 'AGF2026') {
    try {
      criarTriggerRelatorioSemanal();
      return jsonResponse({ success: true, message: 'Trigger semanal criado com sucesso.' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'ativarTriggerDiario' && e.parameter.key === 'AGF2026') {
    try {
      criarTriggerRelatorioDiario();
      return jsonResponse({ success: true, message: 'Trigger diário criado (seg–sex 07h).' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'ativarTodosTriggers' && e.parameter.key === 'AGF2026') {
    try {
      criarTriggerRelatorioSemanal();
      criarTriggerRelatorioDiario();
      return jsonResponse({ success: true, message: 'Triggers semanal e diário criados com sucesso.' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'reconstruirAbertos' && e.parameter.key === 'AGF2026') {
    try {
      const result = reconstruirAbertos();
      return jsonResponse({ success: true, ...result });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'ativarTriggerReconciliacao' && e.parameter.key === 'AGF2026') {
    try {
      criarTriggerReconciliacaoAbertos();
      return jsonResponse({ success: true, message: 'Trigger de reconciliação criado: reconstruirAbertos a cada 30 minutos.' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'normalizarOperadores' && e.parameter.key === 'AGF2026') {
    try {
      const total = normalizarCodigosOperador();
      return jsonResponse({ success: true, message: total + ' código(s) normalizado(s) com sucesso.' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'normalizarRespostas' && e.parameter.key === 'AGF2026') {
    try {
      const total = normalizarRespostasColB();
      return jsonResponse({ success: true, message: total + ' registro(s) da aba Respostas normalizado(s).' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'normalizarDatas' && e.parameter.key === 'AGF2026') {
    try {
      const total = normalizarDatasColA();
      return jsonResponse({ success: true, message: total + ' data(s) normalizada(s) na coluna A.' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'normalizarTodos' && e.parameter.key === 'AGF2026') {
    try {
      const result = normalizarTodosDados();
      return jsonResponse({ success: true, message: result.datas + ' data(s) + ' + result.operadores + ' operador(es) normalizados. Total: ' + result.total, detail: result });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'limparTestes' && e.parameter.key === 'AGF2026') {
    try {
      const total = limparRegistrosTeste();
      return jsonResponse({ success: true, message: total + ' linha(s) de teste removida(s).' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'removerSemTimestamp' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(removerRegistrosSemTimestamp());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'reverterTimestamps' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(reverterTimestampsInterpolados());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'preencherTimestamps' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(preencherTimestampsVazios());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'listarRevisoes' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(listarRevisoesPlanilha());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'recuperarTimestamps' && e.parameter.key === 'AGF2026') {
    try {
      const revId = e.parameter.revId || '';
      if (!revId) return jsonResponse({ success: false, message: 'Parâmetro revId obrigatório.' });
      return jsonResponse(recuperarTimestampsDaRevisao(revId));
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'aplicarTimestamps' && e.parameter.key === 'AGF2026') {
    try {
      const revId = e.parameter.revId || '';
      if (!revId) return jsonResponse({ success: false, message: 'Parâmetro revId obrigatório.' });
      return jsonResponse(aplicarTimestampsDaRevisao(revId));
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'removerRegistroPorId' && e.parameter.key === 'AGF2026') {
    try {
      const idAlvo = e.parameter.id || '';
      if (!idAlvo) return jsonResponse({ success: false, message: 'Parâmetro id obrigatório.' });
      const resultado = removerRegistroPorAbertoId(idAlvo);
      return jsonResponse({ success: true, ...resultado });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'diagRespostas' && e.parameter.key === 'AGF2026') {
    try {
      const ss2 = SpreadsheetApp.openById(SPREADSHEET_ID);
      const aba2 = ss2.getSheetByName(ABA_RESPOSTAS);
      const lastRow = aba2.getLastRow();
      const maxRows = aba2.getMaxRows();
      const dados2 = aba2.getRange(Math.max(1, lastRow - 4), 1, 5, 16).getValues();
      const rows = dados2.map((r, i) => ({
        sheetRow: lastRow - 4 + i,
        colA: String(r[0] || '').substring(0, 30),
        colB: String(r[1] || '').substring(0, 30),
        colC: String(r[2] || ''),
        colH: String(r[7] || ''),
        colP: String(r[15] || '')
      }));
      return jsonResponse({ lastRow, maxRows, lastRows: rows });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'analyzeOrphans' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(analisarOrfaos());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'analisarInconsistencias' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(analisarInconsistencias());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'impactoOrfaos' && e.parameter.key === 'AGF2026') {
    try {
      return jsonResponse(impactoOrfaos());
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'marcarLegado' && e.parameter.key === 'AGF2026') {
    try {
      const cutoff = e.parameter.cutoff || '';
      const result = marcarOrfaosLegado(cutoff);
      return jsonResponse(result);
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'previewAberturasOrfas' && e.parameter.key === 'AGF2026') {
    try {
      const cutoff = e.parameter.cutoff || '2026-05-29';
      return jsonResponse(previewAberturasOrfas(cutoff));
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }
  if (action === 'deletarAberturasOrfas' && e.parameter.key === 'AGF2026') {
    try {
      const cutoff = e.parameter.cutoff || '2026-05-29';
      return jsonResponse(deletarAberturasOrfas(cutoff));
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }

  // Ações via payload GET (contorna CORS)
  if (e.parameter.payload) {
    try {
      const payload = JSON.parse(e.parameter.payload);
      if (payload.action === 'salvarOperador')   return salvarOperador(payload);
      if (payload.action === 'removerOperador')  return removerOperador(payload);
      if (payload.action === 'salvarSerie')      return salvarSerie(payload);
      if (payload.action === 'removerSerie')     return removerSerie(payload);
      if (payload.action === 'salvarOperacao')   return salvarOperacao(payload);
      if (payload.action === 'removerOperacao')  return removerOperacao(payload);
      // Localizar linhas por identificadores (func+ts) para deleção
      if (payload.action === 'localizarLinhas' && payload.key === 'AGF2026') {
        return jsonResponse(localizarLinhasPorIdentificadores(payload.identifiers || [], payload.deletar === true));
      }
      // Deletar linhas por número (mais preciso que func+ts)
      if (payload.action === 'deletarPorLinhas' && payload.key === 'AGF2026') {
        return jsonResponse(deletarLinhasPorNumero(payload.rows || [], payload.dryRun === true, payload.tiposPermitidos || null));
      }
      // Endpoint dedicado: deletar FECHAMENTOs sem ABERTURA (aceita apenas tipo FECHAMENTO)
      if (payload.action === 'deletarFechamentosOrfaos' && payload.key === 'AGF2026') {
        return jsonResponse(deletarLinhasPorNumero(payload.rows || [], payload.dryRun === true, ['FECHAMENTO']));
      }
      return gravarApontamento(payload);
    } catch (err) {
      return jsonResponse({ success: false, message: 'Erro ao processar payload: ' + err.message });
    }
  }

  return jsonResponse({ status: 'ok', message: 'AGRICEF Web App v4 ativo' });
}

function doPost(e) {
  try {
    let payload;
    if (e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    } else {
      payload = JSON.parse(e.postData.contents);
    }
    if (payload.action === 'salvarOperador')   return salvarOperador(payload);
    if (payload.action === 'removerOperador')  return removerOperador(payload);
    if (payload.action === 'salvarSerie')      return salvarSerie(payload);
    if (payload.action === 'removerSerie')     return removerSerie(payload);
    if (payload.action === 'salvarOperacao')   return salvarOperacao(payload);
    if (payload.action === 'removerOperacao')  return removerOperacao(payload);
    if (payload.action === 'localizarLinhas' && payload.key === 'AGF2026') {
      return jsonResponse(localizarLinhasPorIdentificadores(payload.identifiers || [], payload.deletar === true));
    }
    if (payload.action === 'deletarPorLinhas' && payload.key === 'AGF2026') {
      return jsonResponse(deletarLinhasPorNumero(payload.rows || [], payload.dryRun === true, payload.tiposPermitidos || null));
    }
    if (payload.action === 'deletarFechamentosOrfaos' && payload.key === 'AGF2026') {
      return jsonResponse(deletarLinhasPorNumero(payload.rows || [], payload.dryRun === true, ['FECHAMENTO']));
    }
    // Fix#ID-LINK: endpoint de migração histórica — escreve abertoId em linhas específicas
    if (payload.action === 'migrarIds' && payload.key === 'AGF2026') {
      return jsonResponse(migrarIdsHistoricos(payload.updates || []));
    }
    return gravarApontamento(payload);
  } catch (err) {
    return jsonResponse({ success: false, message: 'Erro: ' + err.message });
  }
}

// ================================================================
// VERIFICAR APONTAMENTO ABERTO — retorna dados completos para
// o app pré-preencher o fechamento
// ================================================================

function verificarAberto(operador, implemento) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba   = garantirAbaAbertos(ss);
    const dados = aba.getDataRange().getValues();

    // Colunas aba Abertos:
    // 0:Operador 1:Implemento 2:Tipo 3:Operação 4:Carimbo
    // 5:CodItem  6:QtdPlanejada 7:NrSerie 8:Implemento 9:Cliente 10:OperadorNome

    for (let i = 1; i < dados.length; i++) {
      const row   = dados[i];
      if (!row[0]) continue; // pula linha completamente vazia
      if (mesmoOperador(row[0], operador)) {

        // Bug#29/30/31: a cross-validação foi removida daqui.
        // Motivo: verificarAberto é uma operação de LEITURA — fazer
        // deleções aqui causava condição de corrida com gravarApontamento
        // e apagava registros legítimos quando os IDs ainda não estavam
        // sincronizados (pré-Bug#30) ou quando reconstruirAbertos usava
        // o ID da coluna P de Respostas como abertoId de Abertos.
        //
        // A detecção e remoção de registros fantasmas é feita por:
        //   1. gravarApontamento — cross-valida antes de bloquear por poka-yoke
        //   2. reconstruirAbertos — trigger periódico de 30 min reconstrói
        //      Abertos inteiramente a partir de Respostas
        //
        // Isso garante que: (a) operações de leitura não têm efeitos colaterais,
        // (b) limpeza acontece no momento correto (escrita ou ciclo periódico).

        let loteSeries = null;
        if (row[11]) {
          try { loteSeries = JSON.parse(String(row[11])); } catch(e) {}
        }
        return jsonResponse({
          aberto:        true,
          tipo:          String(row[2]  || ''),
          operacao:      String(row[3]  || ''),  // Bug#24 fix: Sheets retorna number para células numéricas
          carimbo:       formatarCarimboGs(row[4]),
          codItem:       String(row[5]  || ''),
          qtdPlanejada:  row[6]  || '',          // mantém número — frontend usa aritmética
          nrSerie:       String(row[7]  || ''),  // Bug#24b fix: nrSerie numérica (ex: 22000084)
          implemento:    String(row[8]  || ''),
          cliente:       String(row[9]  || ''),
          operadorNome:  String(row[10] || ''),
          loteSeries:    loteSeries,
          abertoId:      String(row[12] || ''),  // ID único gerado na abertura
        });
      }
    }
    return jsonResponse({ aberto: false });
  } catch (err) {
    return jsonResponse({ aberto: false, erro: err.message });
  }
}

// ================================================================
// GRAVAR APONTAMENTO
// ================================================================

function gravarApontamento(payload) {
  // LockService garante que duas requisições simultâneas não passem juntas pela validação
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (lockErr) {
    return jsonResponse({ success: false, message: 'Servidor ocupado. Tente novamente em instantes.' });
  }
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const abaRe = ss.getSheetByName(ABA_RESPOSTAS);
    if (!abaRe) return jsonResponse({ success: false, message: 'Aba "' + ABA_RESPOSTAS + '" não encontrada.' });

    const abaAb = garantirAbaAbertos(ss);
    const tipo  = payload.tipoApontamento || '';

    // ---------------------------------------------------------------
    // VALIDAÇÃO DE CAMPOS OBRIGATÓRIOS (Bug#16, Bug#17, Bug#18)
    // ---------------------------------------------------------------
    if (!tipo) {
      return jsonResponse({ success: false, message: 'Campo tipoApontamento é obrigatório.' });
    }
    if (!TIPOS_APONTAMENTO[tipo]) {
      return jsonResponse({ success: false, message: 'Tipo de apontamento inválido: "' + tipo + '". Use: ' + Object.keys(TIPOS_APONTAMENTO).join(', ') });
    }
    if (!payload.operador || String(payload.operador).trim() === '') {
      return jsonResponse({ success: false, message: 'Campo operador é obrigatório.' });
    }

    // ---------------------------------------------------------------
    // VALIDAÇÕES DE CAMPOS ESPECÍFICOS POR TIPO (Bug#21, #22, #23)
    // ---------------------------------------------------------------
    const _tiposAb = ['ABERTURA', 'INICIO_RETRABALHO', 'INICIO_PARADA'];

    // Bug#22 fix: nrSerie obrigatório para tipos de abertura
    if (_tiposAb.includes(tipo) && (!payload.nrSerie || String(payload.nrSerie).trim() === '')) {
      return jsonResponse({ success: false, message: 'Campo nrSerie é obrigatório para ' + tipo + '.' });
    }
    // Bug#28 fix: nrSerie não pode conter "|" (separador do campo F da planilha)
    if (payload.nrSerie && String(payload.nrSerie).includes('|')) {
      return jsonResponse({ success: false, message: 'Campo nrSerie não pode conter o caractere "|" (usado como separador interno).' });
    }

    // Bug#23 fix: motivo obrigatório para INICIO_RETRABALHO e INICIO_PARADA
    if (tipo === 'INICIO_RETRABALHO' && (!payload.retrabalho || String(payload.retrabalho).trim() === '')) {
      return jsonResponse({ success: false, message: 'Campo retrabalho (motivo do retrabalho) é obrigatório para INICIO_RETRABALHO.' });
    }
    if (tipo === 'INICIO_PARADA' && (!payload.parada || String(payload.parada).trim() === '')) {
      return jsonResponse({ success: false, message: 'Campo parada (tipo de parada) é obrigatório para INICIO_PARADA.' });
    }

    // Bug#21 fix: LOTE requer ao menos uma série em loteSeries
    if (payload.isLote === true || (payload.loteSeries && Array.isArray(payload.loteSeries))) {
      if (!payload.loteSeries || !Array.isArray(payload.loteSeries) || payload.loteSeries.length === 0) {
        return jsonResponse({ success: false, message: 'Apontamento em lote requer pelo menos uma série em loteSeries.' });
      }
    }

    // Lê Abertos UMA VEZ — reaproveitado em validação, loteSeriesFechamento e atualizarAbertos
    const dadosAbertos = abaAb.getDataRange().getValues();

    // ---------------------------------------------------------------
    // VALIDAÇÃO SERVER-SIDE — executada ANTES de qualquer escrita,
    // dentro do lock, para bloquear dupla abertura com segurança.
    // ---------------------------------------------------------------
    const tiposAbertura = ['ABERTURA', 'INICIO_RETRABALHO', 'INICIO_PARADA'];

    // Mapeamento: tipo de fechamento → tipo de abertura esperado na aba Abertos
    const TIPO_COMPATIVEL = {
      'FECHAMENTO':         'ABERTURA',
      'TERMINO_RETRABALHO': 'INÍCIO DE RETRABALHO',
      'TERMINO_PARADA':     'INÍCIO DE PARADA',
    };
    const tiposFechamento = Object.keys(TIPO_COMPATIVEL);

    if (tiposAbertura.includes(tipo)) {
      for (let i = 1; i < dadosAbertos.length; i++) {
        if (!dadosAbertos[i][0]) continue; // linha vazia
        if (mesmoOperador(dadosAbertos[i][0], payload.operador)) {
          const abertoIdExistente = String(dadosAbertos[i][12] || '');
          const tipoExistente     = String(dadosAbertos[i][2]  || '');
          const serieExistente    = String(dadosAbertos[i][7]  || '');

          // Bug#29 fix (movido de verificarAberto para gravarApontamento):
          // Cross-valida registro em Abertos contra Respostas ANTES de bloquear.
          // Se o abertoId não existe em Respostas (col P), o registro é fantasma —
          // remove-o e deixa a abertura continuar.
          if (abertoIdExistente) {
            const abaRePhantom = ss.getSheetByName(ABA_RESPOSTAS);
            if (abaRePhantom) {
              const dadosRePhantom = abaRePhantom.getDataRange().getValues();
              const existeEmRe = dadosRePhantom.some(
                (r, idx) => idx > 0 && String(r[15] || '').trim() === abertoIdExistente
              );
              if (!existeEmRe) {
                // Fantasma confirmado — remove de Abertos e deixa prosseguir
                abaAb.deleteRow(i + 1);
                Logger.log('gravarApontamento [Bug#29]: fantasma removido — op=' +
                  payload.operador + ', abertoId=' + abertoIdExistente);
                break; // sai do loop e continua para gravar a nova abertura
              }
            }
          }

          // Bug#6 fix: idempotência — se a abertura já existe para o mesmo operador+série+tipo,
          // retorna sucesso com o abertoId existente (recuperação de phantom record)
          const mesmoTipo  = tipoExistente === (TIPOS_APONTAMENTO[tipo] || tipo);
          const mesmaSerie = serieExistente === String(payload.nrSerie || '').trim();
          if (mesmoTipo && mesmaSerie && abertoIdExistente) {
            return jsonResponse({
              success: true,
              abertoId: abertoIdExistente,
              message: 'Apontamento já estava em aberto — abertoId retornado para continuidade.',
              jaAberto: true,
            });
          }
          // Bloqueado por abertura diferente — inclui abertoId para o frontend recuperar
          return jsonResponse({
            success: false,
            bloqueado: true,
            message: 'Operador já possui apontamento em aberto. Feche-o antes de iniciar um novo.',
            abertoId: abertoIdExistente, // Bug#6 fix: frontend usa para fechar o phantom
            aberto: {
              tipo:         tipoExistente,
              operacao:     dadosAbertos[i][3] || '',
              carimbo:      formatarCarimboGs(dadosAbertos[i][4]),
              codItem:      dadosAbertos[i][5] || '',
              qtdPlanejada: dadosAbertos[i][6] || '',
              nrSerie:      serieExistente,
              implemento:   dadosAbertos[i][8] || '',
              cliente:      dadosAbertos[i][9] || '',
            }
          });
        }
      }
    }

    // ---------------------------------------------------------------
    // VALIDAÇÃO DE COMPATIBILIDADE DE TIPO — garante que o fechamento
    // corresponde exatamente ao tipo de abertura em aberto.
    // Ex.: TERMINO_PARADA só pode fechar INÍCIO DE PARADA.
    //
    // Também valida:
    //   • semAberto: bloqueia fechamento quando não há abertura em aberto
    //   • serieIncompativel: bloqueia FECHAMENTO com nrSerie diferente da abertura
    // ---------------------------------------------------------------
    if (tiposFechamento.includes(tipo)) {
      const abertoIdPayload = String(payload.abertoId || '').trim();
      let encontrou = false;
      for (let i = 1; i < dadosAbertos.length; i++) {
        if (!dadosAbertos[i][0]) continue;
        const rowId   = String(dadosAbertos[i][12] || '').trim();
        const matchId = abertoIdPayload && rowId && rowId === abertoIdPayload;
        const matchOp = !matchId && mesmoOperador(dadosAbertos[i][0], payload.operador);
        if (matchId || matchOp) {
          // Bug#5 fix: abertoId só pode ser usado pelo próprio operador que o criou
          if (matchId && !mesmoOperador(dadosAbertos[i][0], payload.operador)) {
            return jsonResponse({
              success: false,
              message: 'Não é possível fechar um apontamento de outro operador.',
              semAberto: true,
            });
          }
          encontrou = true;
          const tipoAberto   = String(dadosAbertos[i][2] || '').trim();
          const tipoEsperado = TIPO_COMPATIVEL[tipo];
          if (tipoEsperado && tipoAberto !== tipoEsperado) {
            return jsonResponse({
              success: false,
              message: 'Tipo de fechamento incompatível. Você possui "' + tipoAberto +
                       '" em aberto, mas tentou registrar "' + (TIPOS_APONTAMENTO[tipo] || tipo) + '".',
              incompativel: true,
              tipoAberto: tipoAberto,
            });
          }
          // Valida série apenas para FECHAMENTO (TERMINO_PARADA/RETRABALHO não têm série)
          if (tipo === 'FECHAMENTO') {
            const serieAberto      = String(dadosAbertos[i][7]  || '').trim();
            const loteAberto       = String(dadosAbertos[i][11] || '').trim();
            const serieFechamento  = String(payload.nrSerie || '').trim();
            const loteFechamento   = payload.loteSeries && Array.isArray(payload.loteSeries) && payload.loteSeries.length > 0;
            // Só valida quando a abertura tem série definida e não é lote
            if (serieAberto && !loteAberto && !loteFechamento) {
              if (serieFechamento && serieFechamento !== serieAberto) {
                return jsonResponse({
                  success: false,
                  message: 'Série incompatível. A abertura foi feita para "' + serieAberto +
                           '" mas o fechamento informou "' + serieFechamento + '".',
                  serieIncompativel: true,
                  serieAberto: serieAberto,
                });
              }
            }
          }
          break; // encontrou o registro — compatível → pode continuar
        }
      }
      // Bloqueia fechamento sem nenhuma abertura em aberto para o operador
      if (!encontrou) {
        return jsonResponse({
          success: false,
          message: 'Nenhum apontamento em aberto encontrado para este operador. Registre uma abertura antes de fechar.',
          semAberto: true,
        });
      }
    }

    // Carimbo: vem do browser já formatado (dd/MM/yyyy HH:mm:ss no fuso local)
    // Fallback para o servidor em GMT-3 caso não venha ou seja inválido
    // Bug#19 fix: validar formato dd/MM/yyyy HH:mm:ss e rejeitar datas absurdas
    let carimbo = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
    if (payload.timestamp && typeof payload.timestamp === 'string' && !payload.timestamp.includes('T')) {
      const tsMatch = payload.timestamp.match(/^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/);
      if (tsMatch) {
        const ano = parseInt(tsMatch[3]);
        const anoAtual = new Date().getFullYear();
        if (ano >= 2020 && ano <= anoAtual + 1) {
          carimbo = payload.timestamp; // aceita apenas anos razoáveis
        }
        // Caso contrário usa server time (timestamp futuro/passado distante ignorado)
      }
    }

    // ID único para o registro na aba Abertos (gerado uma vez por abertura)
    // Fix#ID-LINK: Fechamentos gravam o abertoId da Abertura que estão fechando (não um ID aleatório)
    // Isso cria o elo direto ABERTURA↔FECHAMENTO na planilha para pareamento sem algoritmo
    const abertoId = tiposAbertura.includes(tipo) ? gerarIdApontamento() : null;

    const nomeOperador  = payload.operadorNome || payload.operador || '';
    const tipoFormatado = TIPOS_APONTAMENTO[tipo] || tipo;
    const op1           = OPERACOES[payload.operacao]     || payload.operacao     || '';
    const op2           = OPERACOES[payload.opRetrabalho] || payload.opRetrabalho || '';
    const codItem       = payload.codItem || '';

    const qtd = (payload.quantidade === null || payload.quantidade === undefined || payload.quantidade === '')
      ? '' : Number(payload.quantidade);

    // Bug#26 fix: quantidade não-numérica (ex: "abc" → NaN passa typeof number)
    if (typeof qtd === 'number' && isNaN(qtd)) {
      return jsonResponse({ success: false, message: 'Quantidade inválida: valor não numérico. Recebido: "' + payload.quantidade + '"' });
    }
    // Bug#20 fix: quantidade realizada não pode ser negativa
    if (typeof qtd === 'number' && qtd < 0) {
      return jsonResponse({ success: false, message: 'Quantidade realizada não pode ser negativa. Valor recebido: ' + qtd });
    }
    // Bug#12 fix: limite máximo de quantidade (evita valores absurdos como 999999)
    const QTD_MAX = 99999;
    if (typeof qtd === 'number' && qtd > QTD_MAX) {
      return jsonResponse({ success: false, message: 'Quantidade realizada acima do limite máximo permitido (' + QTD_MAX + '). Valor recebido: ' + qtd });
    }
    const qtdPlNum = Number(payload.qtdPlanejada || 0);
    if (qtdPlNum > QTD_MAX) {
      return jsonResponse({ success: false, message: 'Quantidade planejada acima do limite máximo permitido (' + QTD_MAX + '). Valor recebido: ' + qtdPlNum });
    }

    let campoF = '';
    if (payload.nrSerie && payload.implemento && payload.cliente) {
      campoF = payload.nrSerie + ' | ' + payload.implemento + ' | ' + payload.cliente;
    } else {
      campoF = payload.implemento || '';
    }

    let tipoParada = payload.parada || '';
    if (tipoParada === 'Set-up') tipoParada = 'SET - UP';

    // Coluna O: quantidade planejada (só na abertura)
    const qtdPlanejada = payload.qtdPlanejada ? String(payload.qtdPlanejada) : '';

    const linha = [
      carimbo,                       // A
      nomeOperador,                  // B
      tipoFormatado,                 // C
      op1,                           // D
      codItem,                       // E
      campoF,                        // F
      qtd,                           // G — quantidade realizada
      payload.obs1       || '',      // H
      payload.retrabalho || '',      // I
      payload.numRNC     || '',      // J
      tipoParada,                    // K
      op2,                           // L
      payload.setup      || '',      // M
      payload.obs2       || '',      // N
      qtdPlanejada,                  // O — quantidade planejada
      // Fix#ID-LINK: col P = "elo de pareamento"
      // • ABERTURA  → abertoId (novo ID gerado acima)
      // • FECHAMENTO → payload.abertoId (ID da ABERTURA sendo fechada, validado antes)
      // • Parada/Retrabalho → ID próprio (não participa do pareamento AB↔FE)
      tiposAbertura.includes(tipo)
        ? abertoId
        : (tiposFechamento.includes(tipo) && String(payload.abertoId || '').trim())
          ? String(payload.abertoId).trim()
          : gerarIdApontamento(), // P
    ];

    if (payload.loteSeries && Array.isArray(payload.loteSeries) && payload.loteSeries.length > 0) {
      // Bug#25 fix: todos os itens do lote devem ter nrSerie
      const semNrSerie = payload.loteSeries.filter(item => !item.nrSerie || String(item.nrSerie).trim() === '');
      if (semNrSerie.length > 0) {
        return jsonResponse({ success: false, message: semNrSerie.length + ' item(s) do lote sem nrSerie. Todos os itens precisam informar nrSerie.' });
      }
      // Bug#27 fix: não permite séries duplicadas no mesmo lote
      const nrSeriesNoLote = payload.loteSeries.map(item => String(item.nrSerie).trim());
      const nrSeriesUnicas = new Set(nrSeriesNoLote);
      if (nrSeriesUnicas.size !== nrSeriesNoLote.length) {
        const duplicatas = nrSeriesNoLote.filter((s, i) => nrSeriesNoLote.indexOf(s) !== i);
        return jsonResponse({ success: false, message: 'Séries duplicadas no lote: ' + [...new Set(duplicatas)].join(', '), duplicatas: [...new Set(duplicatas)] });
      }
      // Bug#11 fix: validar séries do LOTE contra cadastro
      const abaSeriesLote = ss.getSheetByName(ABA_SERIES);
      if (abaSeriesLote) {
        const seriesData = abaSeriesLote.getDataRange().getValues().slice(1);
        const seriesValidas = new Set(
          seriesData.filter(r => String(r[3]).toUpperCase() !== 'NÃO' && r[0] !== '')
                    .map(r => String(r[0]).trim())
        );
        const seriesInvalidas = payload.loteSeries
          .filter(item => item.nrSerie && !seriesValidas.has(String(item.nrSerie).trim()))
          .map(item => item.nrSerie);
        if (seriesInvalidas.length > 0) {
          return jsonResponse({
            success: false,
            message: 'Séries não encontradas no cadastro: ' + seriesInvalidas.join(', '),
            seriesInvalidas: seriesInvalidas,
          });
        }
      }

      // Bug#4 fix: quantidade total ÷ número de séries (distribuição proporcional)
      const numSeries = payload.loteSeries.length;
      const qtdPlTotal  = Number(payload.qtdPlanejada || 0);
      const qtdReTotal  = (qtd === '' ? 0 : Number(qtd));
      const qtdPlPorSerie = numSeries > 0 ? Math.ceil(qtdPlTotal / numSeries) : qtdPlTotal;
      const qtdRePorSerie = numSeries > 0 ? Math.ceil(qtdReTotal / numSeries) : qtdReTotal;

      // Batch write — uma única chamada de API para todas as séries do lote
      const rows = payload.loteSeries.map(item => {
        const linhaMod = [...linha];
        linhaMod[5]  = item.nrSerie + ' | ' + item.implemento + ' | ' + item.cliente;
        linhaMod[6]  = qtdRePorSerie;          // G — qtd realizada por série
        linhaMod[14] = String(qtdPlPorSerie);  // O — qtd planejada por série
        linhaMod[15] = abertoId; // P — lote: todas as séries compartilham o mesmo abertoId (Fix#ID-LINK)
        return linhaMod;
      });
      const primeiraLinha = abaRe.getLastRow() + 1;
      abaRe.getRange(primeiraLinha, 1, rows.length, rows[0].length).setValues(rows);
    } else if (payload.lote && payload.lote.trim() !== '') {
      // Formato legado — batch write com divisão proporcional também
      const series = payload.lote.split(',').map(s => s.trim()).filter(Boolean);
      const numSeriesLeg = series.length;
      const qtdPlTotLeg  = Number(payload.qtdPlanejada || 0);
      const qtdReTotLeg  = (qtd === '' ? 0 : Number(qtd));
      const qtdPlLeg = numSeriesLeg > 0 ? Math.ceil(qtdPlTotLeg / numSeriesLeg) : qtdPlTotLeg;
      const qtdReLeg = numSeriesLeg > 0 ? Math.ceil(qtdReTotLeg / numSeriesLeg) : qtdReTotLeg;
      const rows = series.map(serie => {
        const linhaMod = [...linha];
        linhaMod[5]  = serie + ' | ' + payload.implemento + ' | ' + payload.cliente;
        linhaMod[6]  = qtdReLeg;
        linhaMod[14] = String(qtdPlLeg);
        linhaMod[15] = abertoId || gerarIdApontamento(); // P — Bug#30: lote legado usa abertoId
        return linhaMod;
      });
      const primeiraLinha = abaRe.getLastRow() + 1;
      abaRe.getRange(primeiraLinha, 1, rows.length, rows[0].length).setValues(rows);
    } else {
      abaRe.appendRow(linha);
    }

    // ---------------------------------------------------------------
    // FECHAMENTO: lê LoteSeries (col 11) da aba Abertos já lida acima.
    // Match preferencial por AbertoId (col 12); fallback por operador.
    // ---------------------------------------------------------------
    let loteSeriesFechamento = null;
    if (tipo === 'FECHAMENTO') {
      const abertoIdPayload = String(payload.abertoId || '').trim();
      for (let i = 1; i < dadosAbertos.length; i++) {
        const rowId   = String(dadosAbertos[i][12] || '').trim();
        const matchId = abertoIdPayload && rowId && rowId === abertoIdPayload;
        const matchOp = !matchId && mesmoOperador(dadosAbertos[i][0], payload.operador);
        if (matchId || matchOp) {
          const loteJson = String(dadosAbertos[i][11] || '');
          if (loteJson) {
            try { loteSeriesFechamento = JSON.parse(loteJson); } catch(e) {}
          }
          break;
        }
      }
    }

    // Passa dadosAbertos já lidos — atualizarAbertos não precisa reler a aba
    atualizarAbertos(abaAb, dadosAbertos, payload, tipo, tipoFormatado, op1, tipoParada, carimbo, abertoId);
    // Bug#1 fix: força commit imediato para que verificarAberto() leia dados atualizados
    SpreadsheetApp.flush();

    // ---------------------------------------------------------------
    // SALDO PARCIAL — salvo no FECHAMENTO
    // Regras:
    //   • saldo = qtdPlanejada − qtdRealizada  (mínimo 0)
    //   • qtdRealizada = 0 quando o operador fecha sem produzir nenhuma peça
    //   • saldo = 0 apaga o registro (produção completa)
    //   • LOTE: salva saldo individualmente para cada série do lote
    //   • Série única: salva saldo para a série específica (se nrSerie não vazio)
    // ---------------------------------------------------------------
    if (tipo === 'FECHAMENTO') {
      const qtdPl = Number(payload.qtdPlanejada || 0);
      const qtdRe = (payload.quantidade === '' || payload.quantidade === null || payload.quantidade === undefined)
        ? 0 : Number(payload.quantidade);

      // Atualiza saldo sempre que houver série — a função cuida de saldo acumulado vs. primeiro ciclo
      const abaSaldo   = garantirAbaSaldo(ss);
      const opCod      = String(payload.operacao || '').substring(0, 4);
      const codItemKey = String(payload.codItem || '').trim();

      if (loteSeriesFechamento && loteSeriesFechamento.length > 0) {
        // LOTE: atualiza saldo de cada série individualmente
        for (const item of loteSeriesFechamento) {
          atualizarSaldoParcial(
            abaSaldo,
            String(item.nrSerie || '').trim(),
            codItemKey,
            opCod,
            qtdPl,
            qtdRe,
            carimbo
          );
        }
      } else {
        // Série única
        const nrSerieKey = String(payload.nrSerie || '').trim();
        if (nrSerieKey) {
          atualizarSaldoParcial(abaSaldo, nrSerieKey, codItemKey, opCod, qtdPl, qtdRe, carimbo);
        }
      }
    }
    return jsonResponse({ success: true, message: 'Apontamento registrado com sucesso' });

  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// CONTROLE DE ABERTOS
// Colunas: 0:Operador 1:Implemento 2:Tipo 3:Operação 4:Carimbo
//          5:CodItem  6:QtdPlanejada 7:NrSerie 8:Implemento 9:Cliente 10:OperadorNome
// ================================================================

// abertoId: ID único gerado na abertura (AP-...) — usado para identificar a linha exata no fechamento
// dadosAbertos: resultado de getDataRange().getValues() já lido pelo caller — evita releitura
function atualizarAbertos(aba, dadosAbertos, payload, tipo, tipoFormatado, op1, tipoParada, carimbo, abertoId) {
  const operador   = normalizarCodigoOp(payload.operador || ''); // sempre 6 dígitos
  const implemento = payload.nrSerie   || payload.implemento || '';
  const dados      = dadosAbertos;

  const tiposAb = ['ABERTURA', 'INICIO_RETRABALHO', 'INICIO_PARADA'];
  const tiposFe = ['FECHAMENTO', 'TERMINO_RETRABALHO', 'TERMINO_PARADA'];

  if (tiposAb.includes(tipo)) {
    const opLabel = op1 || tipoParada || payload.retrabalho || '';
    const novaLinha = [
      operador,                                             // 0  Operador
      implemento,                                           // 1  Implemento
      tipoFormatado,                                        // 2  Tipo
      opLabel,                                              // 3  Operação
      carimbo,                                              // 4  Carimbo
      payload.codItem      || '',                           // 5  CodItem
      payload.qtdPlanejada || '',                           // 6  QtdPlanejada
      payload.nrSerie      || '',                           // 7  NrSerie
      payload.implemento   || '',                           // 8  ImplementoNome
      payload.cliente      || '',                           // 9  Cliente
      payload.operadorNome || '',                           // 10 OperadorNome
      (payload.loteSeries && Array.isArray(payload.loteSeries) && payload.loteSeries.length > 0)
        ? JSON.stringify(payload.loteSeries) : '',          // 11 LoteSeries (JSON)
      abertoId || '',                                       // 12 AbertoId — chave de rastreamento
    ];

    // Verifica por operador (não por série) — um aberto por operador
    for (let i = 1; i < dados.length; i++) {
      if (mesmoOperador(dados[i][0], operador)) {
        aba.getRange(i + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
        aba.getRange(i + 1, 1).setNumberFormat('@');
        return;
      }
    }
    aba.appendRow(novaLinha);
    aba.getRange(aba.getLastRow(), 1).setNumberFormat('@');

  } else if (tiposFe.includes(tipo)) {
    // Preferência: encontrar linha por AbertoId (imune a zeros à esquerda e colisões)
    // Fallback: match por código do operador (compatibilidade com registros sem ID)
    const abertoIdPayload = String(payload.abertoId || '').trim();
    for (let i = dados.length - 1; i >= 1; i--) {
      const rowId   = String(dados[i][12] || '').trim();
      const matchId = abertoIdPayload && rowId && rowId === abertoIdPayload;
      const matchOp = !matchId && mesmoOperador(dados[i][0], operador);
      if (matchId || matchOp) {
        aba.deleteRow(i + 1);
        return;
      }
    }
  }
}

// ================================================================
// CADASTROS DINÂMICOS
// ================================================================

// ================================================================
// GET ABERTOS — lê a aba "Abertos" e retorna os registros abertos
// Chamado pelo dashboard via ?action=getAbertos
// ================================================================
function getAbertosAction() {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName(ABA_ABERTOS);
    if (!aba) return jsonResponse({ success: true, abertos: [] });
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return jsonResponse({ success: true, abertos: [] });
    const abertos = [];
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (!row[0]) continue; // linha vazia
      let loteSeries = null;
      if (row[11]) {
        try { loteSeries = JSON.parse(String(row[11])); } catch(e) {}
      }
      // col 4 (Carimbo) pode ser Date object no GAS — formatar como string dd/MM/yyyy HH:mm:ss
      const carimboRaw = row[4];
      const carimboStr = carimboRaw instanceof Date
        ? Utilities.formatDate(carimboRaw, 'GMT-3', 'dd/MM/yyyy HH:mm:ss')
        : String(carimboRaw || '');
      abertos.push({
        operador:       String(row[0]  || ''),
        implemento:     String(row[1]  || ''),
        tipo:           String(row[2]  || ''),
        operacao:       String(row[3]  || ''),
        carimbo:        carimboStr,
        codItem:        String(row[5]  || ''),
        qtdPlanejada:   String(row[6]  || ''),
        nrSerie:        String(row[7]  || ''),
        implementoNome: String(row[8]  || ''),
        cliente:        String(row[9]  || ''),
        operadorNome:   String(row[10] || ''),
        loteSeries:     loteSeries,
        abertoId:       String(row[12] || ''),
      });
    }
    return jsonResponse({ success: true, abertos: abertos });
  } catch(err) {
    return jsonResponse({ success: false, erro: err.message });
  }
}

function getCadastros() {
  try {
    // Cache de 5 min — evita abrir a planilha a cada carregamento de página
    const cache  = CacheService.getScriptCache();
    const cached = cache.get('cadastros_v3');
    if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const abaOp  = garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo', 'Nome', 'Ativo']);
    const abaSe  = garantirAbaCadastro(ss, ABA_SERIES,     ['NrSerie', 'Implemento', 'Cliente', 'Ativo']);
    const abaOpc = garantirAbaCadastro(ss, ABA_OPERACOES,  ['Codigo', 'Nome', 'ExigeQtd']);

    const operadores = abaOp.getDataRange().getValues().slice(1)
      .filter(r => String(r[2]).toUpperCase() !== 'NÃO' && r[0] !== '')
      .map(r => ({ codigo: normalizarCodigoOp(String(r[0]).trim()), nome: String(r[1]).trim() }));

    const series = abaSe.getDataRange().getValues().slice(1)
      .filter(r => String(r[3]).toUpperCase() !== 'NÃO' && r[0] !== '')
      .map(r => ({ nrSerie: String(r[0]).trim(), implemento: String(r[1]).trim(), cliente: String(r[2]).trim() }));

    // Operações customizadas criadas pelo líder (complementam as built-in do app)
    const operacoesCustom = abaOpc.getDataRange().getValues().slice(1)
      .filter(r => r[0] !== '')
      .map(r => ({
        cod:       String(r[0]).trim(),
        nome:      String(r[1]).trim(),
        requerQtd: String(r[2]).toUpperCase() === 'SIM',
      }));

    const resultado = JSON.stringify({ success: true, operadores, series, operacoesCustom });
    cache.put('cadastros_v3', resultado, 300); // 5 minutos
    return ContentService.createTextOutput(resultado).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function invalidarCacheCadastros() {
  try {
    const c = CacheService.getScriptCache();
    c.remove('cadastros_v3');
    c.remove('cadastros_v2'); // limpa versão antiga se ainda existir
  } catch(e) {}
}

// ================================================================
// DADOS DA PLANILHA — lidos sempre frescos (sem cache de CacheService)
// Chamado pelo dashboard via ?action=getData
// Retorna array de objetos com os cabeçalhos da aba Respostas como chaves,
// no mesmo formato que o script antigo de leitura retornava.
// ================================================================
function getDadosRespostas() {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName(ABA_RESPOSTAS);
    if (!aba) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);

    const dados    = aba.getDataRange().getValues();
    if (dados.length < 2) return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);

    const cabecalhos = dados[0].map(c => String(c));
    const registros  = [];

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      // Pula linhas completamente vazias
      if (!row[0] && !row[1]) continue;
      const obj = {};
      cabecalhos.forEach((col, j) => {
        const val = row[j];
        // Converte Date objects para string legível
        if (val instanceof Date) {
          obj[col] = Utilities.formatDate(val, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
        } else {
          obj[col] = val === null || val === undefined ? '' : val;
        }
      });
      registros.push(obj);
    }

    return ContentService
      .createTextOutput(JSON.stringify(registros))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
  }
}

function salvarOperador(payload) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo', 'Nome', 'Ativo']);
    // Normaliza: remove zeros à esquerda (ex: "000130" → "130")
    const codigoNorm = normalizarCodigoOp(String(payload.codigo || '').trim());
    aba.getRange('A:A').setNumberFormat('@');
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (mesmoOperador(dados[i][0], codigoNorm)) {
        aba.getRange(i+1, 1).setValue(codigoNorm);
        aba.getRange(i+1, 2).setValue(payload.nome);
        aba.getRange(i+1, 3).setValue('Sim');
        invalidarCacheCadastros();
        return jsonResponse({ success: true, message: 'Operador atualizado.' });
      }
    }
    aba.appendRow([codigoNorm, payload.nome, 'Sim']);
    invalidarCacheCadastros();
    return jsonResponse({ success: true, message: 'Operador adicionado.' });
  } catch (err) { return jsonResponse({ success: false, message: err.message }); }
}

function removerOperador(payload) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo', 'Nome', 'Ativo']);
    const dados = aba.getDataRange().getValues();
    for (let i = dados.length-1; i >= 1; i--) {
      if (String(dados[i][0]).trim() === String(payload.codigo).trim()) {
        aba.getRange(i+1,3).setValue('Não');
        invalidarCacheCadastros();
        return jsonResponse({ success: true, message: 'Operador desativado.' });
      }
    }
    return jsonResponse({ success: false, message: 'Operador não encontrado.' });
  } catch (err) { return jsonResponse({ success: false, message: err.message }); }
}

function salvarSerie(payload) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_SERIES, ['NrSerie', 'Implemento', 'Cliente', 'Ativo']);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(payload.nrSerie).trim()) {
        aba.getRange(i+1,2).setValue(payload.implemento);
        aba.getRange(i+1,3).setValue(payload.cliente);
        aba.getRange(i+1,4).setValue('Sim');
        return jsonResponse({ success: true, message: 'Série atualizada.' });
      }
    }
    aba.appendRow([payload.nrSerie, payload.implemento, payload.cliente, 'Sim']);
    invalidarCacheCadastros();
    return jsonResponse({ success: true, message: 'Série adicionada.' });
  } catch (err) { return jsonResponse({ success: false, message: err.message }); }
}

function removerSerie(payload) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_SERIES, ['NrSerie', 'Implemento', 'Cliente', 'Ativo']);
    const dados = aba.getDataRange().getValues();
    for (let i = dados.length-1; i >= 1; i--) {
      if (String(dados[i][0]).trim() === String(payload.nrSerie).trim()) {
        aba.getRange(i+1,4).setValue('Não');
        invalidarCacheCadastros();
        return jsonResponse({ success: true, message: 'Série desativada.' });
      }
    }
    return jsonResponse({ success: false, message: 'Série não encontrada.' });
  } catch (err) { return jsonResponse({ success: false, message: err.message }); }
}

// ================================================================
// OPERAÇÕES CUSTOMIZADAS
// ================================================================

function salvarOperacao(payload) {
  try {
    const cod  = String(payload.cod  || '').trim().toUpperCase();
    const nome = String(payload.nome || '').trim().toUpperCase();
    if (!cod || !nome) return jsonResponse({ success: false, message: 'Código e nome são obrigatórios.' });

    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_OPERACOES, ['Codigo', 'Nome', 'ExigeQtd']);
    const exigeQtd = payload.requerQtd ? 'Sim' : 'Não';

    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim().toUpperCase() === cod) {
        aba.getRange(i+1, 2).setValue(nome);
        aba.getRange(i+1, 3).setValue(exigeQtd);
        invalidarCacheCadastros();
        return jsonResponse({ success: true, message: 'Operação atualizada.' });
      }
    }
    aba.appendRow([cod, nome, exigeQtd]);
    invalidarCacheCadastros();
    return jsonResponse({ success: true, message: 'Operação adicionada.' });
  } catch (err) { return jsonResponse({ success: false, message: err.message }); }
}

function removerOperacao(payload) {
  try {
    const cod = String(payload.cod || '').trim().toUpperCase();
    if (!cod) return jsonResponse({ success: false, message: 'Código obrigatório.' });

    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_OPERACOES, ['Codigo', 'Nome', 'ExigeQtd']);
    const dados = aba.getDataRange().getValues();
    for (let i = dados.length-1; i >= 1; i--) {
      if (String(dados[i][0]).trim().toUpperCase() === cod) {
        aba.deleteRow(i+1);
        invalidarCacheCadastros();
        return jsonResponse({ success: true, message: 'Operação removida.' });
      }
    }
    return jsonResponse({ success: false, message: 'Operação não encontrada.' });
  } catch (err) { return jsonResponse({ success: false, message: err.message }); }
}

// ================================================================
// HELPERS
// ================================================================

function garantirAbaAbertos(ss) {
  let aba = ss.getSheetByName(ABA_ABERTOS);
  if (!aba) {
    aba = ss.insertSheet(ABA_ABERTOS);
    const cab = ['Operador','Implemento','Tipo','Operação','Carimbo','CodItem','QtdPlanejada','NrSerie','ImplementoNome','Cliente','OperadorNome','LoteSeries','AbertoId'];
    aba.appendRow(cab);
    aba.setFrozenRows(1);
    aba.getRange(1,1,1,cab.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
    // Formatos de texto definidos apenas na criação — evita chamada custosa a cada requisição
    aba.getRange('A:A').setNumberFormat('@');
    aba.getRange('E:E').setNumberFormat('@'); // carimbo como texto
  }
  return aba;
}

function garantirAbaCadastro(ss, nome, cabecalhos) {
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.appendRow(cabecalhos);
    aba.setFrozenRows(1);
    aba.getRange(1,1,1,cabecalhos.length).setFontWeight('bold').setBackground('#0f4c81').setFontColor('#fff');
  }
  return aba;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Gera ID único no formato AP-XXXXXXXXXX
function gerarIdApontamento() {
  return 'AP-' + Utilities.getUuid().replace(/-/g,'').substring(0, 10).toUpperCase();
}

// Normaliza código de operador: Sheets converte '000130' → número 130.
// Comparar como número evita falsos negativos por zeros à esquerda.
// Formata valor lido da planilha como data/hora em pt-BR (dd/MM/yyyy HH:mm:ss)
// Cobre três casos:
//   1. Date object  — Sheets auto-converteu a string para Date
//   2. String inglesa — código antigo fez String(date) → "Thu May 14 2026 12:15:07 GMT-0300..."
//   3. String já em pt-BR — retorna como está
function formatarCarimboGs(val) {
  if (!val) return '';
  // Caso 1: objeto Date
  if (val instanceof Date) return Utilities.formatDate(val, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  const s = String(val).trim();
  if (!s) return '';
  // Caso 3: já no formato dd/MM/yyyy
  if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) return s;
  // Caso 2: string inglesa "Thu May 14 2026 12:15:07 GMT-0300 ..."
  // Extrai componentes via regex — não depende de new Date() que pode falhar com caracteres especiais
  const m = s.match(/\w{3}\s+(\w{3})\s+(\d{1,2})\s+(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (m) {
    const MESES = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
    const mes = MESES[m[1]];
    if (mes) {
      return String(m[2]).padStart(2,'0') + '/' + String(mes).padStart(2,'0') + '/' + m[3] +
             ' ' + m[4] + ':' + m[5] + ':' + m[6];
    }
  }
  return s;
}

function mesmoOperador(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return false;
  const sa = String(a).trim();
  const sb = String(b).trim();
  if (!sa || !sb) return false;
  if (sa === sb) return true;
  const na = Number(sa);
  const nb = Number(sb);
  return !isNaN(na) && na !== 0 && !isNaN(nb) && nb !== 0 && na === nb;
}

// Normaliza código de operador: retorna APENAS o número sem zeros à esquerda.
// Ex: "000130" → "130",  "130" → "130",  130 (número) → "130"
// Estratégia sem zeros — novos registros já chegam nesse formato naturalmente.
function normalizarCodigoOp(val) {
  const s = String(val === null || val === undefined ? '' : val).trim();
  if (!s) return s;
  const n = Number(s);
  // número puro positivo: retorna sem zeros à esquerda
  if (!isNaN(n) && n > 0 && String(n) === s) return String(n);
  // string com zeros à esquerda (ex: "000130"): remove zeros
  const semZeros = s.replace(/^0+(\d)/, '$1');
  const nSem = Number(semZeros);
  if (!isNaN(nSem) && nSem > 0) return String(nSem);
  return s;
}

// ================================================================
// SALDO PARCIAL
// Colunas: 0:NrSerie 1:CodItem 2:Operacao 3:QtdRestante 4:UltimaAtualizacao
// ================================================================

function garantirAbaSaldo(ss) {
  let aba = ss.getSheetByName(ABA_SALDO);
  if (!aba) {
    aba = ss.insertSheet(ABA_SALDO);
    const cab = ['NrSerie','CodItem','Operacao','QtdRestante','UltimaAtualizacao'];
    aba.appendRow(cab);
    aba.setFrozenRows(1);
    aba.getRange(1,1,1,cab.length).setFontWeight('bold').setBackground('#4a2060').setFontColor('#fff');
    // Formatos de texto definidos apenas na criação — evita chamada custosa a cada requisição
    aba.getRange('A:A').setNumberFormat('@'); // NrSerie
    aba.getRange('B:B').setNumberFormat('@'); // CodItem
    aba.getRange('C:C').setNumberFormat('@'); // Operacao ← crítico: '0010' vira 10 sem isso
    aba.getRange('E:E').setNumberFormat('@'); // UltimaAtualizacao
  }
  return aba;
}

// qtdPlanejada: planejado do ciclo atual (usado só se não houver saldo anterior)
// qtdRealizada: produzido neste ciclo (sempre subtraído do saldo existente ou do planejado)
// Regra de acumulação:
//   • Há saldo anterior  → novoSaldo = saldoAnterior - qtdRealizada
//   • Sem saldo anterior → novoSaldo = qtdPlanejada  - qtdRealizada  (primeiro ciclo)
//   • novoSaldo ≤ 0      → remove o registro (produção concluída)
function atualizarSaldoParcial(aba, nrSerie, codItem, operacao, qtdPlanejada, qtdRealizada, carimbo) {
  const dados = aba.getDataRange().getValues();
  const codItemNorm = String(codItem || '').trim();

  // Coleta todos os índices que batem na chave composta
  const matches = [];
  for (let i = 1; i < dados.length; i++) {
    if (
      mesmoOperador(dados[i][0], nrSerie) &&
      String(dados[i][1] || '').trim() === codItemNorm &&
      mesmoOperador(dados[i][2], operacao)
    ) {
      matches.push(i);
    }
  }

  // Calcula o novo saldo
  let novoSaldo;
  if (matches.length > 0) {
    // Há saldo acumulado — subtrai o que foi produzido neste ciclo
    const saldoAtual = Number(dados[matches[0]][3]) || 0;
    novoSaldo = Math.max(0, saldoAtual - qtdRealizada);
  } else {
    // Primeiro ciclo — base é a qtdPlanejada
    if (qtdPlanejada <= 0) return; // sem planejado e sem saldo → nada a fazer
    novoSaldo = Math.max(0, qtdPlanejada - qtdRealizada);
  }

  // Remove duplicatas de baixo para cima
  for (let k = matches.length - 1; k >= 1; k--) {
    aba.deleteRow(matches[k] + 1);
  }

  if (matches.length === 0) {
    if (novoSaldo > 0) {
      aba.appendRow([nrSerie, codItem, operacao, novoSaldo, carimbo]);
    }
    return;
  }

  const linha = matches[0] + 1;
  if (novoSaldo <= 0) {
    aba.deleteRow(linha);
  } else {
    aba.getRange(linha, 3).setValue(operacao);
    aba.getRange(linha, 4).setValue(novoSaldo);
    aba.getRange(linha, 5).setValue(carimbo);
  }
}

function verificarSaldoParcialAction(nrSerie, codItem, operacao) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName(ABA_SALDO);
    if (!aba) return jsonResponse({ temSaldo: false });
    const dados = aba.getDataRange().getValues();
    const codItemNorm = String(codItem || '').trim();
    // Chave composta: NrSerie + CodItem + Operacao
    for (let i = 1; i < dados.length; i++) {
      if (
        mesmoOperador(dados[i][0], nrSerie) &&
        String(dados[i][1] || '').trim() === codItemNorm &&
        mesmoOperador(dados[i][2], operacao)
      ) {
        const qtd = Number(dados[i][3]);
        if (qtd > 0) {
          return jsonResponse({ temSaldo: true, qtdRestante: qtd, ultimaAtualizacao: String(dados[i][4]) });
        }
      }
    }
    return jsonResponse({ temSaldo: false });
  } catch (err) {
    return jsonResponse({ temSaldo: false, erro: err.message });
  }
}

// ================================================================
// UTILITÁRIOS
// ================================================================

function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  garantirAbaAbertos(ss);
  garantirAbaSaldo(ss);
  garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo','Nome','Ativo']);
  garantirAbaCadastro(ss, ABA_SERIES,     ['NrSerie','Implemento','Cliente','Ativo']);
  garantirAbaCadastro(ss, ABA_OPERACOES,  ['Codigo','Nome','ExigeQtd']);

  const abaOp = ss.getSheetByName(ABA_OPERADORES);
  if (abaOp.getLastRow() <= 1) {
    abaOp.getRange(2,1,15,3).setValues([
      ['000101','ELENILTON GONÇALVES DOS SANTOS','Sim'],
      ['000108','SIDNEI OLIVEIRA','Sim'],
      ['000109','DANIEL DO NASCIMENTO PISNO SILVA','Sim'],
      ['000117','RAFAEL MENDES DA SILVA','Sim'],
      ['000119','JONAS GABRIEL SILVA SANTOS','Sim'],
      ['000121','JOÃO PEDRO ALMEIDA FERREIRA','Sim'],
      ['000123','ADILSON BATISTA DA SILVA FILHO','Sim'],
      ['000128','GILMARCIO OLIVEIRA DOS SANTOS','Sim'],
      ['000130','EDERSON LUIS CANDINHO','Sim'],
      ['000131','MATHEUS FERREIRA DA SILVA','Sim'],
      ['001943','CELSO RODRIGUES DE OLIVEIRA','Sim'],
      ['003102','LUIS GUILHERME NASCIMENTO DOS SANTOS','Sim'],
      ['003766','MATHEUS STUCHI','Sim'],
      ['004077','PAULO JOAQUIM DE SANTANA','Sim'],
      ['004223','DEIVI RODRIGO DIAS DA ROSA','Sim'],
    ]);
  }

  const abaSe = ss.getSheetByName(ABA_SERIES);
  if (abaSe.getLastRow() <= 1) {
    abaSe.getRange(2,1,22,4).setValues([
      ['22000072','HAULER 8"','COFCO','Sim'],
      ['22000073','HAULER 10"','SÃO MARTINHO','Sim'],
      ['22000074','HAULER 10"','ATVOS','Sim'],
      ['22000075','HAULER 10"','ATVOS','Sim'],
      ['22000076','IRRIGAÍ 3 LINHAS','DEXCO','Sim'],
      ['22000077','HAULER 10"','ATVOS','Sim'],
      ['22000078','HAULER 10"','ATVOS','Sim'],
      ['22000079','IRRIGAÍ 3 LINHAS','DEMONSTRAÇÃO','Sim'],
      ['22000080','HAULER 10"','ATVOS','Sim'],
      ['22000081','HAULER 10"','VAMOS MG','Sim'],
      ['22000082','HAULER 10"','VAMOS MG','Sim'],
      ['22000083','IRRIGAÍ 3 LINHAS','DEMONSTRAÇÃO','Sim'],
      ['22000084','PLANTADORA','J. LUIZ','Sim'],
      ['22000085','IRRIGAÍ 2 LINHAS','GERDAU','Sim'],
      ['22000086','HAULER 10"','VAMOS RJ','Sim'],
      ['22000087','HAULER 10"','VAMOS RJ','Sim'],
      ['22000088','HAULER 10"','ESTOQUE','Sim'],
      ['INSTITUCIONAL','INSTITUCIONAL AGRICEF','INTERNO','Sim'],
      ['SUCESSO','SUCESSO DO CLIENTE','INTERNO','Sim'],
      ['ENGENHARIA','ENGENHARIA','INTERNO','Sim'],
      ['SENSOR','SENSOR VISION','INTERNO','Sim'],
      ['ESTOQUE','ESTOQUE','INTERNO','Sim'],
    ]);
  }
  Logger.log('Setup completo.');
}

// Corrige registros antigos de Saldo_Parcial onde operacao foi salva como número (ex: 10 em vez de '0010')
// Execute UMA VEZ após atualizar o script
function corrigirSaldoParcial() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = garantirAbaSaldo(ss);
  const dados = aba.getDataRange().getValues();
  let corrigidos = 0;
  for (let i = 1; i < dados.length; i++) {
    const opValor = dados[i][2];
    // Se a operação é um número (ex: 10), converte para string com zero à esquerda (ex: '0010')
    if (typeof opValor === 'number' || (typeof opValor === 'string' && !isNaN(Number(opValor)) && String(Number(opValor)) === String(opValor).trim())) {
      const opCorrigida = String(Number(opValor)).padStart(4, '0');
      aba.getRange(i + 1, 3).setValue(opCorrigida);
      corrigidos++;
    }
  }
  Logger.log(corrigidos + ' registro(s) de operação corrigido(s) na aba Saldo_Parcial.');
}

// Atualiza cabeçalhos da aba Respostas para nomear as colunas O e P corretamente
// Execute uma vez após atualizar o script
function atualizarCabecalhos() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) { Logger.log('Aba Respostas não encontrada.'); return; }
  aba.getRange(1, 15).setValue('QTD PLANEJADA');
  aba.getRange(1, 16).setValue('ID');
  Logger.log('Cabeçalhos atualizados: coluna O = QTD PLANEJADA, coluna P = ID');
}

// IMPORTANTE: rode isso para limpar registros ruins da aba Abertos
function limparAbertos() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_ABERTOS);
  if (!aba) { Logger.log('Aba Abertos não existe.'); return; }
  const ultima = aba.getLastRow();
  if (ultima < 1) { Logger.log('Aba vazia (sem cabeçalho). Nada removido.'); return; }
  const removidos = Math.max(0, ultima - 1);
  if (removidos > 0) aba.deleteRows(2, removidos);
  Logger.log(removidos + ' registro(s) removido(s).');
}

// ================================================================
// Bug#29 — RECONSTRUÇÃO DA ABA ABERTOS A PARTIR DE RESPOSTAS
//
// Reconstrói completamente a aba Abertos percorrendo toda a aba
// Respostas em ordem cronológica. Para cada operador mantém estado
// de "último tipo aberto"; se não houver fechamento correspondente,
// o registro vai para o novo Abertos.
//
// Garante consistência mesmo após edição manual das planilhas,
// falha parcial de escrita ou qualquer outra dessincronização.
//
// Invocado:
//   • Manualmente via ?action=reconstruirAbertos&key=AGF2026
//   • Automaticamente pelo trigger a cada 30 min (se ativado)
// ================================================================

function reconstruirAbertos() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaRe = ss.getSheetByName(ABA_RESPOSTAS);
  const abaAb = garantirAbaAbertos(ss);

  if (!abaRe) {
    Logger.log('reconstruirAbertos: aba Respostas não encontrada.');
    return { corrigidos: 0, abertos: 0, mensagem: 'Aba Respostas não encontrada.' };
  }

  const dadosRe = abaRe.getDataRange().getValues();
  if (dadosRe.length < 2) {
    Logger.log('reconstruirAbertos: aba Respostas vazia.');
    return { corrigidos: 0, abertos: 0, mensagem: 'Aba Respostas vazia.' };
  }

  // Bug#31 fix: mapa reverso de TIPOS_APONTAMENTO para lookup robusto
  // Normaliza strings para NFC evitando divergência de encoding UTF-8 entre
  // o código-fonte GAS e os valores armazenados pelo Sheets
  // (ex: "Í" precomposto U+00CD vs. "I" + combining accent U+0049+U+0301)
  const _tiposAberturaKeys   = ['ABERTURA', 'INICIO_RETRABALHO', 'INICIO_PARADA'];
  const _tiposFechamentoKeys = ['FECHAMENTO', 'TERMINO_RETRABALHO', 'TERMINO_PARADA'];
  // Cria conjuntos NFC-normalizados para comparação insensível a forma de normalização
  const tiposAberturaSet   = new Set(_tiposAberturaKeys.map(k => TIPOS_APONTAMENTO[k].normalize('NFC')));
  const tiposFechamentoSet = new Set(_tiposFechamentoKeys.map(k => TIPOS_APONTAMENTO[k].normalize('NFC')));
  // Helper: classifica um tipo de apontamento
  const ehAbertura   = (t) => tiposAberturaSet.has(String(t || '').normalize('NFC'));
  const ehFechamento = (t) => tiposFechamentoSet.has(String(t || '').normalize('NFC'));

  // Map: codigoOperador → dados da última abertura em aberto
  const abertosPorOp = {};

  for (let i = 1; i < dadosRe.length; i++) {
    const row = dadosRe[i];
    if (!row[0] && !row[1]) continue; // linha vazia

    // Ignora registros ISO antigos (carimbo legado 1899-12-30T...)
    const carimboBruto = String(row[0] || '');
    if (carimboBruto.includes('T') || carimboBruto.includes('1899')) continue;

    const tipo        = String(row[2] || '').trim();
    const operadorRaw = String(row[1] || '').trim();

    // Bug#31 fix: extrai código numérico do operador via regex
    // Funciona com qualquer variante de traço (–, —, -) e formato de nome
    // Ex: "117 — RAFAEL..." → "117" | "000117" → "117" | "117" → "117"
    const codMatch    = operadorRaw.match(/^(\d+)/);
    const codPart     = codMatch ? codMatch[1] : operadorRaw.split(/\s*[—\-–]\s*/)[0].trim();
    const operadorCod = normalizarCodigoOp(codPart || operadorRaw);
    if (!operadorCod) continue;

    if (ehAbertura(tipo)) {
      // Campo F: "22000073 | HAULER 10" | SÃO MARTINHO"
      const campoF     = String(row[5] || '');
      const partes     = campoF.split(' | ');
      const nrSerie    = (partes[0] || '').trim();
      const implemento = (partes[1] || '').trim();
      const cliente    = (partes[2] || '').trim();

      abertosPorOp[operadorCod] = {
        operador:      operadorCod,
        implemento:    nrSerie,          // col 1: chave de busca (=nrSerie)
        tipo:          tipo,             // col 2
        operacao:      String(row[3] || '').trim(), // col 3
        carimbo:       carimboBruto,     // col 4
        codItem:       String(row[4] || '').trim(), // col 5
        qtdPlanejada:  row[14] || '',    // col 6 ← coluna O de Respostas
        nrSerie:       nrSerie,          // col 7
        implementoNome: implemento,      // col 8
        cliente:       cliente,          // col 9
        operadorNome:  operadorRaw,      // col 10
        loteSeries:    '',               // col 11 — não reconstruído (não está em Respostas)
        abertoId:      String(row[15] || '').trim(), // col 12 ← coluna P de Respostas
      };

    } else if (ehFechamento(tipo)) {
      delete abertosPorOp[operadorCod];
    }
  }

  // Guarda quantos registros existiam antes
  const registrosAntes = Math.max(0, abaAb.getLastRow() - 1);

  // Monta linhas novas a serem escritas
  const abertos = Object.values(abertosPorOp);
  const novasLinhas = abertos.map(a => [
    a.operador,
    a.implemento,
    a.tipo,
    a.operacao,
    a.carimbo,
    a.codItem,
    a.qtdPlanejada,
    a.nrSerie,
    a.implementoNome,
    a.cliente,
    a.operadorNome,
    a.loteSeries,
    a.abertoId,
  ]);

  // Limpa Abertos de forma segura:
  // 1. Garante ao menos 1 linha extra para não deixar sheet completamente vazia
  // 2. Escreve as novas linhas (ou linha em branco se não houver abertos)
  // 3. Apaga o excedente
  const numNecessario = Math.max(novasLinhas.length, 1); // ao menos 1 linha abaixo do cabeçalho
  const ultimaLinha   = abaAb.getLastRow();

  // Se faltam linhas, adiciona linhas em branco para ter espaço
  if (ultimaLinha < numNecessario + 1) {
    const faltam = numNecessario + 1 - ultimaLinha;
    for (let k = 0; k < faltam; k++) abaAb.appendRow(['', '', '', '', '', '', '', '', '', '', '', '', '']);
  }

  // Escreve as novas linhas (ou linha vazia na linha 2 se não houver abertos)
  if (novasLinhas.length > 0) {
    abaAb.getRange(2, 1, novasLinhas.length, 13).setValues(novasLinhas);
    // Força formato texto na coluna A (operador)
    abaAb.getRange(2, 1, novasLinhas.length, 1).setNumberFormat('@');
  } else {
    // Sem abertos: apenas limpa linha 2
    abaAb.getRange(2, 1, 1, 13).clearContent();
  }

  // Remove linhas excedentes (abaixo das novas linhas + 1 em branco obrigatória)
  const linhasUsadas = Math.max(novasLinhas.length, 1) + 1; // +1 = cabeçalho
  const totalAtual   = abaAb.getLastRow();
  if (totalAtual > linhasUsadas) {
    abaAb.deleteRows(linhasUsadas + 1, totalAtual - linhasUsadas);
  }

  SpreadsheetApp.flush();

  const msg = 'reconstruirAbertos: ' + registrosAntes + ' antes → ' + abertos.length + ' depois.';
  Logger.log(msg);
  return {
    corrigidos: registrosAntes - abertos.length,
    abertos:    abertos.length,
    mensagem:   msg,
  };
}

// ================================================================
// TRIGGER — Reconciliação periódica da aba Abertos (a cada 30 min)
// Execute criarTriggerReconciliacaoAbertos() uma vez para ativar.
// ================================================================

function criarTriggerReconciliacaoAbertos() {
  // Remove triggers anteriores do mesmo handler (evita duplicação)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'reconstruirAbertos')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('reconstruirAbertos')
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log('Trigger criado: reconstruirAbertos a cada 30 minutos.');
}

// ================================================================
// NORMALIZAÇÃO DE CÓDIGOS DE OPERADOR — Execute sempre que necessário
// Remove zeros à esquerda de todos os registros (ex: "000130" → "130").
// Formato padrão: número simples sem padding — compatível com todos os
// sistemas de entrada que não adicionam zeros automaticamente.
// ================================================================
function normalizarCodigosOperador() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let totalCorrigidos = 0;

  // 1. Cadastro_Operadores — coluna A
  const abaOp = garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo', 'Nome', 'Ativo']);
  abaOp.getRange('A:A').setNumberFormat('@'); // força texto para preservar zeros
  const dadosOp = abaOp.getDataRange().getValues();
  for (let i = 1; i < dadosOp.length; i++) {
    const codAtual  = String(dadosOp[i][0] || '').trim();
    const codNorm   = normalizarCodigoOp(codAtual);
    if (codAtual !== codNorm && codNorm) {
      abaOp.getRange(i + 1, 1).setValue(codNorm);
      totalCorrigidos++;
      Logger.log('Operadores: linha ' + (i+1) + ': "' + codAtual + '" → "' + codNorm + '"');
    }
  }

  // 2. Aba Abertos — coluna A (Operador)
  const abaAb = ss.getSheetByName(ABA_ABERTOS);
  if (abaAb) {
    abaAb.getRange('A:A').setNumberFormat('@');
    const dadosAb = abaAb.getDataRange().getValues();
    for (let i = 1; i < dadosAb.length; i++) {
      const codAtual = String(dadosAb[i][0] || '').trim();
      const codNorm  = normalizarCodigoOp(codAtual);
      if (codAtual !== codNorm && codNorm) {
        abaAb.getRange(i + 1, 1).setValue(codNorm);
        totalCorrigidos++;
        Logger.log('Abertos: linha ' + (i+1) + ': "' + codAtual + '" → "' + codNorm + '"');
      }
    }
  }

  // 3. Aba Respostas — coluna B (NOME DO OPERADOR nos registros históricos)
  totalCorrigidos += normalizarRespostasColB();

  // 4. Invalida cache de cadastros para refletir as correções
  invalidarCacheCadastros();

  Logger.log('✅ Normalização concluída. ' + totalCorrigidos + ' código(s) corrigido(s).');
  return totalCorrigidos;
}

// ================================================================
// NORMALIZAÇÃO DA ABA RESPOSTAS — coluna B (NOME DO OPERADOR)
// Corrige entradas como "130 - NOME" ou "130 — NOME" para
// "000130 - NOME" / "000130 — NOME" nos registros históricos.
// Execute via: ?action=normalizarRespostas&key=AGF2026
// ================================================================
function normalizarRespostasColB() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) { Logger.log('Aba Respostas não encontrada.'); return 0; }
  const lastRow = aba.getLastRow();
  if (lastRow < 2) { Logger.log('Nenhum registro a normalizar.'); return 0; }

  // Leitura em lote — coluna B
  const range = aba.getRange(2, 2, lastRow - 1, 1);
  const colB  = range.getValues();
  let corrigidos = 0;

  for (let i = 0; i < colB.length; i++) {
    const val = String(colB[i][0] === null || colB[i][0] === undefined ? '' : colB[i][0]).trim();
    if (!val) continue;

    // Detecta separadores: " - " (hífen), " — " (travessão), " – " (meia-risca), ou só dígitos
    const mHifen = val.match(/^(\d+)\s*[-–]\s*(.*)/s);   // hífen ou en-dash
    const mTravo = val.match(/^(\d+)\s*—\s*(.*)/s);       // em-dash (correto)
    const mSo    = val.match(/^(\d+)$/);

    let code, nome;
    if      (mHifen) { code = mHifen[1]; nome = mHifen[2].trim(); }
    else if (mTravo) { code = mTravo[1]; nome = mTravo[2].trim(); }
    else if (mSo)    { code = mSo[1];    nome = ''; }
    else continue;

    // Padrão novo do sistema: "CODE — NOME" (travessão em-dash, sem zeros à esquerda)
    const codeNorm = normalizarCodigoOp(code);
    const novoVal  = nome ? (codeNorm + ' — ' + nome) : codeNorm;

    if (novoVal !== val) {
      colB[i][0] = novoVal;
      corrigidos++;
      Logger.log('Col B linha ' + (i + 2) + ': "' + val + '" → "' + novoVal + '"');
    }
  }

  if (corrigidos > 0) {
    range.setNumberFormat('@'); // força texto para preservar zeros se houver
    range.setValues(colB);
  }

  Logger.log('✅ Col B operadores: ' + corrigidos + ' registro(s) normalizado(s).');
  return corrigidos;
}

// ================================================================
// NORMALIZAR COLUNA A (DATAS) — converte tudo para dd/MM/yyyy HH:mm:ss
// Padrão do sistema novo (carimboBrowser).
// Trata: Date objects, YYYY/MM/DD texto, string inglesa "Thu May..."
// Execute via: ?action=normalizarDatas&key=AGF2026
//
// IMPORTANTE: usa a sequência correta para evitar que o Sheets
// reconverta strings de data de volta para Date objects:
//   1. setNumberFormat('@') na coluna INTEIRA
//   2. flush() para aplicar o formato antes de qualquer leitura/escrita
//   3. Converte TODOS os valores (não só os alterados) para string
//   4. setValues() com strings puras
//   5. setNumberFormat('@') novamente após gravar (double-lock)
// ================================================================
function normalizarDatasColA() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) { Logger.log('Aba Respostas não encontrada.'); return 0; }
  const lastRow = aba.getLastRow();
  if (lastRow < 2) { Logger.log('Nenhum registro a normalizar.'); return 0; }

  const range = aba.getRange(2, 1, lastRow - 1, 1);

  // PASSO 1: forçar formato texto na coluna ANTES de ler os valores
  // Sem isso, o Sheets reconverte strings de data para Date objects ao gravar.
  range.setNumberFormat('@');
  SpreadsheetApp.flush(); // aplica imediatamente

  const colA  = range.getValues();
  let corrigidos = 0;

  const MESES_EN = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};

  for (let i = 0; i < colA.length; i++) {
    const val = colA[i][0];
    if (val === null || val === undefined || val === '') continue;

    let nova = null;

    if (val instanceof Date) {
      // Date object (Form response ou Sheets auto-convertido) → texto padrão
      nova = Utilities.formatDate(val, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');

    } else {
      const s = String(val).trim();
      if (!s) continue;

      // Já está em dd/MM/yyyy HH:mm:ss — padrão correto
      // Mesmo assim gravar de volta como string para consolidar o formato texto
      if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(s)) {
        colA[i][0] = s; // garante string pura, não Date
        continue;
      }

      // Já está em dd/MM/yyyy (sem hora) — complementa com 00:00:00
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
        nova = s + ' 00:00:00';
      }

      // yyyy/MM/dd HH:mm:ss  →  dd/MM/yyyy HH:mm:ss
      if (!nova) {
        const mISO = s.match(/^(\d{4})\/(\d{2})\/(\d{2})\s+(\d{2}:\d{2}:\d{2})$/);
        if (mISO) nova = mISO[3] + '/' + mISO[2] + '/' + mISO[1] + ' ' + mISO[4];
      }

      // yyyy/MM/dd HH:mm (sem segundos)
      if (!nova) {
        const mISOt = s.match(/^(\d{4})\/(\d{2})\/(\d{2})\s+(\d{2}:\d{2})$/);
        if (mISOt) nova = mISOt[3] + '/' + mISOt[2] + '/' + mISOt[1] + ' ' + mISOt[4] + ':00';
      }

      // yyyy/MM/dd (sem hora)
      if (!nova) {
        const mISOd = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
        if (mISOd) nova = mISOd[3] + '/' + mISOd[2] + '/' + mISOd[1] + ' 00:00:00';
      }

      // dd/MM/yyyy HH:mm (sem segundos) — complementa
      if (!nova) {
        const mPT = s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}:\d{2})$/);
        if (mPT) nova = mPT[1] + '/' + mPT[2] + '/' + mPT[3] + ' ' + mPT[4] + ':00';
      }

      // String inglesa "Thu May 14 2026 12:15:07 GMT..."
      if (!nova) {
        const mEng = s.match(/\w{3}\s+(\w{3})\s+(\d{1,2})\s+(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
        if (mEng) {
          const mes = MESES_EN[mEng[1]];
          if (mes) {
            nova = String(mEng[2]).padStart(2,'0') + '/' + String(mes).padStart(2,'0') + '/' + mEng[3] +
                   ' ' + mEng[4] + ':' + mEng[5] + ':' + mEng[6];
          }
        }
      }
    }

    if (nova) {
      colA[i][0] = nova;
      corrigidos++;
    }
  }

  // PASSO 2: gravar todos os valores como strings
  range.setValues(colA);

  // PASSO 3: forçar formato texto novamente após gravar (double-lock)
  range.setNumberFormat('@');
  SpreadsheetApp.flush();

  Logger.log('✅ Col A datas: ' + corrigidos + ' registro(s) normalizado(s).');
  return corrigidos;
}

// ================================================================
// NORMALIZAR TUDO — roda datas (col A) + operadores (col B) de uma vez
// Execute via: ?action=normalizarTodos&key=AGF2026
// ================================================================
function normalizarTodosDados() {
  const datas     = normalizarDatasColA();
  const operadores = normalizarRespostasColB();
  return { datas, operadores, total: datas + operadores };
}

function verificarEstrutura() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) { Logger.log('Aba não encontrada!'); return; }
  const cab = aba.getRange(1,1,1,aba.getLastColumn()).getValues()[0];
  cab.forEach((c,i) => Logger.log(String.fromCharCode(65+i) + ': ' + c));
}

function testeGravar() {
  const now = new Date();
  const p2  = n => String(n).padStart(2,'0');
  const ts  = p2(now.getDate())+'/'+p2(now.getMonth()+1)+'/'+now.getFullYear()+' '+p2(now.getHours())+':'+p2(now.getMinutes())+':'+p2(now.getSeconds());
  const payload = {
    timestamp: ts, hora: p2(now.getHours())+':'+p2(now.getMinutes()),
    operador:'004077', operadorNome:'004077 - PAULO JOAQUIM DE SANTANA',
    nrSerie:'22000074', implemento:'HAULER 10"', cliente:'ATVOS',
    tipoApontamento:'ABERTURA', operacao:'0010',
    codItem:'509419', quantidade:'', qtdPlanejada:'20', obs1:'Teste v4',
    setup:'', obs2:'', retrabalho:'', numRNC:'', opRetrabalho:'', parada:'', lote:'',
  };
  Logger.log(gravarApontamento(payload).getContent());
}

// ================================================================
// RELATÓRIO SEMANAL AUTOMATIZADO — v1.0
//
// Envio: toda segunda-feira às 07h00 (horário de Brasília)
// Para ativar o envio automático, abra o Editor do Apps Script
// e execute UMA VEZ a função: criarTriggerRelatorioSemanal()
//
// Para testar manualmente: execute enviarRelatorioSemanal()
// ================================================================

const EMAIL_RELATORIO    = 'guilherme.souza@agricef.com.br';
const DIAS_ALERTA_ATRASO = 3; // ordens abertas há mais que isso → alerta vermelho

// ---------------------------------------------------------------
// PONTO DE ENTRADA — chamado pelo trigger ou manualmente
// ---------------------------------------------------------------
function enviarRelatorioSemanal() {
  try {
    const rel = _rsMontarRelatorio();
    _rsEnviarEmail(rel);
    Logger.log('✅ Relatório semanal enviado para ' + EMAIL_RELATORIO);
  } catch (err) {
    Logger.log('❌ Erro no relatório: ' + err.message + '\n' + err.stack);
    throw err;
  }
}

// ---------------------------------------------------------------
// CONFIGURAR TRIGGER SEMANAL — executar UMA VEZ pelo Editor
// ---------------------------------------------------------------
function criarTriggerRelatorioSemanal() {
  // Remove triggers duplicados da mesma função
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarRelatorioSemanal')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('enviarRelatorioSemanal')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();

  Logger.log('✅ Trigger criado: toda segunda-feira às 07h (GMT-3).');
}

// ---------------------------------------------------------------
// MONTAR OBJETO DO RELATÓRIO — lê planilha e computa todos os KPIs
// ---------------------------------------------------------------
function _rsMontarRelatorio() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) throw new Error('Aba "' + ABA_RESPOSTAS + '" não encontrada.');

  const dados = aba.getDataRange().getValues();
  const agora = new Date();

  // Início da semana atual (segunda-feira 00h00)
  const dow = agora.getDay(); // 0=dom, 1=seg ...
  const diasAteSeg = (dow === 0) ? 6 : dow - 1;
  const seg = new Date(agora);
  seg.setDate(agora.getDate() - diasAteSeg);
  seg.setHours(0, 0, 0, 0);

  const segAnterior = new Date(seg);
  segAnterior.setDate(seg.getDate() - 7);

  // ----- Parse de todas as linhas relevantes -----
  const linhas = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (!r[0] || !r[2]) continue;
    const ts = _rsParseData(r[0]);
    if (!ts) continue;
    const tipoRaw = String(r[2]).toUpperCase();
    let tipo = null;
    if (tipoRaw.includes('ABERTURA') && !tipoRaw.includes('RETRABALHO')) tipo = 'ABERTURA';
    else if (tipoRaw.includes('FECHAMENTO')) tipo = 'FECHAMENTO';
    else continue; // ignora paradas e retrabalhos neste contexto
    const campoF = String(r[5] || '');
    const pts    = campoF.split('|').map(x => x.trim());
    const funcRaw = String(r[1] || '').trim();
    linhas.push({
      ts,
      tipo,
      func:     funcRaw,
      funcCod:  _rdCodigoOp(funcRaw), // código normalizado para pareamento
      op:     String(r[3] || '').trim(),
      item:   String(r[4] || '').trim(),
      serie:  pts[0] || '',
      impl:   pts[1] || '',
      client: pts[2] || '',
      qty:    Number(r[6])  || 0,
      qtdPl:  Number(r[14]) || 0,
    });
  }

  // ----- Algoritmo de pareamento ABERTURA ↔ FECHAMENTO -----
  const aberturas = linhas
    .filter(r => r.tipo === 'ABERTURA')
    .map(r => ({ ...r, _used: false, _fechTs: null, _leadMs: 0 }));

  linhas
    .filter(r => r.tipo === 'FECHAMENTO')
    .forEach(fech => {
      let melhor = null, melhorScore = -1;
      for (let i = 0; i < aberturas.length; i++) {
        const a = aberturas[i];
        // funcCod tolera hífen vs em-dash e variações de zeros
        if (a._used || a.funcCod !== fech.funcCod || a.ts > fech.ts) continue;
        let score = 1;
        if (a.op    === fech.op)    score += 4;
        if (a.serie === fech.serie) score += 2;
        if (a.item  === fech.item)  score += 1;
        // Empate: prefere a ABERTURA mais recente (mais próxima do FECHAMENTO)
        if (score > melhorScore || (score === melhorScore && melhor !== null && a.ts > aberturas[melhor].ts)) {
          melhorScore = score; melhor = i;
        }
      }
      if (melhor !== null) {
        aberturas[melhor]._used   = true;
        aberturas[melhor]._fechTs = fech.ts;
        aberturas[melhor]._leadMs = fech.ts - aberturas[melhor].ts;
        aberturas[melhor]._fechOp = fech.op;
      }
    });

  const pares   = aberturas.filter(a => a._used);
  const backlog = aberturas.filter(a => !a._used).sort((x, y) => x.ts - y.ts);

  // ----- KPIs semana atual e anterior -----
  const kpiAtual    = _rsKpiSemana(pares, seg,         new Date(seg.getTime()          + 7 * 86400000));
  const kpiAnterior = _rsKpiSemana(pares, segAnterior, new Date(segAnterior.getTime()  + 7 * 86400000));

  // ----- Tendência 4 semanas -----
  const trend = [];
  for (let w = 3; w >= 0; w--) {
    const ini = new Date(seg); ini.setDate(seg.getDate() - w * 7);
    const fim = new Date(ini); fim.setDate(ini.getDate() + 7);
    trend.push({ ini, fim, ..._rsKpiSemana(pares, ini, fim) });
  }

  // ----- Top operadores (semana atual) — agrupa por código normalizado -----
  const contOp = {};
  pares.filter(p => p._fechTs >= seg).forEach(p => {
    const cod  = p.funcCod || p.func;
    const nome = _rdNomeOp(p.func);
    if (!contOp[cod]) contOp[cod] = { nome, count: 0 };
    contOp[cod].count++;
  });
  const topOperadores = Object.values(contOp)
    .sort((a, b) => b.count - a.count)
    .slice(0, 5)
    .map(e => [e.nome, e.count]);

  // ----- Top operações (semana atual) -----
  const contOperacao = {};
  pares.filter(p => p._fechTs >= seg).forEach(p => {
    const op = (p._fechOp || p.op || '(sem operação)');
    contOperacao[op] = (contOperacao[op] || 0) + 1;
  });
  const topOperacoes = Object.entries(contOperacao).sort((a, b) => b[1] - a[1]).slice(0, 5);

  // ----- Alertas e recomendações -----
  const alertas = _rsAlertas(backlog, kpiAtual, kpiAnterior, agora);
  const recos   = _rsRecomendacoes(backlog, kpiAtual, kpiAnterior, topOperadores, topOperacoes, agora);

  return {
    agora, seg, segAnterior,
    kpiAtual, kpiAnterior,
    trend, topOperadores, topOperacoes,
    backlog, alertas, recos,
  };
}

// ---------------------------------------------------------------
// KPIs DE UM PERÍODO
// ---------------------------------------------------------------
function _rsKpiSemana(pares, inicio, fim) {
  const entregas = pares.filter(p => p._fechTs >= inicio && p._fechTs < fim);
  const entradas = pares.filter(p => p.ts      >= inicio && p.ts      < fim);

  const leadTimes = entregas
    .map(p => p._leadMs / 3600000)
    .filter(h => h > 0 && h < 24 * 30); // sanidade: entre 0 e 30 dias

  const leadMedio = leadTimes.length
    ? leadTimes.reduce((s, v) => s + v, 0) / leadTimes.length
    : 0;

  return {
    throughput: entregas.length,
    entradas:   entradas.length,
    leadMedio,  // em horas
  };
}

// ---------------------------------------------------------------
// ALERTAS INTELIGENTES
// ---------------------------------------------------------------
function _rsAlertas(backlog, kpiAtual, kpiAnterior, agora) {
  const lista = [];

  // 1. Volume de backlog
  if (backlog.length >= 15) {
    lista.push({ nivel: 'CRITICO', msg: 'Backlog crítico: ' + backlog.length + ' ordens em aberto. Verifique gargalos e redistribua a carga imediatamente.' });
  } else if (backlog.length >= 8) {
    lista.push({ nivel: 'ATENCAO', msg: 'Backlog elevado: ' + backlog.length + ' ordens em aberto. Monitorar evolução ao longo da semana.' });
  }

  // 2. Ordens antigas sem fechamento
  const limiteMs = DIAS_ALERTA_ATRASO * 86400 * 1000;
  const antigas  = backlog.filter(o => (agora - o.ts) > limiteMs);
  if (antigas.length > 0) {
    const series = [...new Set(antigas.map(o => o.serie).filter(Boolean))].slice(0, 3).join(', ');
    lista.push({
      nivel: 'ATENCAO',
      msg: antigas.length + ' ordem(ns) aberta(s) há mais de ' + DIAS_ALERTA_ATRASO + ' dias sem fechamento.'
           + (series ? ' Séries: ' + series + '.' : ''),
    });
  }

  // 3. Queda de throughput ≥ 25%
  if (kpiAnterior.throughput > 0) {
    const pct = Math.round(((kpiAtual.throughput - kpiAnterior.throughput) / kpiAnterior.throughput) * 100);
    if (pct <= -25) {
      lista.push({ nivel: 'ATENCAO', msg: 'Queda de ' + Math.abs(pct) + '% no throughput vs. semana anterior (' + kpiAnterior.throughput + ' → ' + kpiAtual.throughput + ' entregas).' });
    }
  }

  // 4. Semana sem nenhum fechamento
  if (kpiAtual.throughput === 0) {
    lista.push({ nivel: 'CRITICO', msg: 'Nenhum fechamento registrado nesta semana. Verifique se os operadores estão utilizando o sistema.' });
  }

  // 5. Lead time alto (> 8h)
  if (kpiAtual.leadMedio > 8) {
    lista.push({ nivel: 'ATENCAO', msg: 'Lead time médio elevado: ' + kpiAtual.leadMedio.toFixed(1) + 'h. Pode indicar espera de material, retrabalho ou operações subdimensionadas.' });
  }

  // 6. Entradas muito acima de saídas
  if (kpiAtual.entradas > 0 && kpiAtual.throughput > 0 && kpiAtual.entradas > kpiAtual.throughput * 1.5) {
    lista.push({ nivel: 'ATENCAO', msg: 'Entradas (' + kpiAtual.entradas + ') superam entregas (' + kpiAtual.throughput + ') em ' + Math.round(((kpiAtual.entradas / kpiAtual.throughput) - 1) * 100) + '%. O backlog tende a crescer.' });
  }

  return lista;
}

// ---------------------------------------------------------------
// RECOMENDAÇÕES CONTEXTUAIS
// ---------------------------------------------------------------
function _rsRecomendacoes(backlog, kpiAtual, kpiAnterior, topOp, topOperacoes, agora) {
  const lista = [];
  const limiteMs = DIAS_ALERTA_ATRASO * 86400 * 1000;
  const antigas  = backlog.filter(o => (agora - o.ts) > limiteMs);

  if (antigas.length > 0) {
    lista.push('⚡ Priorizar o fechamento das ' + antigas.length + ' ordem(ns) com mais de ' + DIAS_ALERTA_ATRASO + ' dias em aberto. Conversar diretamente com os operadores envolvidos para identificar impedimentos.');
  }

  if (kpiAnterior.throughput > 0 && kpiAtual.throughput < kpiAnterior.throughput) {
    lista.push('📉 Throughput abaixo do período anterior. Revisar distribuição de tarefas, checar se houve paradas não registradas e avaliar se há recursos disponíveis para reforço.');
  }

  if (backlog.length > 5) {
    const opBacklog = {};
    backlog.forEach(o => { const k = o.op || ''; if (k) opBacklog[k] = (opBacklog[k] || 0) + 1; });
    const gargalos = Object.entries(opBacklog).sort((a, b) => b[1] - a[1]).slice(0, 2).map(e => e[0]).join(', ');
    if (gargalos) {
      lista.push('🔍 Operações com maior concentração no backlog: ' + gargalos + '. Avaliar reforço de equipe ou resequenciamento nessas etapas.');
    }
  }

  if (kpiAtual.leadMedio > 6) {
    lista.push('⏱ Lead time médio de ' + kpiAtual.leadMedio.toFixed(1) + 'h. Investigar se operações podem ser subdivididas, paralelizadas ou se há tempos de espera desnecessários no processo.');
  }

  if (kpiAtual.entradas > kpiAtual.throughput * 1.3 && kpiAtual.entradas > 0) {
    lista.push('📥 Volume de entradas supera entregas em mais de 30%. Considerar priorização de carteira ou alocação de capacidade extra para evitar crescimento contínuo do backlog.');
  }

  if (topOp.length > 0) {
    lista.push('🏆 Destaque da semana: ' + topOp[0][0] + ' com ' + topOp[0][1] + ' fechamentos. Identificar boas práticas deste operador e replicar para a equipe.');
  }

  if (lista.length === 0) {
    lista.push('✅ Indicadores dentro do esperado. Manter cadência de apontamentos e foco na antecipação de demandas para as próximas semanas.');
  }

  return lista;
}

// ---------------------------------------------------------------
// GERAÇÃO DO HTML DO EMAIL
// ---------------------------------------------------------------
function _rsGerarHtml(rel) {
  const { agora, seg, kpiAtual, kpiAnterior, trend, topOperadores, topOperacoes, backlog, alertas, recos } = rel;

  const fmt  = d => Utilities.formatDate(d, 'GMT-3', 'dd/MM/yyyy');
  const fmtH = h => h < 1 ? Math.round(h * 60) + 'min' : h.toFixed(1) + 'h';

  const fimSem     = new Date(seg.getTime() + 6 * 86400000);
  const semanaStr  = fmt(seg) + ' – ' + fmt(fimSem);
  const dataEnvio  = Utilities.formatDate(agora, 'GMT-3', "dd/MM/yyyy 'às' HH:mm");

  // Delta throughput
  const deltaTP    = kpiAnterior.throughput > 0
    ? Math.round(((kpiAtual.throughput - kpiAnterior.throughput) / kpiAnterior.throughput) * 100)
    : null;
  const deltaTPStr   = deltaTP === null ? '—' : (deltaTP >= 0 ? '▲ ' + deltaTP + '%' : '▼ ' + Math.abs(deltaTP) + '%');
  const deltaTPColor = deltaTP === null ? '#8b949e' : (deltaTP >= 0 ? '#4ade80' : '#f87171');

  // Cor do KPI backlog
  const corBacklog = backlog.length >= 15 ? '#f87171' : backlog.length >= 8 ? '#fb923c' : '#4ade80';

  // ----- HTML: Alertas -----
  let htmlAlertas = '';
  if (alertas.length === 0) {
    htmlAlertas = '<tr><td style="padding:12px 16px;color:#4ade80;font-size:14px;">✅ Nenhum alerta crítico esta semana.</td></tr>';
  } else {
    alertas.forEach(function(a) {
      var cor = a.nivel === 'CRITICO' ? '#f87171' : '#fb923c';
      var ico = a.nivel === 'CRITICO' ? '🔴' : '🟡';
      htmlAlertas += '<tr><td style="padding:9px 14px 9px 16px;border-left:3px solid ' + cor + ';background:#1a2030;border-radius:0 4px 4px 0;font-size:13px;color:#e6edf3;margin-bottom:4px;">'
        + ico + ' ' + a.msg + '</td></tr>';
    });
  }

  // ----- HTML: Recomendações -----
  var htmlRecos = '';
  recos.forEach(function(r) {
    htmlRecos += '<tr><td style="padding:9px 16px;font-size:13px;color:#e6edf3;border-bottom:1px solid #252d40;line-height:1.5;">' + r + '</td></tr>';
  });

  // ----- HTML: Top Operadores -----
  var htmlTopOp = '';
  var medalhas = ['🥇','🥈','🥉','4°','5°'];
  topOperadores.forEach(function(entry, i) {
    htmlTopOp += '<tr>'
      + '<td style="padding:7px 12px;font-size:20px;width:32px;">' + (medalhas[i] || (i+1)+'°') + '</td>'
      + '<td style="padding:7px 8px;font-size:13px;color:#e6edf3;">' + entry[0] + '</td>'
      + '<td style="padding:7px 12px;font-size:14px;font-weight:700;color:#f0b429;text-align:right;">' + entry[1] + '</td>'
      + '</tr>';
  });
  if (!htmlTopOp) htmlTopOp = '<tr><td colspan="3" style="padding:12px;color:#8b949e;font-size:13px;text-align:center;">Sem fechamentos nesta semana.</td></tr>';

  // ----- HTML: Top Operações -----
  var htmlTopOper = '';
  topOperacoes.forEach(function(entry) {
    htmlTopOper += '<tr>'
      + '<td style="padding:7px 14px;font-size:13px;color:#e6edf3;">' + entry[0] + '</td>'
      + '<td style="padding:7px 14px;font-size:13px;color:#38bdf8;font-weight:700;text-align:right;">' + entry[1] + '</td>'
      + '</tr>';
  });
  if (!htmlTopOper) htmlTopOper = '<tr><td colspan="2" style="padding:12px;color:#8b949e;font-size:13px;text-align:center;">Sem dados.</td></tr>';

  // ----- HTML: Backlog -----
  var htmlBacklog = '';
  var backlogTop10 = backlog.slice(0, 10);
  if (backlogTop10.length === 0) {
    htmlBacklog = '<tr><td colspan="4" style="padding:14px;color:#4ade80;font-size:13px;text-align:center;">✅ Nenhuma ordem em aberto.</td></tr>';
  } else {
    backlogTop10.forEach(function(o) {
      var diasAberto = Math.floor((agora - o.ts) / 86400000);
      var corDias    = diasAberto >= DIAS_ALERTA_ATRASO ? '#f87171' : '#8b949e';
      var peso       = diasAberto >= DIAS_ALERTA_ATRASO ? '700' : '400';
      var nome       = _rdNomeOp(o.func);
      htmlBacklog += '<tr>'
        + '<td style="padding:6px 10px;font-size:12px;color:#8b949e;border-bottom:1px solid #252d40;">' + fmt(new Date(o.ts)) + '</td>'
        + '<td style="padding:6px 10px;font-size:12px;color:#e6edf3;border-bottom:1px solid #252d40;">' + nome + '</td>'
        + '<td style="padding:6px 10px;font-size:12px;color:#8b949e;border-bottom:1px solid #252d40;">' + (o.op || '—') + '</td>'
        + '<td style="padding:6px 10px;font-size:12px;color:' + corDias + ';font-weight:' + peso + ';border-bottom:1px solid #252d40;text-align:right;">' + diasAberto + 'd</td>'
        + '</tr>';
    });
    if (backlog.length > 10) {
      htmlBacklog += '<tr><td colspan="4" style="padding:8px 10px;font-size:12px;color:#8b949e;text-align:center;">… e mais ' + (backlog.length - 10) + ' ordem(ns) em aberto</td></tr>';
    }
  }

  // ----- HTML: Trend 4 semanas (barras inline) -----
  var maxTrend = Math.max.apply(null, trend.map(function(t){ return Math.max(t.throughput, t.entradas, 1); }));
  var htmlTrend = '';
  trend.forEach(function(t, i) {
    var isAtual  = (i === 3);
    var hTp      = Math.max(4, Math.round((t.throughput / maxTrend) * 80));
    var hEn      = Math.max(4, Math.round((t.entradas   / maxTrend) * 80));
    var corTp    = isAtual ? '#f0b429' : '#4ade80';
    var semLabel = fmt(t.ini).substring(0, 5);
    htmlTrend += '<td style="text-align:center;padding:0 10px;vertical-align:bottom;">'
      + '<div style="display:inline-block;vertical-align:bottom;">'
      + '<div style="width:14px;height:' + hEn + 'px;background:#38bdf8;opacity:0.7;border-radius:2px 2px 0 0;display:inline-block;vertical-align:bottom;margin-right:2px;" title="Entradas: ' + t.entradas + '"></div>'
      + '<div style="width:14px;height:' + hTp + 'px;background:' + corTp + ';border-radius:2px 2px 0 0;display:inline-block;vertical-align:bottom;" title="Saídas: ' + t.throughput + '"></div>'
      + '</div>'
      + '<div style="font-size:10px;color:' + (isAtual ? '#f0b429' : '#8b949e') + ';margin-top:5px;font-weight:' + (isAtual ? '700' : '400') + ';">' + semLabel + '</div>'
      + '<div style="font-size:12px;color:' + corTp + ';font-weight:700;">' + t.throughput + '</div>'
      + '</td>';
  });

  // ===== MONTAGEM FINAL DO HTML =====
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="margin:0;padding:0;background:#0a0e17;font-family:Arial,Helvetica,sans-serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#0a0e17;padding:24px 12px;">'
    + '<tr><td align="center">'
    + '<table width="620" cellpadding="0" cellspacing="0" style="max-width:620px;width:100%;">'

    // ── CABEÇALHO ──
    + '<tr><td style="background:#0d1117;border-bottom:3px solid #f0b429;padding:24px 28px;border-radius:12px 12px 0 0;">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
    + '<td><div style="font-size:26px;font-weight:900;color:#f0b429;letter-spacing:3px;line-height:1;">AGRICEF</div>'
    + '<div style="font-size:11px;color:#58677a;letter-spacing:2px;margin-top:3px;text-transform:uppercase;">Sistema de Apontamento de Produção</div></td>'
    + '<td align="right">'
    + '<div style="font-size:16px;font-weight:700;color:#e6edf3;">📊 Relatório Semanal</div>'
    + '<div style="font-size:12px;color:#8b949e;margin-top:4px;">Semana: ' + semanaStr + '</div>'
    + '<div style="font-size:11px;color:#58677a;margin-top:2px;">Gerado em: ' + dataEnvio + '</div>'
    + '</td></tr></table></td></tr>'

    // ── KPI CARDS ──
    + '<tr><td style="background:#111722;padding:20px 28px 12px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'

    + '<td style="width:25%;padding:4px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;border-top:3px solid #4ade80;">'
    + '<tr><td style="padding:14px 12px;">'
    + '<div style="font-size:10px;color:#8b949e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">ENTREGAS</div>'
    + '<div style="font-size:30px;font-weight:900;color:#4ade80;line-height:1;">' + kpiAtual.throughput + '</div>'
    + '<div style="font-size:11px;color:' + deltaTPColor + ';margin-top:4px;">' + deltaTPStr + ' vs sem. ant.</div>'
    + '</td></tr></table></td>'

    + '<td style="width:25%;padding:4px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;border-top:3px solid #38bdf8;">'
    + '<tr><td style="padding:14px 12px;">'
    + '<div style="font-size:10px;color:#8b949e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">ENTRADAS</div>'
    + '<div style="font-size:30px;font-weight:900;color:#38bdf8;line-height:1;">' + kpiAtual.entradas + '</div>'
    + '<div style="font-size:11px;color:#8b949e;margin-top:4px;">novas ordens</div>'
    + '</td></tr></table></td>'

    + '<td style="width:25%;padding:4px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;border-top:3px solid ' + corBacklog + ';">'
    + '<tr><td style="padding:14px 12px;">'
    + '<div style="font-size:10px;color:#8b949e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">BACKLOG</div>'
    + '<div style="font-size:30px;font-weight:900;color:' + corBacklog + ';line-height:1;">' + backlog.length + '</div>'
    + '<div style="font-size:11px;color:#8b949e;margin-top:4px;">em aberto</div>'
    + '</td></tr></table></td>'

    + '<td style="width:25%;padding:4px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;border-top:3px solid #a78bfa;">'
    + '<tr><td style="padding:14px 12px;">'
    + '<div style="font-size:10px;color:#8b949e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">LEAD TIME</div>'
    + '<div style="font-size:30px;font-weight:900;color:#a78bfa;line-height:1;">' + (kpiAtual.leadMedio > 0 ? fmtH(kpiAtual.leadMedio) : '—') + '</div>'
    + '<div style="font-size:11px;color:#8b949e;margin-top:4px;">tempo médio/op</div>'
    + '</td></tr></table></td>'

    + '</tr></table></td></tr>'

    // ── ALERTAS ──
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">🚨 Alertas da Semana</div>'
    + '<table width="100%" cellpadding="0" cellspacing="3">' + htmlAlertas + '</table>'
    + '</td></tr>'

    // ── TENDÊNCIA 4 SEMANAS ──
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">📈 Tendência de Produção (4 semanas)</div>'
    + '<table cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;padding:12px 8px;width:100%;">'
    + '<tr><td style="padding:4px 14px 10px;" colspan="4">'
    + '<span style="display:inline-block;width:10px;height:10px;background:#38bdf8;border-radius:2px;margin-right:4px;opacity:0.7;"></span>'
    + '<span style="font-size:11px;color:#8b949e;">Entradas &nbsp;</span>'
    + '<span style="display:inline-block;width:10px;height:10px;background:#4ade80;border-radius:2px;margin-right:4px;"></span>'
    + '<span style="font-size:11px;color:#8b949e;">Saídas</span>'
    + '</td></tr>'
    + '<tr>' + htmlTrend + '</tr>'
    + '</table></td></tr>'

    // ── TOP OPERADORES + TOP OPERAÇÕES ──
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'

    + '<td style="width:50%;vertical-align:top;padding-right:8px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">👷 Top Operadores</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">' + htmlTopOp + '</table>'
    + '</td>'

    + '<td style="width:50%;vertical-align:top;padding-left:8px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">⚙️ Operações com + Saídas</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">' + htmlTopOper + '</table>'
    + '</td>'

    + '</tr></table></td></tr>'

    // ── BACKLOG DETALHADO ──
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">📋 Backlog Atual — ' + backlog.length + ' ordem' + (backlog.length !== 1 ? 's' : '') + ' em aberto</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">'
    + '<tr style="background:#141a28;">'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;font-weight:600;border-bottom:1px solid #252d40;">ABERTURA</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;font-weight:600;border-bottom:1px solid #252d40;">OPERADOR</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;font-weight:600;border-bottom:1px solid #252d40;">OPERAÇÃO</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:right;font-weight:600;border-bottom:1px solid #252d40;">DIAS</th>'
    + '</tr>'
    + htmlBacklog
    + '</table></td></tr>'

    // ── RECOMENDAÇÕES ──
    + '<tr><td style="background:#111722;padding:4px 28px 24px;border-radius:0 0 0 0;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">💡 Recomendações</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">' + htmlRecos + '</table>'
    + '</td></tr>'

    // ── FOOTER ──
    + '<tr><td style="background:#0d1117;padding:16px 28px;border-radius:0 0 12px 12px;border-top:1px solid #1c2230;text-align:center;">'
    + '<div style="font-size:11px;color:#58677a;">Este relatório é gerado automaticamente toda segunda-feira às 07h00 (Brasília).</div>'
    + '<div style="font-size:11px;color:#3a4460;margin-top:4px;">AGRICEF — Dashboard de Produção · Sistema de Apontamento v4</div>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}

// ---------------------------------------------------------------
// ENVIO DO EMAIL
// ---------------------------------------------------------------
function _rsEnviarEmail(rel) {
  var dataStr = Utilities.formatDate(rel.agora, 'GMT-3', 'dd/MM/yyyy');
  var html    = _rsGerarHtml(rel);
  MailApp.sendEmail({
    to:       EMAIL_RELATORIO,
    subject:  '📊 AGRICEF | Relatório Semanal de Produção — ' + dataStr,
    htmlBody: html,
  });
}

// ---------------------------------------------------------------
// PARSER DE DATA — converte string dd/MM/yyyy HH:mm:ss → Date
// ---------------------------------------------------------------
function _rsParseData(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  // Formato principal: dd/MM/yyyy HH:mm:ss
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], +m[6]);
  // Fallback genérico
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ================================================================
//  RELATÓRIO DIÁRIO — AGRICEF
//
//  Enviado de segunda a sexta-feira às 06h45 (GMT-3).
//  Foco operacional: onde agir hoje, quem conversar, gargalos.
//
//  Para ativar, execute UMA VEZ: criarTriggerRelatorioDiario()
//  Para testar manualmente: execute enviarRelatorioDiario()
// ================================================================

// ---------------------------------------------------------------
// PONTO DE ENTRADA
// ---------------------------------------------------------------
function enviarRelatorioDiario() {
  try {
    const rel = _rdMontarRelatorio();
    _rdEnviarEmail(rel);
    Logger.log('✅ Relatório diário enviado para ' + EMAIL_RELATORIO);
  } catch (err) {
    Logger.log('❌ Erro no relatório diário: ' + err.message + '\n' + err.stack);
    throw err;
  }
}

// ---------------------------------------------------------------
// CONFIGURAR TRIGGER DIÁRIO — executar UMA VEZ pelo Editor
// ---------------------------------------------------------------
function criarTriggerRelatorioDiario() {
  // Remove triggers duplicados
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'enviarRelatorioDiario'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  // Segunda a sexta às 06h45 (dispara entre 06h e 07h — não tem minutos exatos no Apps Script)
  var dias = [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY,
  ];
  dias.forEach(function(dia) {
    ScriptApp.newTrigger('enviarRelatorioDiario')
      .timeBased()
      .onWeekDay(dia)
      .atHour(7)   // entre 07h e 08h; antes do relatório semanal de segunda
      .create();
  });

  Logger.log('✅ Trigger diário criado: seg-sex às 07h (GMT-3).');
}

// ---------------------------------------------------------------
// HELPERS DE OPERADOR — tratam hífen (' - ') e em-dash (' — ')
// ---------------------------------------------------------------

// Extrai código numérico do campo func (ex: "130 — EDERSON" → "130")
function _rdCodigoOp(func) {
  var s = String(func || '').trim();
  var cod = s.split(/[\s\-—]+/)[0].trim();
  var n = Number(cod);
  return (!isNaN(n) && n > 0) ? String(n) : cod;
}

// Extrai nome limpo do campo func (ex: "130 — EDERSON LUIS" → "EDERSON LUIS")
function _rdNomeOp(func) {
  var s = String(func || '').trim();
  // Remove prefixo numérico e separador (hífen ou em-dash)
  var m = s.match(/^\d+\s*[\-—]\s*(.+)/);
  return m ? m[1].trim() : s;
}

// ---------------------------------------------------------------
// MONTAR RELATÓRIO DIÁRIO
// ---------------------------------------------------------------
function _rdMontarRelatorio() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) throw new Error('Aba "' + ABA_RESPOSTAS + '" não encontrada.');

  const dados = aba.getDataRange().getValues();
  const agora = new Date();

  // Janelas temporais
  const hoje      = new Date(agora); hoje.setHours(0,0,0,0);
  const ontem     = new Date(hoje);  ontem.setDate(hoje.getDate() - 1);
  const hojeFim   = new Date(hoje);  hojeFim.setHours(23,59,59,999);
  const ontemFim  = new Date(ontem); ontemFim.setHours(23,59,59,999);
  // "Esta semana" = últimos 7 dias para tendência rápida
  const semanaIni = new Date(hoje);  semanaIni.setDate(hoje.getDate() - 7);

  // --- Parse linhas (todos os tipos relevantes) ---
  const linhas = [];
  for (var i = 1; i < dados.length; i++) {
    var r = dados[i];
    if (!r[0] || !r[2]) continue;
    var ts = _rsParseData(r[0]);
    if (!ts) continue;
    var tipoRaw = String(r[2]).toUpperCase();
    var tipo = null;
    if      (tipoRaw.includes('ABERTURA') && !tipoRaw.includes('RETRABALHO')) tipo = 'ABERTURA';
    else if (tipoRaw.includes('FECHAMENTO'))                                   tipo = 'FECHAMENTO';
    else if (tipoRaw.includes('INÍCIO DE PARADA'))                             tipo = 'PARADA_INI';
    else if (tipoRaw.includes('INÍCIO DE RETRABALHO'))                         tipo = 'RETRAB_INI';
    else if (tipoRaw.includes('TÉRMINO DE RETRABALHO'))                        tipo = 'RETRAB_FIM';
    else continue;

    var campoF = String(r[5] || '');
    var pts    = campoF.split('|').map(function(x){ return x.trim(); });
    var funcRaw = String(r[1] || '').trim();
    linhas.push({
      ts,
      tipo,
      func:     funcRaw,
      funcCod:  _rdCodigoOp(funcRaw), // código numérico para pareamento robusto
      funcNome: _rdNomeOp(funcRaw),   // nome limpo para exibição
      op:    String(r[3] || '').trim(),
      item:  String(r[4] || '').trim(),
      serie: pts[0] || '',
      impl:  pts[1] || '',
      qty:   Number(r[6])  || 0,
    });
  }

  // --- Pareamento ABERTURA ↔ FECHAMENTO (two-pass greedy, usa código normalizado) ---
  var aberturas = linhas
    .filter(function(r){ return r.tipo === 'ABERTURA'; })
    .map(function(r){ return Object.assign({}, r, { _used: false, _fechTs: null, _leadMs: 0 }); });

  linhas
    .filter(function(r){ return r.tipo === 'FECHAMENTO'; })
    .forEach(function(fech) {
      var melhor = null, melhorScore = -1;
      for (var i = 0; i < aberturas.length; i++) {
        var a = aberturas[i];
        // Usa funcCod para tolerância a hífen vs em-dash e variações de código
        if (a._used || a.funcCod !== fech.funcCod || a.ts > fech.ts) continue;
        var score = 1;
        if (a.op    === fech.op)    score += 4;
        if (a.serie === fech.serie) score += 2;
        if (a.item  === fech.item)  score += 1;
        // Empate: prefere a ABERTURA mais recente (mais próxima do FECHAMENTO)
        if (score > melhorScore || (score === melhorScore && melhor !== null && a.ts > aberturas[melhor].ts)) {
          melhorScore = score; melhor = i;
        }
      }
      if (melhor !== null) {
        aberturas[melhor]._used   = true;
        aberturas[melhor]._fechTs = fech.ts;
        aberturas[melhor]._leadMs = fech.ts - aberturas[melhor].ts;
      }
    });

  var backlog   = aberturas.filter(function(a){ return !a._used; })
                            .sort(function(x, y){ return x.ts - y.ts; }); // mais antigas primeiro
  var fechadas  = aberturas.filter(function(a){ return a._used; });

  // --- Aberturas e fechamentos HOJE e ONTEM ---
  var aberturasHoje  = linhas.filter(function(r){ return r.tipo==='ABERTURA' && r.ts >= hoje && r.ts <= hojeFim; });
  var fechamentosHoje= fechadas.filter(function(r){ return r._fechTs >= hoje && r._fechTs <= hojeFim; });
  var aberturasOntem = linhas.filter(function(r){ return r.tipo==='ABERTURA' && r.ts >= ontem && r.ts <= ontemFim; });
  var fechamentosOntem=fechadas.filter(function(r){ return r._fechTs >= ontem && r._fechTs <= ontemFim; });

  // Prefere ontem se hoje não tem dados ainda (relatório às 07h)
  var refDiaLabel = aberturasHoje.length > 0 || fechamentosHoje.length > 0 ? 'Hoje' : 'Ontem';
  var refAberturas = refDiaLabel === 'Hoje' ? aberturasHoje  : aberturasOntem;
  var refFechados  = refDiaLabel === 'Hoje' ? fechamentosHoje : fechamentosOntem;

  // --- Backlog por operador (quem conversar) — agrupa por código normalizado ---
  var backlogPorOp = {};
  backlog.forEach(function(o){
    var cod  = o.funcCod  || o.func;
    var nome = o.funcNome || o.func;
    if (!backlogPorOp[cod]) backlogPorOp[cod] = { count: 0, ordens: [], nome: nome };
    backlogPorOp[cod].count++;
    backlogPorOp[cod].ordens.push(o);
  });
  var topBacklogOp = Object.values(backlogPorOp)
    .sort(function(a,b){ return b.count - a.count; })
    .slice(0, 5);

  // --- Ordens mais antigas em aberto ---
  var ordensAntigas = backlog.slice(0, 8);

  // --- Backlog por operação (gargalos) ---
  var backlogPorOperacao = {};
  backlog.forEach(function(o){
    var op = o.op || '(sem operação)';
    backlogPorOperacao[op] = (backlogPorOperacao[op] || 0) + 1;
  });
  var topGargalos = Object.entries(backlogPorOperacao)
    .sort(function(a,b){ return b[1]-a[1]; })
    .slice(0, 3);

  // --- Backlog por série (foco em produto) ---
  var backlogPorSerie = {};
  backlog.forEach(function(o){
    if (!o.serie) return;
    if (!backlogPorSerie[o.serie]) backlogPorSerie[o.serie] = { count: 0, impl: o.impl || '' };
    backlogPorSerie[o.serie].count++;
  });
  var topSeries = Object.entries(backlogPorSerie)
    .sort(function(a,b){ return b[1].count-a[1].count; })
    .slice(0, 5);

  // --- Operadores sem atividade nos últimos 2 dias úteis ---
  // Agrupados por código normalizado para evitar duplicatas (hífen vs em-dash)
  var dois_dias_atras = new Date(hoje); dois_dias_atras.setDate(hoje.getDate() - 2);
  var ativosRecentesCod = new Set(
    linhas.filter(function(r){ return r.ts >= dois_dias_atras; })
          .map(function(r){ return r.funcCod || r.func; })
  );
  var todosOpsCod = {}; // cod → nome
  linhas.forEach(function(r){
    var cod = r.funcCod || r.func;
    if (!todosOpsCod[cod]) todosOpsCod[cod] = r.funcNome || r.func;
  });
  var semAtividade = Object.entries(todosOpsCod)
    .filter(function(e){ return !ativosRecentesCod.has(e[0]); })
    .map(function(e){ return e[1]; })
    .sort();

  // --- Retrabalhos ativos ---
  var retrabs = linhas.filter(function(r){ return r.tipo === 'RETRAB_INI'; });
  var retrAbertos = retrabs.filter(function(ri){
    return !linhas.some(function(rf){
      return rf.tipo === 'RETRAB_FIM' && rf.funcCod === ri.funcCod && rf.ts > ri.ts;
    });
  });

  // --- Alertas diários ---
  var alertas = _rdAlertas(backlog, refFechados, refAberturas, topBacklogOp, ordensAntigas, agora);

  // --- Ações recomendadas ---
  var acoes = _rdAcoes(backlog, topBacklogOp, topGargalos, ordensAntigas, semAtividade, refFechados, retrAbertos);

  return {
    agora, hoje, refDiaLabel,
    aberturasRef: refAberturas, fechadosRef: refFechados,
    backlog, topBacklogOp, ordensAntigas, topGargalos, topSeries,
    semAtividade, retrAbertos, alertas, acoes,
  };
}

// ---------------------------------------------------------------
// ALERTAS DIÁRIOS
// ---------------------------------------------------------------
function _rdAlertas(backlog, fechados, aberturas, topOps, antigas, agora) {
  var lista = [];
  var limiteMs = 3 * 86400000; // 3 dias

  if (backlog.length >= 20) {
    lista.push({ nivel: 'CRITICO', msg: 'Backlog crítico: ' + backlog.length + ' ordens em aberto. Intervenção imediata necessária.' });
  } else if (backlog.length >= 10) {
    lista.push({ nivel: 'ATENCAO', msg: 'Backlog elevado: ' + backlog.length + ' ordens em aberto. Monitorar e acelerar fechamentos.' });
  }

  var antigas3d = backlog.filter(function(o){ return (agora - o.ts) > limiteMs; });
  if (antigas3d.length > 0) {
    lista.push({ nivel: 'CRITICO', msg: antigas3d.length + ' ordem(ns) aberta(s) há mais de 3 dias. Investigar imediatamente.' });
  }

  if (fechados.length === 0 && aberturas.length === 0) {
    lista.push({ nivel: 'ATENCAO', msg: 'Nenhum apontamento registrado no dia de referência. Verificar se operadores usaram o sistema.' });
  }

  if (topOps.length > 0 && topOps[0].count >= 5) {
    lista.push({ nivel: 'ATENCAO', msg: topOps[0].nome + ' acumula ' + topOps[0].count + ' ordens em aberto — conversa prioritária hoje.' });
  }

  return lista;
}

// ---------------------------------------------------------------
// AÇÕES RECOMENDADAS PARA O DIA
// ---------------------------------------------------------------
function _rdAcoes(backlog, topOps, topGargalos, antigas, semAtividade, fechados, retrabs) {
  var lista = [];
  var limiteMs = 3 * 86400000;
  var agora = new Date();

  // 1. Conversar com operadores com mais backlog
  if (topOps.length > 0) {
    var nomes = topOps.slice(0, 3).map(function(o){ return o.nome + ' (' + o.count + ')'; }).join(', ');
    lista.push('💬 <b>Conversar hoje com:</b> ' + nomes + ' — entender o que impede o fechamento das ordens em aberto.');
  }

  // 2. Ordens mais antigas
  var antigas3d = backlog.filter(function(o){ return (agora - o.ts) > limiteMs; });
  if (antigas3d.length > 0) {
    var seriesAntigas = [...new Set(antigas3d.map(function(o){ return o.serie; }).filter(Boolean))].slice(0,3).join(', ');
    lista.push('🔴 <b>Prioridade máxima:</b> ' + antigas3d.length + ' ordem(ns) com mais de 3 dias em aberto.'
      + (seriesAntigas ? ' Séries: ' + seriesAntigas + '.' : '')
      + ' Verificar se há impedimento de material, máquina ou informação.');
  }

  // 3. Gargalo de operação
  if (topGargalos.length > 0 && topGargalos[0][1] >= 3) {
    lista.push('🔧 <b>Gargalo detectado:</b> operação <i>' + topGargalos[0][0] + '</i> concentra ' + topGargalos[0][1]
      + ' ordens do backlog. Avaliar reforço de equipe ou resequenciamento nesta etapa.');
  }

  // 4. Operadores sem atividade
  if (semAtividade.length > 0 && semAtividade.length <= 5) {
    lista.push('📋 <b>Sem apontamento nos últimos 2 dias:</b> ' + semAtividade.slice(0,4).join(', ')
      + '. Verificar presença e garantir que os registros estão sendo feitos corretamente.');
  } else if (semAtividade.length > 5) {
    lista.push('📋 <b>' + semAtividade.length + ' operadores</b> sem apontamento nos últimos 2 dias. Verificar aderência ao sistema de apontamento.');
  }

  // 5. Retrabalhos em aberto
  if (retrabs.length > 0) {
    var opRetrabs = [...new Set(retrabs.map(function(r){ return _rdNomeOp(r.func); }))];
    lista.push('🔄 <b>Retrabalho em andamento:</b> ' + opRetrabs.join(', ')
      + '. Acompanhar conclusão e registrar término de retrabalho no sistema.');
  }

  // 6. Desempenho ontem/hoje
  if (fechados.length >= 5) {
    lista.push('✅ <b>Bom ritmo:</b> ' + fechados.length + ' fechamentos registrados. Manter a cadência e garantir novos apontamentos de abertura.');
  } else if (fechados.length > 0) {
    lista.push('📈 <b>Apenas ' + fechados.length + ' fechamento(s).</b> Estimular operadores a concluir e registrar as operações em andamento.');
  }

  if (lista.length === 0) {
    lista.push('✅ Nenhum ponto crítico identificado para hoje. Manter cadência de apontamentos e foco na qualidade dos registros.');
  }

  return lista;
}

// ---------------------------------------------------------------
// GERAÇÃO DO HTML DO EMAIL DIÁRIO
// ---------------------------------------------------------------
function _rdGerarHtml(rel) {
  var { agora, refDiaLabel, aberturasRef, fechadosRef,
        backlog, topBacklogOp, ordensAntigas, topGargalos, topSeries,
        semAtividade, retrAbertos, alertas, acoes } = rel;

  var fmt      = function(d){ return Utilities.formatDate(d, 'GMT-3', 'dd/MM/yyyy'); };
  var fmtHora  = function(d){ return Utilities.formatDate(d, 'GMT-3', 'HH:mm'); };
  var dataEnvio= Utilities.formatDate(agora, 'GMT-3', "dd/MM/yyyy 'às' HH:mm");
  var diaSemana= ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'][agora.getDay()];

  // Cor do backlog
  var corBacklog = backlog.length >= 20 ? '#f87171' : backlog.length >= 10 ? '#fb923c' : '#4ade80';

  // ── HTML Alertas ──
  var htmlAlertas = '';
  if (alertas.length === 0) {
    htmlAlertas = '<tr><td style="padding:10px 14px;color:#4ade80;font-size:13px;">✅ Sem alertas críticos para hoje.</td></tr>';
  } else {
    alertas.forEach(function(a) {
      var cor = a.nivel === 'CRITICO' ? '#f87171' : '#fb923c';
      var ico = a.nivel === 'CRITICO' ? '🔴' : '🟡';
      htmlAlertas += '<tr><td style="padding:9px 14px;border-left:3px solid ' + cor + ';background:#1a2030;border-radius:0 4px 4px 0;font-size:13px;color:#e6edf3;margin-bottom:4px;">' + ico + ' ' + a.msg + '</td></tr>';
    });
  }

  // ── HTML Ações ──
  var htmlAcoes = '';
  acoes.forEach(function(a) {
    htmlAcoes += '<tr><td style="padding:9px 16px;font-size:13px;color:#e6edf3;border-bottom:1px solid #252d40;line-height:1.55;">' + a + '</td></tr>';
  });

  // ── HTML KPIs cards ──
  var card = function(cor, label, valor, sub) {
    return '<td style="width:25%;padding:4px;">'
      + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;border-top:3px solid ' + cor + ';">'
      + '<tr><td style="padding:12px 10px;">'
      + '<div style="font-size:10px;color:#8b949e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">' + label + '</div>'
      + '<div style="font-size:28px;font-weight:900;color:' + cor + ';line-height:1;">' + valor + '</div>'
      + '<div style="font-size:11px;color:#8b949e;margin-top:3px;">' + sub + '</div>'
      + '</td></tr></table></td>';
  };

  var htmlCards = ''
    + card('#fca5a5', 'BACKLOG TOTAL', backlog.length, 'ordens sem fechamento')
    + card('#4ade80', refDiaLabel + ' — ENTREGAS', fechadosRef.length, 'ops concluídas')
    + card('#38bdf8', refDiaLabel + ' — ENTRADAS', aberturasRef.length, 'novas aberturas')
    + card('#fb923c', 'MAIS ANTIGO', ordensAntigas.length > 0
        ? Math.floor((agora - ordensAntigas[0].ts) / 86400000) + 'd'
        : '—',
        ordensAntigas.length > 0 ? 'dias sem fechar' : 'tudo ok');

  // ── HTML Quem conversar ──
  var htmlConversas = '';
  if (topBacklogOp.length === 0) {
    htmlConversas = '<tr><td colspan="2" style="padding:12px;color:#4ade80;font-size:13px;text-align:center;">✅ Nenhum backlog pendente.</td></tr>';
  } else {
    topBacklogOp.forEach(function(op) {
      var corQ = op.count >= 5 ? '#f87171' : op.count >= 3 ? '#fb923c' : '#8b949e';
      var series = [...new Set(op.ordens.map(function(o){ return o.serie; }).filter(Boolean))].slice(0,3).join(', ');
      htmlConversas += '<tr>'
        + '<td style="padding:8px 10px;font-size:13px;color:#e6edf3;border-bottom:1px solid #252d40;">' + op.nome + '</td>'
        + '<td style="padding:8px 10px;font-size:13px;font-weight:700;color:' + corQ + ';border-bottom:1px solid #252d40;text-align:center;">' + op.count + ' ordens</td>'
        + '<td style="padding:8px 10px;font-size:11px;color:#8b949e;border-bottom:1px solid #252d40;">' + (series||'—') + '</td>'
        + '</tr>';
    });
  }

  // ── HTML Ordens Antigas ──
  var htmlAntigas = '';
  if (ordensAntigas.length === 0) {
    htmlAntigas = '<tr><td colspan="4" style="padding:12px;color:#4ade80;font-size:13px;text-align:center;">✅ Nenhuma ordem com mais de 3 dias.</td></tr>';
  } else {
    ordensAntigas.forEach(function(o) {
      var dias     = Math.floor((agora - o.ts) / 86400000);
      var corDias  = dias >= 5 ? '#f87171' : dias >= 3 ? '#fb923c' : '#8b949e';
      var nome     = _rdNomeOp(o.func);
      htmlAntigas += '<tr>'
        + '<td style="padding:6px 10px;font-size:11px;color:#8b949e;border-bottom:1px solid #252d40;">' + fmt(o.ts) + '</td>'
        + '<td style="padding:6px 10px;font-size:12px;color:#e6edf3;border-bottom:1px solid #252d40;">' + nome + '</td>'
        + '<td style="padding:6px 10px;font-size:12px;color:#38bdf8;border-bottom:1px solid #252d40;">' + (o.op||'—') + '</td>'
        + '<td style="padding:6px 10px;font-size:12px;font-weight:700;color:' + corDias + ';border-bottom:1px solid #252d40;text-align:right;">' + dias + 'd</td>'
        + '</tr>';
    });
  }

  // ── HTML Gargalos ──
  var htmlGargalos = '';
  if (topGargalos.length === 0) {
    htmlGargalos = '<tr><td colspan="2" style="padding:12px;color:#4ade80;font-size:13px;text-align:center;">Sem gargalos identificados.</td></tr>';
  } else {
    var maxG = topGargalos[0][1];
    topGargalos.forEach(function(g, i) {
      var pct = Math.max(8, Math.round(g[1] / maxG * 100));
      var cor = i === 0 ? '#f87171' : '#fb923c';
      htmlGargalos += '<tr>'
        + '<td style="padding:8px 12px;font-size:12px;color:#e6edf3;border-bottom:1px solid #252d40;width:50%;">' + g[0] + '</td>'
        + '<td style="padding:8px 12px;border-bottom:1px solid #252d40;">'
        + '<div style="display:flex;align-items:center;gap:8px;">'
        + '<div style="flex:1;height:10px;background:#252d40;border-radius:3px;overflow:hidden;">'
        + '<div style="width:' + pct + '%;height:100%;background:' + cor + ';border-radius:3px;"></div></div>'
        + '<span style="font-size:12px;font-weight:700;color:' + cor + ';min-width:24px;text-align:right;">' + g[1] + '</span>'
        + '</div></td>'
        + '</tr>';
    });
  }

  // ── HTML Séries com Backlog ──
  var htmlSeries = '';
  if (topSeries.length === 0) {
    htmlSeries = '<tr><td colspan="3" style="padding:12px;color:#4ade80;font-size:13px;text-align:center;">Sem backlog por série.</td></tr>';
  } else {
    topSeries.forEach(function(entry) {
      var s = entry[0], v = entry[1];
      htmlSeries += '<tr>'
        + '<td style="padding:7px 10px;font-size:13px;font-weight:700;color:#38bdf8;border-bottom:1px solid #252d40;">' + s + '</td>'
        + '<td style="padding:7px 10px;font-size:12px;color:#8b949e;border-bottom:1px solid #252d40;">' + (v.impl||'—') + '</td>'
        + '<td style="padding:7px 10px;font-size:12px;font-weight:700;color:#fca5a5;border-bottom:1px solid #252d40;text-align:right;">' + v.count + ' abertas</td>'
        + '</tr>';
    });
  }

  // ── MONTAGEM FINAL ──
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="margin:0;padding:0;background:#0a0e17;font-family:Arial,Helvetica,sans-serif;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#0a0e17;padding:24px 12px;">'
    + '<tr><td align="center">'
    + '<table width="620" cellpadding="0" cellspacing="0" style="max-width:620px;width:100%;">'

    // CABEÇALHO
    + '<tr><td style="background:#0d1117;border-bottom:3px solid #38bdf8;padding:20px 28px;border-radius:12px 12px 0 0;">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
    + '<td><div style="font-size:26px;font-weight:900;color:#f0b429;letter-spacing:3px;line-height:1;">AGRICEF</div>'
    + '<div style="font-size:11px;color:#58677a;letter-spacing:2px;margin-top:3px;text-transform:uppercase;">Relatório Operacional Diário</div></td>'
    + '<td align="right">'
    + '<div style="font-size:15px;font-weight:700;color:#38bdf8;">🗓️ ' + diaSemana + ', ' + fmt(agora) + '</div>'
    + '<div style="font-size:11px;color:#58677a;margin-top:3px;">Gerado às ' + fmtHora(agora) + ' (Brasília)</div>'
    + '</td></tr></table></td></tr>'

    // ALERTAS
    + '<tr><td style="background:#111722;padding:16px 28px 4px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f87171;text-transform:uppercase;letter-spacing:1.5px;padding:0 0 8px;">🚨 Pontos Críticos de Hoje</div>'
    + '<table width="100%" cellpadding="0" cellspacing="3">' + htmlAlertas + '</table>'
    + '</td></tr>'

    // AÇÕES DO DIA
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<div style="font-size:12px;font-weight:700;color:#4ade80;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">⚡ Ações para Hoje</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">' + htmlAcoes + '</table>'
    + '</td></tr>'

    // KPI CARDS
    + '<tr><td style="background:#111722;padding:4px 28px 12px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f0b429;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">📊 Situação do Dia</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>' + htmlCards + '</tr></table>'
    + '</td></tr>'

    // QUEM CONVERSAR
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<div style="font-size:12px;font-weight:700;color:#fb923c;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">💬 Com Quem Falar Hoje (Top Backlog por Operador)</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">'
    + '<tr style="background:#141a28;">'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;border-bottom:1px solid #252d40;">OPERADOR</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:center;border-bottom:1px solid #252d40;">BACKLOG</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;border-bottom:1px solid #252d40;">SÉRIES</th>'
    + '</tr>'
    + htmlConversas
    + '</table></td></tr>'

    // ORDENS MAIS ANTIGAS
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<div style="font-size:12px;font-weight:700;color:#f87171;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">⏰ Ordens Mais Antigas em Aberto</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">'
    + '<tr style="background:#141a28;">'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;border-bottom:1px solid #252d40;">ABERTURA</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;border-bottom:1px solid #252d40;">OPERADOR</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:left;border-bottom:1px solid #252d40;">OPERAÇÃO</th>'
    + '<th style="padding:8px 10px;font-size:11px;color:#8b949e;text-align:right;border-bottom:1px solid #252d40;">DIAS</th>'
    + '</tr>'
    + htmlAntigas
    + '</table></td></tr>'

    // GARGALOS + SÉRIES
    + '<tr><td style="background:#111722;padding:4px 28px 16px;">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'

    + '<td style="width:50%;vertical-align:top;padding-right:8px;">'
    + '<div style="font-size:12px;font-weight:700;color:#a78bfa;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">🔧 Gargalos (Op. com Mais Backlog)</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">' + htmlGargalos + '</table>'
    + '</td>'

    + '<td style="width:50%;vertical-align:top;padding-left:8px;">'
    + '<div style="font-size:12px;font-weight:700;color:#67e8f9;text-transform:uppercase;letter-spacing:1.5px;padding:10px 0 8px;">📦 Séries com Mais Backlog</div>'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#1c2230;border:1px solid #252d40;border-radius:8px;">' + htmlSeries + '</table>'
    + '</td>'

    + '</tr></table></td></tr>'

    // FOOTER
    + '<tr><td style="background:#0d1117;padding:14px 28px;border-radius:0 0 12px 12px;border-top:1px solid #1c2230;text-align:center;">'
    + '<div style="font-size:11px;color:#58677a;">Relatório gerado automaticamente de segunda a sexta às 07h (Brasília).</div>'
    + '<div style="font-size:11px;color:#3a4460;margin-top:3px;">AGRICEF — Dashboard de Produção · Sistema de Apontamento v4 '
    + '| <a href="https://agricefprocessos-tech.github.io/agricef-dashboard/" style="color:#38bdf8;text-decoration:none;">Abrir Dashboard</a></div>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}

// ---------------------------------------------------------------
// ENVIO DO EMAIL DIÁRIO
// ---------------------------------------------------------------
function _rdEnviarEmail(rel) {
  var dataStr = Utilities.formatDate(rel.agora, 'GMT-3', 'dd/MM/yyyy');
  var html = _rdGerarHtml(rel);
  MailApp.sendEmail({
    to:       EMAIL_RELATORIO,
    subject:  '⚡ AGRICEF | Relatório Diário — ' + dataStr,
    htmlBody: html,
  });
}

// ================================================================
// ANÁLISE DE ÓRFÃOS (ABERTURAs sem FECHAMENTO correspondente)
// Execute via: ?action=analyzeOrphans&key=AGF2026
// ================================================================
function analisarOrfaos() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) throw new Error('Aba Respostas não encontrada.');

  const dados = aba.getDataRange().getValues();
  const agora = new Date();

  // Parse de todas as linhas de ABERTURA e FECHAMENTO
  const linhas = [];
  for (var i = 1; i < dados.length; i++) {
    var r = dados[i];
    if (!r[0] || !r[2]) continue;
    var ts = new Date(r[0]);
    if (isNaN(ts.getTime())) continue;
    var tipoRaw = String(r[2]).toUpperCase();
    var tipo = null;
    if (tipoRaw.includes('ABERTURA') && !tipoRaw.includes('RETRABALHO')) tipo = 'ABERTURA';
    else if (tipoRaw.includes('FECHAMENTO')) tipo = 'FECHAMENTO';
    else continue;
    var funcRaw = String(r[1] || '').trim();
    var campoF  = String(r[5] || '');
    var pts     = campoF.split('|').map(function(x){ return x.trim(); });
    linhas.push({
      ts:      ts,
      tipo:    tipo,
      func:    funcRaw,
      funcCod: _rdCodigoOp(funcRaw),
      op:      String(r[3] || '').trim(),
      item:    String(r[4] || '').trim(),
      serie:   pts[0] || '',
      impl:    pts[1] || '',
      row:     i + 1,  // linha na planilha (base 1, +1 pelo cabeçalho)
    });
  }

  // Algoritmo de pareamento idêntico ao relatório
  var aberturas = linhas.filter(function(r){ return r.tipo === 'ABERTURA'; })
    .map(function(r){ return Object.assign({}, r, { _used: false }); });

  linhas.filter(function(r){ return r.tipo === 'FECHAMENTO'; })
    .forEach(function(fech){
      var melhor = null, melhorScore = -1;
      for (var i = 0; i < aberturas.length; i++) {
        var a = aberturas[i];
        if (a._used || a.funcCod !== fech.funcCod || a.ts > fech.ts) continue;
        var score = 1;
        if (a.op    === fech.op)    score += 4;
        if (a.serie === fech.serie) score += 2;
        if (a.item  === fech.item)  score += 1;
        // Empate: prefere a ABERTURA mais recente (mais próxima do FECHAMENTO)
        if (score > melhorScore || (score === melhorScore && melhor !== null && a.ts > aberturas[melhor].ts)) {
          melhorScore = score; melhor = i;
        }
      }
      if (melhor !== null) aberturas[melhor]._used = true;
    });

  var orfaos = aberturas.filter(function(a){ return !a._used; });

  // Agrupa órfãos por faixa de data
  var cortes = {
    antes_fev2026:  new Date('2026-02-01T00:00:00'),
    antes_mar2026:  new Date('2026-03-01T00:00:00'),
    antes_abr2026:  new Date('2026-04-01T00:00:00'),
    antes_mai2026:  new Date('2026-05-01T00:00:00'),
  };

  var porFaixa = {
    'Antes de Fev/2026':  [],
    'Fev/2026':           [],
    'Mar/2026':           [],
    'Abr/2026':           [],
    'Mai/2026 em diante': [],
  };

  orfaos.forEach(function(o){
    var d = o.ts;
    if      (d < cortes.antes_fev2026) porFaixa['Antes de Fev/2026'].push(o);
    else if (d < cortes.antes_mar2026) porFaixa['Fev/2026'].push(o);
    else if (d < cortes.antes_abr2026) porFaixa['Mar/2026'].push(o);
    else if (d < cortes.antes_mai2026) porFaixa['Abr/2026'].push(o);
    else                               porFaixa['Mai/2026 em diante'].push(o);
  });

  // Agrupa por operador
  var porOperador = {};
  orfaos.forEach(function(o){
    var k = o.funcCod || o.func;
    if (!porOperador[k]) porOperador[k] = { func: o.func, funcCod: o.funcCod, count: 0, datas: [] };
    porOperador[k].count++;
    var dtStr = Utilities.formatDate(o.ts, 'GMT-3', 'dd/MM/yyyy');
    if (porOperador[k].datas.indexOf(dtStr) === -1) porOperador[k].datas.push(dtStr);
  });

  // Agrupa por operação
  var porOperacao = {};
  orfaos.forEach(function(o){
    var k = o.op || '(sem operação)';
    porOperacao[k] = (porOperacao[k] || 0) + 1;
  });

  // Impacto na assiduidade: dias únicos por operador contribuídos apenas por órfãos
  // (dias em que o operador SÓ tem órfão, sem outros registros)
  var diasPorOp = {};   // todos os registros (para comparar)
  var diasOrfaoPorOp = {};
  linhas.forEach(function(r){
    var k = r.funcCod || r.func;
    var d = Utilities.formatDate(r.ts, 'GMT-3', 'dd/MM/yyyy');
    if (!diasPorOp[k]) diasPorOp[k] = new Set();
    diasPorOp[k].add(d);
  });
  orfaos.forEach(function(o){
    var k = o.funcCod || o.func;
    var d = Utilities.formatDate(o.ts, 'GMT-3', 'dd/MM/yyyy');
    if (!diasOrfaoPorOp[k]) diasOrfaoPorOp[k] = new Set();
    diasOrfaoPorOp[k].add(d);
  });

  // Linhas de não-órfãos por operador e data
  var diasComOutros = {};
  var orfaoRows = new Set(orfaos.map(function(o){ return o.row; }));
  linhas.forEach(function(r){
    if (orfaoRows.has(r.row)) return;  // é órfão ele mesmo — pula
    var k = r.funcCod || r.func;
    var d = Utilities.formatDate(r.ts, 'GMT-3', 'dd/MM/yyyy');
    if (!diasComOutros[k]) diasComOutros[k] = new Set();
    diasComOutros[k].add(d);
  });

  // Dias que seriam perdidos na assiduidade (só existem por causa do órfão)
  var diasPerdidosTotal = 0;
  var assiduidadeImpacto = [];
  Object.keys(diasOrfaoPorOp).forEach(function(k){
    var perdidos = 0;
    diasOrfaoPorOp[k].forEach(function(d){
      if (!diasComOutros[k] || !diasComOutros[k].has(d)) perdidos++;
    });
    if (perdidos > 0) {
      diasPerdidosTotal += perdidos;
      assiduidadeImpacto.push({ funcCod: k, diasPerdidos: perdidos });
    }
  });

  // Resumo por faixa (simplificado para o JSON de resposta)
  var resumoFaixas = {};
  Object.keys(porFaixa).forEach(function(k){
    resumoFaixas[k] = porFaixa[k].length;
  });

  // Top operadores por qtd de órfãos
  var topOps = Object.values(porOperador)
    .sort(function(a,b){ return b.count - a.count; })
    .slice(0, 15)
    .map(function(o){ return { func: o.func, funcCod: o.funcCod, count: o.count }; });

  // Top operações
  var topOps2 = Object.entries(porOperacao)
    .sort(function(a,b){ return b[1]-a[1]; })
    .map(function(e){ return { operacao: e[0], count: e[1] }; });

  // Data do órfão mais antigo e mais recente
  var datas = orfaos.map(function(o){ return o.ts.getTime(); });
  var maisAntigo = datas.length ? new Date(Math.min.apply(null, datas)) : null;
  var maisRecente = datas.length ? new Date(Math.max.apply(null, datas)) : null;

  return {
    success: true,
    totais: {
      totalRegistros:    linhas.length,
      totalAberturas:    aberturas.length,
      totalFechamentos:  linhas.filter(function(r){ return r.tipo === 'FECHAMENTO'; }).length,
      totalPareados:     aberturas.filter(function(a){ return a._used; }).length,
      totalOrfaos:       orfaos.length,
      maisAntigo:        maisAntigo ? Utilities.formatDate(maisAntigo, 'GMT-3', 'dd/MM/yyyy') : null,
      maisRecente:       maisRecente ? Utilities.formatDate(maisRecente, 'GMT-3', 'dd/MM/yyyy') : null,
    },
    porFaixa:           resumoFaixas,
    porOperador:        topOps,
    porOperacao:        topOps2,
    impactoAssiduidade: {
      operadoresAfetados: assiduidadeImpacto.length,
      diasPerdidosTotal:  diasPerdidosTotal,
      detalhe:            assiduidadeImpacto,
    },
    geradoEm: Utilities.formatDate(agora, 'GMT-3', 'dd/MM/yyyy HH:mm:ss'),
  };
}

// ================================================================
// MARCAR ÓRFÃOS COMO LEGADO (adiciona marcação na coluna de obs)
// Parâmetro cutoff: data no formato "YYYY-MM-DD" (ex: "2026-03-01")
// Marca todos os órfãos ANTERIORES à data informada.
// Execute via: ?action=marcarLegado&key=AGF2026&cutoff=2026-03-01
// ================================================================
function marcarOrfaosLegado(cutoffStr) {
  var cutoff = cutoffStr ? new Date(cutoffStr + 'T00:00:00') : null;
  if (!cutoff || isNaN(cutoff.getTime())) throw new Error('Parâmetro cutoff inválido. Use formato YYYY-MM-DD.');

  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) throw new Error('Aba Respostas não encontrada.');

  var dados = aba.getDataRange().getValues();

  // Mesmo algoritmo de pareamento para identificar órfãos
  var linhas = [];
  for (var i = 1; i < dados.length; i++) {
    var r = dados[i];
    if (!r[0] || !r[2]) continue;
    var ts = new Date(r[0]);
    if (isNaN(ts.getTime())) continue;
    var tipoRaw = String(r[2]).toUpperCase();
    var tipo = null;
    if (tipoRaw.includes('ABERTURA') && !tipoRaw.includes('RETRABALHO')) tipo = 'ABERTURA';
    else if (tipoRaw.includes('FECHAMENTO')) tipo = 'FECHAMENTO';
    else continue;
    var funcRaw = String(r[1] || '').trim();
    var pts = String(r[5] || '').split('|').map(function(x){ return x.trim(); });
    linhas.push({ ts, tipo, funcCod: _rdCodigoOp(funcRaw), op: String(r[3]||'').trim(),
                  serie: pts[0]||'', item: String(r[4]||'').trim(), row: i+1, _used: false });
  }

  var aberturas = linhas.filter(function(r){ return r.tipo === 'ABERTURA'; });
  linhas.filter(function(r){ return r.tipo === 'FECHAMENTO'; }).forEach(function(fech){
    var melhor = null, melhorScore = -1;
    for (var i = 0; i < aberturas.length; i++) {
      var a = aberturas[i];
      if (a._used || a.funcCod !== fech.funcCod || a.ts > fech.ts) continue;
      var score = 1;
      if (a.op === fech.op) score += 4;
      if (a.serie === fech.serie) score += 2;
      if (a.item === fech.item) score += 1;
      // Empate: prefere a ABERTURA mais recente (mais próxima do FECHAMENTO)
      if (score > melhorScore || (score === melhorScore && melhor !== null && a.ts > aberturas[melhor].ts)) {
        melhorScore = score; melhor = i;
      }
    }
    if (melhor !== null) aberturas[melhor]._used = true;
  });

  var orfaosAntigos = aberturas.filter(function(a){ return !a._used && a.ts < cutoff; });

  if (orfaosAntigos.length === 0) return { success: true, message: 'Nenhum órfão anterior a ' + cutoffStr + ' encontrado.', marcados: 0 };

  // Coluna N (índice 13) = OBSERVAÇÃO 2. Adiciona prefixo [LEGADO] se ainda não tiver.
  var colN = aba.getRange(1, 14, dados.length, 1).getValues();
  var marcados = 0;
  orfaosAntigos.forEach(function(o){
    var idx = o.row - 1; // base 0 no array colN (que começa na linha 1 da planilha)
    var atual = String(colN[idx][0] || '');
    if (!atual.includes('[LEGADO]')) {
      colN[idx][0] = '[LEGADO] ' + atual;
      marcados++;
    }
  });

  aba.getRange(1, 14, dados.length, 1).setValues(colN);
  return { success: true, message: marcados + ' órfão(s) anteriores a ' + cutoffStr + ' marcados como [LEGADO].', marcados: marcados };
}

// ================================================================
// LIMPAR REGISTROS DE TESTE — remove linhas com 'TESTE-AUTO' nas obs
// Execute via: ?action=limparTestes&key=AGF2026
// ================================================================
function limparRegistrosTeste() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return 0;
  const dados = aba.getDataRange().getValues();
  if (dados.length <= 1) return 0; // só cabeçalho

  const cabecalho = dados[0];
  const linhasValidas = [cabecalho]; // sempre mantém o cabeçalho
  let removidos = 0;

  // Estratégia eficiente: filtrar → limpar tudo → reescrever
  // Muito mais rápido que deleteRow() linha a linha
  for (let i = 1; i < dados.length; i++) {
    const obs1 = String(dados[i][7]  || '');
    const obs2 = String(dados[i][13] || '');
    const nome = String(dados[i][1]  || ''); // Col B — nome operador
    const isTest = obs1.includes('TESTE-AUTO') || obs2.includes('TESTE-AUTO')
                || obs1.includes('AUTOTESTE')  || obs2.includes('AUTOTESTE')
                || nome.includes('AUTOTESTE')
                || nome.includes('HUMANO')     // registros legados sem nome real
                || obs1.includes('CLEANUP')    || obs2.includes('CLEANUP');
    if (isTest) {
      removidos++;
    } else {
      linhasValidas.push(dados[i]);
    }
  }

  if (removidos === 0) return 0;

  // Limpar tudo e reescrever apenas as linhas válidas
  aba.clearContents();
  aba.getRange(1, 1, linhasValidas.length, cabecalho.length).setValues(linhasValidas);
  SpreadsheetApp.flush();

  Logger.log('🧹 Registros de teste removidos: ' + removidos + ' | Válidas mantidas: ' + (linhasValidas.length - 1));
  return removidos;
}

// ================================================================
// REMOVER REGISTRO POR ABERTOid — deleta linha em Respostas (col P)
// e remove da aba Abertos (col 12). Usado para limpar registros
// de teste que não têm marcador AUTOTESTE.
// Execute via: ?action=removerRegistroPorId&key=AGF2026&id=AP-XXXXXX
// ================================================================
function removerRegistroPorAbertoId(idAlvo) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaRe = ss.getSheetByName(ABA_RESPOSTAS);
  const abaAb = garantirAbaAbertos(ss);
  let removidosRe = 0;
  let removidosAb = 0;

  // Remove de Respostas (col P = índice 15)
  if (abaRe) {
    const dadosRe = abaRe.getDataRange().getValues();
    const linhasValidas = [dadosRe[0]]; // mantém cabeçalho
    for (let i = 1; i < dadosRe.length; i++) {
      if (String(dadosRe[i][15] || '').trim() === idAlvo) {
        removidosRe++;
      } else {
        linhasValidas.push(dadosRe[i]);
      }
    }
    if (removidosRe > 0) {
      abaRe.clearContents();
      abaRe.getRange(1, 1, linhasValidas.length, dadosRe[0].length).setValues(linhasValidas);
    }
  }

  // Remove de Abertos (col 12 = índice 12)
  if (abaAb) {
    const dadosAb = abaAb.getDataRange().getValues();
    for (let i = dadosAb.length - 1; i >= 1; i--) {
      if (String(dadosAb[i][12] || '').trim() === idAlvo) {
        abaAb.deleteRow(i + 1);
        removidosAb++;
      }
    }
  }

  SpreadsheetApp.flush();
  return {
    removidosRespostas: removidosRe,
    removidosAbertos: removidosAb,
    mensagem: `ID ${idAlvo}: ${removidosRe} linha(s) removida(s) de Respostas, ${removidosAb} de Abertos.`
  };
}

// ================================================================
// REMOVER REGISTROS SEM TIMESTAMP — deleta linhas onde col A está vazia.
// Usa deleteRow de baixo para cima (evita deslocamento de índices).
// NÃO usa clearContents — opera célula a célula para máxima segurança.
// Execute via: ?action=removerSemTimestamp&key=AGF2026
// ================================================================
function removerRegistrosSemTimestamp() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba não encontrada.' };

  const dados = aba.getDataRange().getValues();
  const totalAntes = dados.length - 1; // exclui cabeçalho

  // Coleta linhas a deletar (de cima para baixo, mas deleta de baixo para cima)
  const linhasParaDeletar = [];
  for (let i = 1; i < dados.length; i++) {
    const colA = dados[i][0];
    const vazio = (!colA) ||
      (colA instanceof Date && colA.getFullYear() <= 1900) ||
      (typeof colA === 'string' && colA.trim() === '');
    if (vazio) {
      linhasParaDeletar.push(i + 1); // número real da linha na planilha (1-indexed)
    }
  }

  // Deleta de baixo para cima para não deslocar índices
  for (let k = linhasParaDeletar.length - 1; k >= 0; k--) {
    aba.deleteRow(linhasParaDeletar[k]);
  }

  if (linhasParaDeletar.length > 0) SpreadsheetApp.flush();

  // Verifica o estado final
  const totalDepois = aba.getLastRow() - 1; // exclui cabeçalho

  return {
    success: true,
    removidos: linhasParaDeletar.length,
    totalAntes,
    totalDepois,
    mensagem: linhasParaDeletar.length + ' registro(s) sem timestamp removido(s). ' +
              totalDepois + ' registros válidos mantidos.'
  };
}

// ================================================================
// REVERTER TIMESTAMPS INTERPOLADOS — desfaz o preencherTimestampsVazios
// Remove timestamps que são idênticos ao registro imediatamente anterior
// (padrão característico de interpolação). Restaura a célula para vazio.
// Execute via: ?action=reverterTimestamps&key=AGF2026
// ================================================================
function reverterTimestampsInterpolados() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba não encontrada.' };

  const dados = aba.getDataRange().getValues();
  let revertidos = 0;
  const updates = []; // { row (1-indexed) }

  for (let i = 1; i < dados.length; i++) {
    const colA = dados[i][0];
    const colAPrev = dados[i - 1][0];
    if (!colA || !colAPrev) continue;

    // Converte para string para comparação
    let tsAtual = '', tsPrev = '';
    if (colA instanceof Date)    tsAtual = Utilities.formatDate(colA,    'GMT-3', 'dd/MM/yyyy HH:mm:ss');
    else if (typeof colA === 'string') tsAtual = colA.trim();
    if (colAPrev instanceof Date)  tsPrev  = Utilities.formatDate(colAPrev, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
    else if (typeof colAPrev === 'string') tsPrev = colAPrev.trim();

    if (tsAtual && tsPrev && tsAtual === tsPrev) {
      updates.push({ row: i + 1 });
      revertidos++;
    }
  }

  // Limpa os timestamps interpolados (seta célula para string vazia)
  for (const upd of updates) {
    aba.getRange(upd.row, 1).setValue('');
  }
  if (revertidos > 0) SpreadsheetApp.flush();

  return {
    success: true,
    revertidos,
    mensagem: revertidos + ' timestamp(s) interpolado(s) removido(s). Células voltaram para vazio.'
  };
}

// ================================================================
// PREENCHER TIMESTAMPS VAZIOS — interpolação pelo registro anterior
// Preenche col A vazia com o timestamp do último registro com data.
// Registros sem vizinho anterior são deixados em branco.
// Execute via: ?action=preencherTimestamps&key=AGF2026
// ================================================================
function preencherTimestampsVazios() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const dados = aba.getDataRange().getValues();
  let ultimoTs = null;
  let preenchidos = 0;
  const updates = []; // { row (1-indexed), ts }

  for (let i = 1; i < dados.length; i++) {
    const colA = dados[i][0];
    // Verifica se há timestamp válido nesta célula
    let tsStr = '';
    if (colA instanceof Date && colA.getFullYear() > 1900) {
      tsStr = Utilities.formatDate(colA, 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
    } else if (typeof colA === 'string' && colA.trim()) {
      tsStr = colA.trim();
    }

    if (tsStr) {
      ultimoTs = tsStr;
    } else if (ultimoTs) {
      // Célula vazia — preenche com o último timestamp conhecido
      updates.push({ row: i + 1, ts: ultimoTs }); // +1 porque dados[0] é header (linha 1)
      preenchidos++;
    }
  }

  // Aplica os updates linha a linha
  // (não usa setValues em batch pois as linhas não são contíguas)
  for (const upd of updates) {
    aba.getRange(upd.row, 1).setValue(upd.ts);
  }

  if (preenchidos > 0) SpreadsheetApp.flush();

  return {
    success: true,
    preenchidos,
    mensagem: preenchidos + ' timestamp(s) preenchido(s) por interpolação do registro anterior.'
  };
}

// ================================================================
// LISTAR REVISÕES DA PLANILHA — via Drive REST API (UrlFetchApp)
// Não requer re-autorização do DriveApp service.
// Execute via: ?action=listarRevisoes&key=AGF2026
// ================================================================
function listarRevisoesPlanilha() {
  // Usa Drive Advanced Service (já autorizado junto com o scope drive)
  const lista = [];
  let pageToken = '';
  do {
    const params = { fields: 'revisions(id,modifiedTime,lastModifyingUser/emailAddress),nextPageToken', pageSize: 1000 };
    if (pageToken) params.pageToken = pageToken;
    const resp = Drive.Revisions.list(SPREADSHEET_ID, params);
    (resp.revisions || []).forEach(r => lista.push({
      id: r.id,
      data: r.modifiedTime ? Utilities.formatDate(new Date(r.modifiedTime), 'GMT-3', 'dd/MM/yyyy HH:mm:ss') : '',
      autor: (r.lastModifyingUser || {}).emailAddress || 'desconhecido'
    }));
    pageToken = resp.nextPageToken || '';
  } while (pageToken);

  lista.reverse(); // mais recentes primeiro
  return { success: true, total: lista.length, ultimas30: lista.slice(0, 30) };
}

// ================================================================
// RECUPERAR TIMESTAMPS DE UMA REVISÃO — exporta a revisão como CSV
// e extrai a coluna A (timestamps) mapeando pelo ID da col P
// Execute via: ?action=recuperarTimestamps&key=AGF2026&revId=XXX
// ================================================================
function recuperarTimestampsDaRevisao(revId) {
  const token = ScriptApp.getOAuthToken();
  const headers = { Authorization: 'Bearer ' + token };

  // Verifica se a revisão existe via Drive REST API
  const revUrl = 'https://www.googleapis.com/drive/v3/files/' + SPREADSHEET_ID + '/revisions/' + revId +
    '?fields=id,modifiedTime';
  const revResp = UrlFetchApp.fetch(revUrl, { headers, muteHttpExceptions: true });
  if (revResp.getResponseCode() !== 200) {
    return { success: false, message: 'Revisão ' + revId + ' não encontrada. HTTP ' + revResp.getResponseCode() };
  }
  const revInfo = JSON.parse(revResp.getContentText());
  const dataRevisao = revInfo.modifiedTime ?
    Utilities.formatDate(new Date(revInfo.modifiedTime), 'GMT-3', 'dd/MM/yyyy HH:mm:ss') : '?';
  // (revAlvo substituído por revInfo/dataRevisao acima — nada a verificar aqui)

  // Exporta a revisão como CSV (primeira aba)
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
    '/export?format=csv&gid=971425155&revision=' + revId;
  const options = {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(exportUrl, options);
  if (resp.getResponseCode() !== 200) {
    return { success: false, message: 'Erro ao exportar revisão: HTTP ' + resp.getResponseCode() };
  }

  // Parseia o CSV linha a linha
  const csv = resp.getContentText();
  const linhas = Utilities.parseCsv(csv);
  if (linhas.length < 2) return { success: false, message: 'Revisão vazia.' };

  // Mapeia: abertoId (col P = index 15) → timestamp (col A = index 0)
  const mapa = {};
  let comTimestamp = 0;
  for (let i = 1; i < linhas.length; i++) {
    const row = linhas[i];
    const ts = String(row[0] || '').trim();
    const id = String(row[15] || '').trim();
    if (ts && id) {
      mapa[id] = ts;
      comTimestamp++;
    }
  }

  return {
    success: true,
    revisao: revId,
    dataRevisao: dataRevisao,
    totalLinhas: linhas.length - 1,
    comTimestamp,
    amostra: Object.entries(mapa).slice(0, 5).map(([id, ts]) => ({ id, ts }))
  };
}

// ================================================================
// APLICAR TIMESTAMPS RECUPERADOS — preenche col A vazia usando
// o mapa de abertoId→timestamp de uma revisão anterior
// Execute via: ?action=aplicarTimestamps&key=AGF2026&revId=XXX
// ================================================================
function aplicarTimestampsDaRevisao(revId) {
  // 1. Obtém o mapa da revisão
  const mapaResult = recuperarTimestampsDaRevisao(revId);
  if (!mapaResult.success) return mapaResult;

  // Reconstrói o mapa a partir do CSV da revisão
  const tokenApp2 = ScriptApp.getOAuthToken();
  const exportUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
    '/export?format=csv&gid=971425155&revision=' + revId;
  const options = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  const csv = UrlFetchApp.fetch(exportUrl, options).getContentText();
  const linhas = Utilities.parseCsv(csv);
  const mapa = {};
  for (let i = 1; i < linhas.length; i++) {
    const ts = String(linhas[i][0] || '').trim();
    const id = String(linhas[i][15] || '').trim();
    if (ts && id) mapa[id] = ts;
  }

  // 2. Lê Respostas atual e preenche timestamps vazios
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const dados = aba.getDataRange().getValues();
  let preenchidos = 0;
  const updates = []; // [row, timestamp]

  for (let i = 1; i < dados.length; i++) {
    const tsAtual = String(dados[i][0] || '').trim();
    if (tsAtual) continue; // já tem timestamp
    const id = String(dados[i][15] || '').trim();
    if (id && mapa[id]) {
      updates.push({ row: i + 1, ts: mapa[id] });
      preenchidos++;
    }
  }

  // Aplica os updates em batch
  for (const upd of updates) {
    aba.getRange(upd.row, 1).setValue(upd.ts);
  }
  if (updates.length > 0) SpreadsheetApp.flush();

  return {
    success: true,
    revisaoUsada: revId,
    preenchidos,
    mensagem: preenchidos + ' timestamp(s) recuperado(s) da revisão ' + revId
  };
}

// ================================================================
// ANÁLISE DE INCONSISTÊNCIAS — replica a lógica da aba "Inconsistências"
// do dashboard AGRICEF para identificação server-side
// Execute via: ?action=analisarInconsistencias&key=AGF2026
// ================================================================
function analisarInconsistencias() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return { success: false, message: 'Sem dados.' };

  // ── PARSE REGISTROS ──────────────────────────────────────────
  // Colunas: A=ts, B=func, C=tipo, D=op, E=item, F=serie, K=tipoParada, I=tipoRet
  const records = [];
  for (let i = 1; i < dados.length; i++) {
    const row  = dados[i];
    const tsRaw = row[0];
    const func  = String(row[1] || '').trim();
    const tipo  = String(row[2] || '').trim();
    const op    = String(row[3] || '').trim();
    const item  = String(row[4] || '').trim();
    const serieRaw = String(row[5] || '').trim();
    const serie = serieRaw.split(' | ')[0].trim();

    if (!func || !tipo) continue;

    let tsMs;
    if (tsRaw instanceof Date) {
      tsMs = tsRaw.getTime();
    } else if (tsRaw) {
      tsMs = new Date(String(tsRaw)).getTime();
    }
    if (!tsMs || isNaN(tsMs)) continue;

    records.push({ tsMs, func, tipo, op, item, serie });
  }

  // Ordena por timestamp (mesmo que o dashboard faz ao iterar)
  records.sort((a, b) => a.tsMs - b.tsMs);

  // ── WORK MINUTES (replica dashboard: seg-sex 08:00-17:00, pausa 12:00-13:00) ──
  function workMinutes(tsStart, tsEnd) {
    if (tsEnd <= tsStart) return 0;
    const WS = 480, WE = 1020, LS = 720, LE = 780; // minutos do dia
    function mid(ms) { const d = new Date(ms); return d.getHours()*60 + d.getMinutes(); }
    function wmd(s, e) {
      s = Math.max(s, WS); e = Math.min(e, WE);
      if (e <= s) return 0;
      return e - s - Math.max(0, Math.min(e, LE) - Math.max(s, LS));
    }
    function isWD(ms) { const d = new Date(ms); return d.getDay() >= 1 && d.getDay() <= 5; }
    function nextWDStart(ms) {
      const d = new Date(ms);
      d.setHours(8, 0, 0, 0);
      d.setDate(d.getDate() + 1);
      while (new Date(d).getDay() < 1 || new Date(d).getDay() > 5) d.setDate(d.getDate() + 1);
      return d.getTime();
    }
    let tot = 0, cur = tsStart, it = 0;
    const mCur = mid(cur);
    if (!isWD(cur) || mCur >= WE) { cur = nextWDStart(cur); }
    else if (mCur < WS) { const d = new Date(cur); d.setHours(8, 0, 0, 0); cur = d.getTime(); }
    while (cur < tsEnd && it++ < 500) {
      if (!isWD(cur)) { cur = nextWDStart(cur); continue; }
      const eod = new Date(cur); eod.setHours(17, 0, 0, 0);
      tot += wmd(mid(cur), mid(tsEnd < eod.getTime() ? tsEnd : eod.getTime()));
      cur = nextWDStart(cur);
    }
    return Math.max(0, tot);
  }

  // ── PAIR ROWS (replica exata do dashboard) ────────────────────
  function pairRows(rows, openType, closeType) {
    const pairs = [];
    const opens = [];
    rows.forEach(r => {
      if (r.tipo === openType) {
        opens.push({ ...r, _used: false });
      } else if (r.tipo === closeType) {
        let best = null, bestScore = -1;
        for (let i = 0; i < opens.length; i++) {
          const o = opens[i];
          if (o._used || o.func !== r.func || o.tsMs > r.tsMs) continue;
          let score = 1;
          if (o.op    === r.op)    score += 4;
          if (o.serie === r.serie) score += 2;
          if (o.item  === r.item)  score += 1;
          if (score > bestScore) { bestScore = score; best = i; }
        }
        if (best !== null) {
          const o = opens[best];
          const dur = workMinutes(o.tsMs, r.tsMs);
          if (dur > 0 && dur < 4800) {
            pairs.push({ func: r.func, op: r.op || o.op, serie: r.serie || o.serie,
              item: r.item || o.item, dur, openTs: o.tsMs, closeTs: r.tsMs });
          }
          opens[best]._used = true;
        }
      }
    });
    return pairs;
  }

  // ── CHECK 1: ABERTURAs sem FECHAMENTO ─────────────────────────
  const pairsAudit = pairRows(records, 'ABERTURA', 'FECHAMENTO');
  const pairedOpenKeys = new Set(pairsAudit.map(p => p.func + '||' + p.openTs));
  const backlog = records.filter(r =>
    r.tipo === 'ABERTURA' && !pairedOpenKeys.has(r.func + '||' + r.tsMs)
  );

  // ── CHECK 2: Durações suspeitas (> 10h work ou < 1 min work) ──
  const allPairs = [
    ...pairRows(records, 'ABERTURA', 'FECHAMENTO'),
    ...pairRows(records, 'INÍCIO DE PARADA', 'TÉRMINO DE PARADA'),
    ...pairRows(records, 'INÍCIO DE RETRABALHO', 'TÉRMINO DE RETRABALHO'),
  ];
  const tooLong  = allPairs.filter(p => p.dur > 600);
  const tooShort = allPairs.filter(p => p.dur > 0 && p.dur < 1);

  // ── CHECK 3: Apontamentos simultâneos (< 2 min entre ABERTURAs, ops distintas) ──
  const aberturas = records.filter(r => r.tipo === 'ABERTURA').sort((a, b) => a.tsMs - b.tsMs);
  const byFunc = {};
  aberturas.forEach(r => { if (!byFunc[r.func]) byFunc[r.func] = []; byFunc[r.func].push(r); });
  const simultaneos = [];
  Object.entries(byFunc).forEach(([func, recs]) => {
    for (let i = 0; i < recs.length - 1; i++) {
      const a = recs[i], b = recs[i+1];
      const diffMin = (b.tsMs - a.tsMs) / 60000;
      if (diffMin < 2 && a.op !== b.op) {
        simultaneos.push({ func, tsA: a.tsMs, tsB: b.tsMs, opA: a.op, opB: b.op });
      }
    }
  });

  // ── CHECK 4: FECHAMENTOs sem ABERTURA ────────────────────────
  const pairedCloseKeys = new Set(pairsAudit.map(p => p.func + '||' + p.closeTs));
  const orphanClose = records.filter(r =>
    r.tipo === 'FECHAMENTO' && !pairedCloseKeys.has(r.func + '||' + r.tsMs)
  );

  // ── TOTAIS ──────────────────────────────────────────────────
  const total = backlog.length + tooLong.length + tooShort.length + simultaneos.length + orphanClose.length;

  function fmtTs(ms) {
    return Utilities.formatDate(new Date(ms), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  }

  return {
    success: true,
    totalRegistros: records.length,
    totalPares: pairsAudit.length,
    totalInconsistencias: total,
    aberturaSemFechamento: {
      count: backlog.length, sev: 'alta',
      exemplos: backlog.slice(0, 15).map(r => ({
        func: r.func, op: r.op, serie: r.serie, ts: fmtTs(r.tsMs)
      }))
    },
    fechamentoSemAbertura: {
      count: orphanClose.length, sev: 'baixa',
      exemplos: orphanClose.slice(0, 15).map(r => ({
        func: r.func, op: r.op, ts: fmtTs(r.tsMs)
      }))
    },
    duracaoLonga: {
      count: tooLong.length, sev: 'media',
      exemplos: tooLong.slice(0, 10).map(p => ({
        func: p.func, op: p.op, durMin: Math.round(p.dur),
        de: fmtTs(p.openTs), ate: fmtTs(p.closeTs)
      }))
    },
    duracaoCurta: {
      count: tooShort.length, sev: 'baixa',
      exemplos: tooShort.slice(0, 10).map(p => ({
        func: p.func, op: p.op, durSeg: Math.round(p.dur * 60),
        ts: fmtTs(p.openTs)
      }))
    },
    simultaneos: {
      count: simultaneos.length, sev: 'media',
      exemplos: simultaneos.slice(0, 10).map(s => ({
        func: s.func, opA: s.opA, opB: s.opB,
        tsA: fmtTs(s.tsA), tsB: fmtTs(s.tsB)
      }))
    }
  };
}

// ── Placeholder functions referenciadas em doGet ──────────────────────────────
function analisarOrfaos() {
  return analisarInconsistencias();
}

function marcarOrfaosLegado(cutoff) {
  return { success: false, message: 'Função não implementada. Use analisarInconsistencias.' };
}

// ================================================================
// IMPACTO DA REMOÇÃO DE ÓRFÃOS — detalha os registros inconsistentes
// para avaliação antes de qualquer remoção
// Execute via: ?action=impactoOrfaos&key=AGF2026
// ================================================================
function impactoOrfaos() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return { success: false, message: 'Sem dados.' };

  function fmtTs(ms) {
    return Utilities.formatDate(new Date(ms), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  }

  // ── PARSE (igual a analisarInconsistencias) ──────────────────
  const records = [];
  for (let i = 1; i < dados.length; i++) {
    const row    = dados[i];
    const tsRaw  = row[0];
    const func   = String(row[1] || '').trim();
    const tipo   = String(row[2] || '').trim();
    const op     = String(row[3] || '').trim();
    const item   = String(row[4] || '').trim();
    const serie  = String(row[5] || '').trim().split(' | ')[0].trim();

    if (!func || !tipo) continue;

    let tsMs;
    if (tsRaw instanceof Date) tsMs = tsRaw.getTime();
    else if (tsRaw)            tsMs = new Date(String(tsRaw)).getTime();
    if (!tsMs || isNaN(tsMs))  continue;

    records.push({ tsMs, func, tipo, op, item, serie, sheetRow: i + 1 });
  }

  records.sort((a, b) => a.tsMs - b.tsMs);

  // ── PAIR ROWS ────────────────────────────────────────────────
  function pairRows(rows, openType, closeType) {
    const pairs = [];
    const opens = [];
    rows.forEach(r => {
      if (r.tipo === openType) {
        opens.push({ ...r, _used: false });
      } else if (r.tipo === closeType) {
        let best = null, bestScore = -1;
        for (let i = 0; i < opens.length; i++) {
          const o = opens[i];
          if (o._used || o.func !== r.func || o.tsMs > r.tsMs) continue;
          let score = 1;
          if (o.op    === r.op)    score += 4;
          if (o.serie === r.serie) score += 2;
          if (o.item  === r.item)  score += 1;
          if (score > bestScore) { bestScore = score; best = i; }
        }
        if (best !== null) {
          const o = opens[best];
          const durRaw = (r.tsMs - o.tsMs) / 60000; // minutos reais (sem filtro workMinutes)
          if (durRaw > 0) {
            pairs.push({ func: r.func, op: r.op || o.op, serie: r.serie || o.serie,
              item: r.item || o.item, durRaw, openTs: o.tsMs, closeTs: r.tsMs });
          }
          opens[best]._used = true;
        }
      }
    });
    return pairs;
  }

  const pairsAudit = pairRows(records, 'ABERTURA', 'FECHAMENTO');
  const pairedOpenKeys  = new Set(pairsAudit.map(p => p.func + '||' + p.openTs));
  const pairedCloseKeys = new Set(pairsAudit.map(p => p.func + '||' + p.closeTs));

  const backlog     = records.filter(r => r.tipo === 'ABERTURA'   && !pairedOpenKeys.has(r.func  + '||' + r.tsMs));
  const orphanClose = records.filter(r => r.tipo === 'FECHAMENTO' && !pairedCloseKeys.has(r.func + '||' + r.tsMs));

  // ── AGRUPA POR FUNCIONÁRIO ──────────────────────────────────
  function agrupar(lista) {
    const m = {};
    lista.forEach(r => {
      if (!m[r.func]) m[r.func] = { func: r.func, count: 0, ops: new Set(), datas: [] };
      m[r.func].count++;
      if (r.op) m[r.func].ops.add(r.op);
      m[r.func].datas.push(r.tsMs);
    });
    return Object.values(m).map(v => ({
      func: v.func.split(' - ').slice(1).join(' ') || v.func,
      funcRaw: v.func,
      count: v.count,
      ops: [...v.ops].join(', '),
      primeiro: fmtTs(Math.min(...v.datas)),
      ultimo:   fmtTs(Math.max(...v.datas))
    })).sort((a, b) => b.count - a.count);
  }

  // ── DISTRIBUIÇÃO POR MÊS ────────────────────────────────────
  function porMes(lista) {
    const m = {};
    lista.forEach(r => {
      const mes = Utilities.formatDate(new Date(r.tsMs), 'America/Sao_Paulo', 'yyyy-MM');
      m[mes] = (m[mes] || 0) + 1;
    });
    return Object.entries(m).sort().map(([mes, cnt]) => ({ mes, cnt }));
  }

  // ── ESTIMATIVA DE HORAS PERDIDAS ─────────────────────────────
  // Para ABERTURAs órfãs: tentamos estimar duração pelo próximo FECHAMENTO do mesmo func
  // (mesmo que não pareado), para ter ideia do volume perdido
  let minutosPerdidosEstimados = 0;
  backlog.forEach(ab => {
    // Procura o FECHAMENTO mais próximo após esta ABERTURA para o mesmo funcionário
    const prox = orphanClose.find(fc => fc.func === ab.func && fc.tsMs > ab.tsMs);
    if (prox) {
      const d = (prox.tsMs - ab.tsMs) / 60000;
      if (d > 0 && d < 24 * 60) minutosPerdidosEstimados += d; // só conta se < 24h
    }
  });

  return {
    success: true,
    totalRegistros: records.length,
    totalOrfaos: backlog.length + orphanClose.length,
    pctDoTotal: ((backlog.length + orphanClose.length) / records.length * 100).toFixed(2) + '%',

    aberturaSemFechamento: {
      total: backlog.length,
      porFuncionario: agrupar(backlog),
      porMes: porMes(backlog),
      detalhes: backlog.map(r => ({
        func: r.func, op: r.op, serie: r.serie, item: r.item,
        ts: fmtTs(r.tsMs), sheetRow: r.sheetRow
      }))
    },

    fechamentoSemAbertura: {
      total: orphanClose.length,
      porFuncionario: agrupar(orphanClose),
      porMes: porMes(orphanClose),
      detalhes: orphanClose.map(r => ({
        func: r.func, op: r.op, serie: r.serie, item: r.item,
        ts: fmtTs(r.tsMs), sheetRow: r.sheetRow
      }))
    },

    impactoEstimado: {
      horasEstimadas: (minutosPerdidosEstimados / 60).toFixed(1),
      obs: 'Estimativa de horas de produção representadas pelos pares órfãos (ABERTURA+FECHAMENTO sem par entre si)'
    }
  };
}

// ================================================================
// PREVIEW / DELETAR ABERTURAS ÓRFÃS HISTÓRICAS
//
// Identifica ABERTURAs sem FECHAMENTO correspondente com timestamp
// até a data de corte (cutoff), e opcionalmente as remove.
//
// Preview:  ?action=previewAberturasOrfas&key=AGF2026&cutoff=2026-05-29
// Deletar:  ?action=deletarAberturasOrfas&key=AGF2026&cutoff=2026-05-29
//
// A lógica de emparelhamento usa durReal (minutos reais de relógio),
// idêntica à correção aplicada no dashboard (não usa workMinutes para
// validar o par, evitando falsos positivos em fins de semana).
// ================================================================

function previewAberturasOrfas(cutoffDateStr) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const dados = aba.getDataRange().getValues();
  if (dados.length < 2) return { success: false, message: 'Sem dados.' };

  // Cutoff: fim do dia da data informada
  const cutoff   = cutoffDateStr ? new Date(cutoffDateStr + 'T23:59:59') : new Date('2026-05-29T23:59:59');
  const cutoffMs = cutoff.getTime();

  // ── PARSER DE DATAS — suporta Date objects, formato BR (dd/MM/yyyy HH:mm:ss)
  //    e ISO. Garante que strings no formato brasileiro (saídas de getDadosRespostas)
  //    sejam corretamente interpretadas.
  function parseTs(val) {
    if (!val && val !== 0) return null;
    if (val instanceof Date) {
      const ms = val.getTime();
      return (!ms || isNaN(ms)) ? null : ms;
    }
    const s = String(val).trim();
    if (!s) return null;
    // Formato BR: dd/MM/yyyy HH:mm:ss  ou  dd/MM/yyyy, HH:mm:ss
    const brMatch = s.match(/^(\d{2})\/(\d{2})\/(\d{4})[,\s]+(\d{2}):(\d{2}):(\d{2})/);
    if (brMatch) {
      const [, dd, mm, yyyy, hh, min, sec] = brMatch;
      const ms = new Date(+yyyy, +mm - 1, +dd, +hh, +min, +sec).getTime();
      return isNaN(ms) ? null : ms;
    }
    // Fallback: ISO ou qualquer outro formato reconhecível
    const ms = new Date(s).getTime();
    return isNaN(ms) ? null : ms;
  }

  // ── PARSE — inclui todas as linhas com timestamp válido ─────────
  // Usa Utilities.formatDate para gerar a chave de timestamp no mesmo
  // formato que getDadosRespostas → alinhado com o algoritmo do dashboard.
  const records = [];
  for (let i = 1; i < dados.length; i++) {
    const row   = dados[i];
    const tsRaw = row[0];
    if (!tsRaw || (tsRaw instanceof Date && isNaN(tsRaw.getTime()))) continue;

    // Gera string canônica: "dd/MM/yyyy HH:mm:ss" em GMT-3 (igual a getData)
    let tsStr;
    if (tsRaw instanceof Date) {
      try { tsStr = Utilities.formatDate(tsRaw, 'GMT-3', 'dd/MM/yyyy HH:mm:ss'); }
      catch(e) { continue; }
    } else {
      tsStr = String(tsRaw).trim();
    }
    if (!tsStr) continue;

    // Converte para ms para comparação de datas (cutoff, ordenação)
    const tsMs = parseTs(tsRaw) || parseTs(tsStr);
    if (!tsMs) continue;

    const func  = String(row[1]  || '').trim();
    const tipo  = String(row[2]  || '').trim();
    // Coluna D (index 3) = TIPO DE OPERAÇÃO 1  |  Coluna L (index 11) = TIPO DE OPERAÇÃO 2
    // Dashboard usa: op = D || L  (mesmo fallback)
    const op    = (String(row[3]  || '').trim()) || (String(row[11] || '').trim());
    const item  = String(row[4]  || '').trim();
    const serie = String(row[5]  || '').trim().split(' | ')[0].trim();

    records.push({ sheetRow: i + 1, tsMs, tsStr, func, tipo, op, item, serie });
  }

  // Ordena cronologicamente (igual ao dashboard)
  records.sort((a, b) => a.tsMs - b.tsMs);

  // ── PAIR ROWS — réplica exata do dashboard ────────────────────────
  // Chave de identificação: func + '||' + tsStr  (mesmo formato da dashboard)
  function pairRows(rows, openType, closeType) {
    const opens = [];
    const pairedOpenKeys = new Set();  // chaves: func||tsStr
    rows.forEach(r => {
      if (r.tipo === openType) {
        opens.push({ ...r, _used: false });
      } else if (r.tipo === closeType) {
        let best = null, bestScore = -1;
        for (let i = 0; i < opens.length; i++) {
          const o = opens[i];
          if (o._used || o.func !== r.func || o.tsMs > r.tsMs) continue;
          let score = 1;
          if (o.op    === r.op)    score += 4;
          if (o.serie === r.serie) score += 2;
          if (o.item  === r.item)  score += 1;
          if (score > bestScore) { bestScore = score; best = i; }
        }
        if (best !== null) {
          const o       = opens[best];
          const durReal = (r.tsMs - o.tsMs) / 60000;  // minutos reais de relógio
          if (durReal > 0 && durReal < 10080) {         // < 7 dias reais
            pairedOpenKeys.add(o.func + '||' + o.tsStr);
          }
          opens[best]._used = true;
        }
      }
    });
    return pairedOpenKeys;
  }

  const pairedOpenKeys = pairRows(records, 'ABERTURA', 'FECHAMENTO');

  // ── SELECIONA ÓRFÃS ATÉ O CUTOFF ─────────────────────────────────
  const orfas = records.filter(r =>
    r.tipo  === 'ABERTURA' &&
    r.func  !== ''          &&   // descarta linhas sem operador
    r.tsMs  <= cutoffMs     &&
    !pairedOpenKeys.has(r.func + '||' + r.tsStr)   // chave alinhada com dashboard
  );

  const fmtTs = ms => new Date(ms).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });

  return {
    success:          true,
    cutoff:           cutoff.toLocaleDateString('pt-BR'),
    totalRegistros:   records.length,
    totalPares:       pairedOpenKeys.size,
    orfasEncontradas: orfas.length,
    items: orfas.map(r => ({
      sheetRow: r.sheetRow,
      ts:       r.tsStr,
      func:     r.func,
      op:       r.op,
      item:     r.item,
      serie:    r.serie
    }))
  };
}

function deletarAberturasOrfas(cutoffDateStr) {
  // 1. Identificar as linhas a deletar via preview
  const preview = previewAberturasOrfas(cutoffDateStr);
  if (!preview.success) return preview;
  if (preview.orfasEncontradas === 0) {
    return { success: true, removidos: 0, message: 'Nenhuma ABERTURA órfã encontrada até ' + (cutoffDateStr || '29/05/2026') + '.' };
  }

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  // 2. Ordena de baixo para cima para não deslocar índices durante a deleção
  const linhas = preview.items.map(r => r.sheetRow).sort((a, b) => b - a);

  for (const linha of linhas) {
    aba.deleteRow(linha);
  }
  SpreadsheetApp.flush();

  // 3. Reconstrói a aba Abertos para refletir o estado atual
  let recResult = null;
  try { recResult = reconstruirAbertos(); } catch(e) { recResult = { error: e.message }; }

  return {
    success:         true,
    removidos:       linhas.length,
    linhasDeletadas: [...linhas].reverse(), // ordem crescente p/ leitura
    reconstruirAbertos: recResult,
    message:         linhas.length + ' ABERTURA(s) órfã(s) removida(s). Aba Abertos reconstruída.'
  };
}

// ================================================================
// LOCALIZAR LINHAS POR IDENTIFICADORES (func + ts)
//
// Recebe uma lista de {func, ts} vindos do algoritmo do dashboard
// (executado no browser via getData) e localiza os números de linha
// exatos na planilha Respostas, para deleção segura.
//
// Chamada via POST com payload:
//   { action: 'localizarLinhas', key: 'AGF2026',
//     identifiers: [{func, ts}, ...], deletar: false }
//
// ts deve estar no formato 'dd/MM/yyyy HH:mm:ss' (sem vírgula),
// igual ao produzido por getDadosRespostas / Utilities.formatDate.
// ================================================================

// ================================================================
// Fix#ID-LINK — MIGRAÇÃO HISTÓRICA DE IDs
//
// Recebe lista de {sheetRow, abertoId} e escreve o abertoId na
// coluna P (col 16) de cada linha especificada da aba Respostas.
//
// Usado APENAS UMA VEZ para linkar registros históricos.
// Operação idempotente: se col P já tem o mesmo ID, não faz nada.
// ================================================================
function migrarIdsHistoricos(updates) {
  if (!updates || updates.length === 0) {
    return { success: false, message: 'Lista de updates vazia.' };
  }
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Aba sem dados.' };

  // Leitura em bulk: toda a coluna P de uma vez (muito mais rápido que célula a célula)
  const numDataRows = lastRow - 1; // linha 1 = cabeçalho
  const colPRange   = aba.getRange(2, 16, numDataRows, 1);
  const colPValues  = colPRange.getValues(); // [[val], [val], ...]

  let atualizados = 0, ignorados = 0, erros = [], dirty = false;

  for (const upd of updates) {
    const row = parseInt(upd.sheetRow);
    const id  = String(upd.abertoId || '').trim();
    if (!row || row < 2 || row > lastRow || !id) {
      erros.push('Linha inválida ou ID vazio: ' + JSON.stringify(upd));
      continue;
    }
    const idx   = row - 2; // 0-indexed no array
    const atual = String(colPValues[idx][0] || '').trim();
    if (atual === id) { ignorados++; continue; } // idempotente: já correto
    colPValues[idx][0] = id;
    atualizados++;
    dirty = true;
  }

  // Escrita em bulk: uma única chamada para toda a coluna
  if (dirty) {
    colPRange.setValues(colPValues);
    SpreadsheetApp.flush();
  }

  return {
    success: true,
    total: updates.length,
    atualizados,
    ignorados,
    erros: erros.length > 0 ? erros : undefined,
    message: atualizados + ' linha(s) atualizadas, ' + ignorados + ' já corretas.',
  };
}

// ================================================================
// Deletar linhas por número de linha da planilha (col A = Carimbo de data/hora).
// Verifica que cada linha é do tipo ABERTURA antes de remover.
// dryRun=true → apenas lista o que seria deletado, sem apagar.
// ================================================================
function deletarLinhasPorNumero(rows, dryRun, tiposPermitidos) {
  // tiposPermitidos: array de tipos aceitos para deleção (default: ['ABERTURA'])
  const tipos = Array.isArray(tiposPermitidos) && tiposPermitidos.length > 0
    ? tiposPermitidos
    : ['ABERTURA'];
  if (!rows || rows.length === 0) {
    return { success: false, message: 'Lista de linhas vazia.' };
  }
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  const lastRow = aba.getLastRow();

  // Ler todos os dados uma vez (bulk)
  const dados = aba.getRange(1, 1, lastRow, 20).getValues();

  const confirmadas = [];  // linhas válidas (ABERTURA confirmada)
  const rejeitadas  = [];  // linhas que não são ABERTURA ou fora do range

  for (const rowNum of rows) {
    const r = parseInt(rowNum);
    if (!r || r < 2 || r > lastRow) {
      rejeitadas.push({ row: r, motivo: 'fora do range' });
      continue;
    }
    const tipo = String(dados[r - 1][2] || '').trim(); // col C = tipo
    const func = String(dados[r - 1][1] || '').trim(); // col B = operador
    const ts   = dados[r - 1][0];                      // col A = timestamp
    if (!tipos.includes(tipo)) {
      rejeitadas.push({ row: r, tipo, func, motivo: 'tipo não permitido (permitidos: ' + tipos.join(',') + ')' });
      continue;
    }
    let tsStr;
    try { tsStr = ts instanceof Date ? Utilities.formatDate(ts, 'GMT-3', 'dd/MM/yyyy HH:mm:ss') : String(ts).trim(); }
    catch(e) { tsStr = String(ts); }
    confirmadas.push({ row: r, ts: tsStr, func });
  }

  if (dryRun) {
    return {
      success: true,
      dryRun: true,
      totalSolicitado: rows.length,
      confirmadas: confirmadas.length,
      rejeitadas: rejeitadas.length,
      detalheRejeitadas: rejeitadas,
      amostraConfirmadas: confirmadas.slice(0, 10),
      message: 'DRY-RUN: ' + confirmadas.length + ' linha(s) seriam deletadas, ' + rejeitadas.length + ' rejeitadas.'
    };
  }

  // Deletar de baixo para cima para não deslocar índices
  const linhasOrdenadas = confirmadas.map(r => r.row).sort((a, b) => b - a);
  for (const linha of linhasOrdenadas) {
    aba.deleteRow(linha);
  }
  SpreadsheetApp.flush();

  // Reconstruir aba de abertos
  let recResult = null;
  try { recResult = reconstruirAbertos(); } catch(e) { recResult = { error: e.message }; }

  return {
    success: true,
    totalSolicitado: rows.length,
    deletados: linhasOrdenadas.length,
    rejeitadas: rejeitadas.length,
    detalheRejeitadas: rejeitadas.length > 0 ? rejeitadas : undefined,
    reconstruirAbertos: recResult,
    message: linhasOrdenadas.length + ' linha(s) deletada(s). ' + rejeitadas.length + ' rejeitada(s) (tipo != ABERTURA ou fora do range).'
  };
}

function localizarLinhasPorIdentificadores(identifiers, deletar) {
  if (!identifiers || identifiers.length === 0) {
    return { success: false, message: 'Lista de identificadores vazia.' };
  }

  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = ss.getSheetByName(ABA_RESPOSTAS);
  if (!aba) return { success: false, message: 'Aba Respostas não encontrada.' };

  // Constrói Set de chaves: "func|||ts"
  const keysAlvo = new Set(identifiers.map(id => id.func + '|||' + id.ts));

  const dados = aba.getDataRange().getValues();
  const linhasEncontradas = [];
  const naoEncontrados = [...keysAlvo]; // debug: quais não foram achados

  for (let i = 1; i < dados.length; i++) {
    const row    = dados[i];
    const tsRaw  = row[0];
    const func   = String(row[1] || '').trim();
    const tipo   = String(row[2] || '').trim();
    if (tipo !== 'ABERTURA') continue; // só ABERTURAs

    let tsStr;
    if (tsRaw instanceof Date) {
      try { tsStr = Utilities.formatDate(tsRaw, 'GMT-3', 'dd/MM/yyyy HH:mm:ss'); }
      catch(e) { continue; }
    } else if (tsRaw) {
      tsStr = String(tsRaw).trim();
    } else {
      continue;
    }

    const chave = func + '|||' + tsStr;
    if (keysAlvo.has(chave)) {
      const op    = String(row[3]  || '').trim() || String(row[11] || '').trim();
      const item  = String(row[4]  || '').trim();
      const serie = String(row[5]  || '').trim().split(' | ')[0].trim();
      linhasEncontradas.push({ sheetRow: i + 1, ts: tsStr, func, op, item, serie });
      // Remove do conjunto de não-encontrados
      const idx = naoEncontrados.indexOf(chave);
      if (idx >= 0) naoEncontrados.splice(idx, 1);
    }
  }

  if (!deletar) {
    return {
      success:        true,
      buscados:       keysAlvo.size,
      encontrados:    linhasEncontradas.length,
      naoEncontrados: naoEncontrados.length,
      items:          linhasEncontradas
    };
  }

  // ── DELETAR ─────────────────────────────────────────────────────
  if (linhasEncontradas.length === 0) {
    return { success: true, removidos: 0, message: 'Nenhuma linha localizada para deleção.' };
  }

  const linhas = linhasEncontradas.map(r => r.sheetRow).sort((a, b) => b - a); // desc
  for (const linha of linhas) {
    aba.deleteRow(linha);
  }
  SpreadsheetApp.flush();

  let recResult = null;
  try { recResult = reconstruirAbertos(); } catch(e) { recResult = { error: e.message }; }

  return {
    success:         true,
    buscados:        keysAlvo.size,
    removidos:       linhas.length,
    naoEncontrados:  naoEncontrados.length,
    linhasDeletadas: [...linhas].reverse(),
    reconstruirAbertos: recResult,
    message:         linhas.length + ' linha(s) deletada(s). ' + naoEncontrados.length + ' identificador(es) não localizado(s).'
  };
}

