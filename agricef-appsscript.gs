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

const SPREADSHEET_ID = '15vtJ2eOw3Zd9f5MmwqEj18nsGAvVkFYFpsUsRbZM6Ik';
const ABA_RESPOSTAS  = 'Respostas do Formulário 1';
const ABA_ABERTOS    = 'Abertos';
const ABA_OPERADORES = 'Cadastro_Operadores';
const ABA_SALDO      = 'Saldo_Parcial';
const ABA_SERIES     = 'Cadastro_Series';

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
  if (action === 'getData')          return getDadosRespostas();
  if (action === 'triggerRelatorio' && e.parameter.key === 'AGF2026') {
    try {
      enviarRelatorioSemanal();
      return jsonResponse({ success: true, message: 'Relatório enviado para ' + EMAIL_RELATORIO });
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
  if (action === 'normalizarOperadores' && e.parameter.key === 'AGF2026') {
    try {
      const total = normalizarCodigosOperador();
      return jsonResponse({ success: true, message: total + ' código(s) normalizado(s) com sucesso.' });
    } catch(err) {
      return jsonResponse({ success: false, message: err.message });
    }
  }

  // Ações via payload GET (contorna CORS)
  if (e.parameter.payload) {
    try {
      const payload = JSON.parse(e.parameter.payload);
      if (payload.action === 'salvarOperador')  return salvarOperador(payload);
      if (payload.action === 'removerOperador') return removerOperador(payload);
      if (payload.action === 'salvarSerie')     return salvarSerie(payload);
      if (payload.action === 'removerSerie')    return removerSerie(payload);
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
    if (payload.action === 'salvarOperador')  return salvarOperador(payload);
    if (payload.action === 'removerOperador') return removerOperador(payload);
    if (payload.action === 'salvarSerie')     return salvarSerie(payload);
    if (payload.action === 'removerSerie')    return removerSerie(payload);
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
        let loteSeries = null;
        if (row[11]) {
          try { loteSeries = JSON.parse(String(row[11])); } catch(e) {}
        }
        return jsonResponse({
          aberto:        true,
          tipo:          row[2]  || '',
          operacao:      row[3]  || '',
          carimbo:       formatarCarimboGs(row[4]),
          codItem:       row[5]  || '',
          qtdPlanejada:  row[6]  || '',
          nrSerie:       row[7]  || '',
          implemento:    row[8]  || '',
          cliente:       row[9]  || '',
          operadorNome:  row[10] || '',
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
          return jsonResponse({
            success: false,
            bloqueado: true,
            message: 'Operador já possui apontamento em aberto. Feche-o antes de iniciar um novo.',
            aberto: {
              tipo:         dadosAbertos[i][2] || '',
              operacao:     dadosAbertos[i][3] || '',
              carimbo:      formatarCarimboGs(dadosAbertos[i][4]),
              codItem:      dadosAbertos[i][5] || '',
              qtdPlanejada: dadosAbertos[i][6] || '',
              nrSerie:      dadosAbertos[i][7] || '',
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
    // ---------------------------------------------------------------
    if (tiposFechamento.includes(tipo)) {
      const abertoIdPayload = String(payload.abertoId || '').trim();
      for (let i = 1; i < dadosAbertos.length; i++) {
        if (!dadosAbertos[i][0]) continue;
        const rowId   = String(dadosAbertos[i][12] || '').trim();
        const matchId = abertoIdPayload && rowId && rowId === abertoIdPayload;
        const matchOp = !matchId && mesmoOperador(dadosAbertos[i][0], payload.operador);
        if (matchId || matchOp) {
          const tipoAberto     = String(dadosAbertos[i][2] || '').trim();
          const tipoEsperado   = TIPO_COMPATIVEL[tipo];
          if (tipoEsperado && tipoAberto !== tipoEsperado) {
            return jsonResponse({
              success: false,
              message: 'Tipo de fechamento incompatível. Você possui "' + tipoAberto +
                       '" em aberto, mas tentou registrar "' + (TIPOS_APONTAMENTO[tipo] || tipo) + '".',
              incompativel: true,
              tipoAberto: tipoAberto,
            });
          }
          break; // encontrou o registro — compatível → pode continuar
        }
      }
    }

    // Carimbo: vem do browser já formatado (dd/MM/yyyy HH:mm:ss no fuso local)
    // Fallback para o servidor em GMT-3 caso não venha
    const carimbo = (payload.timestamp && !payload.timestamp.includes('T'))
      ? payload.timestamp
      : Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss');

    // ID único para o registro na aba Abertos (gerado uma vez por abertura)
    // Fechamentos recebem o ID do frontend via payload.abertoId
    const abertoId = tiposAbertura.includes(tipo) ? gerarIdApontamento() : null;

    const nomeOperador  = payload.operadorNome || payload.operador || '';
    const tipoFormatado = TIPOS_APONTAMENTO[tipo] || tipo;
    const op1           = OPERACOES[payload.operacao]     || payload.operacao     || '';
    const op2           = OPERACOES[payload.opRetrabalho] || payload.opRetrabalho || '';
    const codItem       = payload.codItem || '';

    const qtd = (payload.quantidade === null || payload.quantidade === undefined || payload.quantidade === '')
      ? '' : Number(payload.quantidade);

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
      gerarIdApontamento(),          // P — ID único
    ];

    if (payload.loteSeries && Array.isArray(payload.loteSeries) && payload.loteSeries.length > 0) {
      // Batch write — uma única chamada de API para todas as séries do lote
      const rows = payload.loteSeries.map(item => {
        const linhaMod = [...linha];
        linhaMod[5]  = item.nrSerie + ' | ' + item.implemento + ' | ' + item.cliente;
        linhaMod[15] = gerarIdApontamento();
        return linhaMod;
      });
      const primeiraLinha = abaRe.getLastRow() + 1;
      abaRe.getRange(primeiraLinha, 1, rows.length, rows[0].length).setValues(rows);
    } else if (payload.lote && payload.lote.trim() !== '') {
      // Formato legado — batch write também
      const series = payload.lote.split(',').map(s => s.trim()).filter(Boolean);
      const rows = series.map(serie => {
        const linhaMod = [...linha];
        linhaMod[5]  = serie + ' | ' + payload.implemento + ' | ' + payload.cliente;
        linhaMod[15] = gerarIdApontamento();
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

function getCadastros() {
  try {
    // Cache de 5 min — evita abrir a planilha a cada carregamento de página
    const cache  = CacheService.getScriptCache();
    const cached = cache.get('cadastros_v2');
    if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const abaOp = garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo', 'Nome', 'Ativo']);
    const abaSe = garantirAbaCadastro(ss, ABA_SERIES,     ['NrSerie', 'Implemento', 'Cliente', 'Ativo']);

    const operadores = abaOp.getDataRange().getValues().slice(1)
      .filter(r => String(r[2]).toUpperCase() !== 'NÃO' && r[0] !== '')
      .map(r => ({ codigo: normalizarCodigoOp(String(r[0]).trim()), nome: String(r[1]).trim() }));

    const series = abaSe.getDataRange().getValues().slice(1)
      .filter(r => String(r[3]).toUpperCase() !== 'NÃO' && r[0] !== '')
      .map(r => ({ nrSerie: String(r[0]).trim(), implemento: String(r[1]).trim(), cliente: String(r[2]).trim() }));

    const resultado = JSON.stringify({ success: true, operadores, series });
    cache.put('cadastros_v2', resultado, 300); // 5 minutos
    return ContentService.createTextOutput(resultado).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function invalidarCacheCadastros() {
  try { CacheService.getScriptCache().remove('cadastros_v2'); } catch(e) {}
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
    // Garante que o código sempre é salvo com 6 dígitos (ex: "000130")
    const codigoNorm = normalizarCodigoOp(String(payload.codigo || '').trim());
    // Força formato texto na coluna A para preservar zeros à esquerda
    aba.getRange('A:A').setNumberFormat('@');
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (mesmoOperador(dados[i][0], codigoNorm)) {
        aba.getRange(i+1, 1).setValue(codigoNorm); // corrige o código se estava sem zeros
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

// Normaliza código de operador para sempre 6 dígitos com zeros à esquerda.
// Se o valor da célula foi convertido pelo Sheets para número (ex: 130),
// reconstitui o formato canônico "000130".
function normalizarCodigoOp(val) {
  const s = String(val === null || val === undefined ? '' : val).trim();
  if (!s) return s;
  const n = Number(s);
  // só normaliza se for número puro positivo (sem letras, pontada ou traço)
  if (!isNaN(n) && n > 0 && String(n) === s) return String(n).padStart(6, '0');
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
// NORMALIZAÇÃO DE CÓDIGOS DE OPERADOR — Execute UMA VEZ
// Corrige registros existentes em Abertos e Cadastro_Operadores
// para garantir que todos os códigos têm 6 dígitos (ex: "000130").
// Isso resolve divergências de backlog causadas pelo formato
// "130" vs "000130" nas diferentes fontes de dados.
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

  // 3. Invalida cache de cadastros para refletir as correções
  invalidarCacheCadastros();

  Logger.log('✅ Normalização concluída. ' + totalCorrigidos + ' código(s) corrigido(s).');
  return totalCorrigidos;
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
    linhas.push({
      ts,
      tipo,
      func:   String(r[1] || '').trim(),
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
        if (a._used || a.func !== fech.func || a.ts > fech.ts) continue;
        let score = 1;
        if (a.op    === fech.op)    score += 4;
        if (a.serie === fech.serie) score += 2;
        if (a.item  === fech.item)  score += 1;
        if (score > melhorScore) { melhorScore = score; melhor = i; }
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

  // ----- Top operadores (semana atual) -----
  const contOp = {};
  pares.filter(p => p._fechTs >= seg).forEach(p => {
    const nome = p.func.includes(' - ') ? p.func.split(' - ').slice(1).join(' ') : p.func;
    contOp[nome] = (contOp[nome] || 0) + 1;
  });
  const topOperadores = Object.entries(contOp).sort((a, b) => b[1] - a[1]).slice(0, 5);

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
      var nome       = o.func.includes(' - ') ? o.func.split(' - ').slice(1).join(' ') : o.func;
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
    linhas.push({
      ts,
      tipo,
      func:  String(r[1] || '').trim(),
      op:    String(r[3] || '').trim(),
      item:  String(r[4] || '').trim(),
      serie: pts[0] || '',
      impl:  pts[1] || '',
      qty:   Number(r[6])  || 0,
    });
  }

  // --- Pareamento ABERTURA ↔ FECHAMENTO (two-pass greedy) ---
  var aberturas = linhas
    .filter(function(r){ return r.tipo === 'ABERTURA'; })
    .map(function(r){ return Object.assign({}, r, { _used: false, _fechTs: null, _leadMs: 0 }); });

  linhas
    .filter(function(r){ return r.tipo === 'FECHAMENTO'; })
    .forEach(function(fech) {
      var melhor = null, melhorScore = -1;
      for (var i = 0; i < aberturas.length; i++) {
        var a = aberturas[i];
        if (a._used || a.func !== fech.func || a.ts > fech.ts) continue;
        var score = 1;
        if (a.op    === fech.op)    score += 4;
        if (a.serie === fech.serie) score += 2;
        if (a.item  === fech.item)  score += 1;
        if (score > melhorScore) { melhorScore = score; melhor = i; }
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

  // --- Backlog por operador (quem conversar) ---
  var backlogPorOp = {};
  backlog.forEach(function(o){
    var nome = o.func.includes(' - ') ? o.func.split(' - ').slice(1).join(' ') : o.func;
    if (!backlogPorOp[nome]) backlogPorOp[nome] = { count: 0, ordens: [], nome: nome };
    backlogPorOp[nome].count++;
    backlogPorOp[nome].ordens.push(o);
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
  var dois_dias_atras = new Date(hoje); dois_dias_atras.setDate(hoje.getDate() - 2);
  var ativosRecentes = new Set(
    linhas.filter(function(r){ return r.ts >= dois_dias_atras; })
          .map(function(r){ return r.func; })
  );
  var todosOps = new Set(linhas.map(function(r){ return r.func; }));
  var semAtividade = [...todosOps].filter(function(f){ return !ativosRecentes.has(f); })
    .map(function(f){ return f.includes(' - ') ? f.split(' - ').slice(1).join(' ') : f; })
    .sort();

  // --- Retrabalhos ativos ---
  var retrabs = linhas.filter(function(r){ return r.tipo === 'RETRAB_INI'; });
  var retrAbertos = retrabs.filter(function(ri){
    return !linhas.some(function(rf){
      return rf.tipo === 'RETRAB_FIM' && rf.func === ri.func && rf.ts > ri.ts;
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
    var opRetrabs = [...new Set(retrabs.map(function(r){ return r.func.includes(' - ') ? r.func.split(' - ').slice(1).join(' ') : r.func; }))];
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
      var nome     = o.func.includes(' - ') ? o.func.split(' - ').slice(1).join(' ') : o.func;
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

