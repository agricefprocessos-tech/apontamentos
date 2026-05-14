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
  const operador   = payload.operador  || '';
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
      .map(r => ({ codigo: String(r[0]).trim(), nome: String(r[1]).trim() }));

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

function salvarOperador(payload) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = garantirAbaCadastro(ss, ABA_OPERADORES, ['Codigo', 'Nome', 'Ativo']);
    const dados = aba.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(payload.codigo).trim()) {
        aba.getRange(i+1,2).setValue(payload.nome);
        aba.getRange(i+1,3).setValue('Sim');
        return jsonResponse({ success: true, message: 'Operador atualizado.' });
      }
    }
    aba.appendRow([payload.codigo, payload.nome, 'Sim']);
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

