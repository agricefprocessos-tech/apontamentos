// ================================================================
// AGRICEF — FormPCP.gs
// Módulo 2: Inclusão de Projetos PCP — Hauler (PPH1 + PPH2)
//
// ⚠️  PRÉ-REQUISITO: Desativar Regras Jira 07 (PPH1), 08 (PPH2),
//     09 (PPH2) antes de usar este módulo. A criação de tarefas
//     pai e subtarefas é feita 100% via API por este script.
//     Manter ativa a regra "casca" (calcula Nivel de Prioridade).
//
// Credenciais: Script Properties → JIRA_EMAIL + JIRA_TOKEN
// ================================================================

const JIRA_BASE    = 'https://agricefprojetos.atlassian.net';
const JIRA_PROJECT = 'AGTK';

// ─── SERVE O FORMULÁRIO ──────────────────────────────────────

function servirFormularioPCP_() {
  return HtmlService.createHtmlOutputFromFile('FormularioPCP')
    .setTitle('Novo Projeto PCP — Hauler | Agricef')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── FUNÇÃO PRINCIPAL — chamada via google.script.run ─────────
// dados = {
//   diametro: "8" | "10",
//   destino: "cliente" | "estoque" | "interno",
//   clienteNome: "ATVOS",
//   clienteNumero: "4",         // ex: "ATVOS 4"
//   serial: "S22000086",
//   startDate: "2026-05-21",    // Data início F1
//   diaC: "2026-07-15",         // DIA C: chegada do caminhão (F2)
//   criarFases: "ambas" | "f1" | "f2",
//   alongamento: false,
//   extensao: false,
//   fabricacaoCarroceria: "externa" | "interna" | "agricef",
//   classificacao: "Receita direta",
//   impactoFinanceiro: "Estratégico",
//   alinhamento: "Core Business",
//   urgencia: "Alta (Proximo Trimestre)",
//   complexidade: "Simples",
//   baselineDate: "2026-09-01",
//   alvoDate: "2026-08-15",
// }

function criarHaulerJira(dados) {
  try {
    const resultado = { chaves: {} };

    if (dados.criarFases !== 'f2') {
      // ── PPH1 ─────────────────────────────────────────────
      const datasF1    = calcularDatasF1_(dados.startDate);
      const summaryF1  = buildSummary_('F1', dados);
      const dueDateF1  = dados.baselineDate || datasF1.preMontagem.end;

      const bodyF1 = buildBodyPai_('PPH1', summaryF1, dados, dados.startDate, dueDateF1);
      const resF1  = jiraRequest_('POST', '/rest/api/3/issue', bodyF1);
      if (!resF1.key) throw new Error('Erro ao criar PPH1: ' + JSON.stringify(resF1));

      resultado.chaves.f1 = resF1.key;
      resultado.chaves.f1_url = JIRA_BASE + '/browse/' + resF1.key;

      // Subtarefas PPH1 (5 fixas)
      const subtasksF1 = [
        { nome: 'Preparação e Plan.', ...datasF1.prep        },
        { nome: 'Compras',           ...datasF1.compras      },
        { nome: 'Fabricação (Osti)', ...datasF1.fabOsti      },
        { nome: 'Fabricação (Agricef)', ...datasF1.fabAgricef },
        { nome: 'Pré – Montagem',   ...datasF1.preMontagem  },
      ];

      resultado.chaves.f1_subtasks = [];
      for (const st of subtasksF1) {
        const r = criarSubtarefa_(resF1.key, st.nome, st.start, st.end);
        if (r.key) resultado.chaves.f1_subtasks.push(r.key);
      }
    }

    if (dados.criarFases !== 'f1') {
      // ── PPH2 ─────────────────────────────────────────────
      const diaC       = dados.diaC || dados.startDate;
      const stF2       = calcularSubtasksF2_(diaC, dados);
      const summaryF2  = buildSummary_('F2', dados);
      const dueDateF2  = stF2[stF2.length - 1].end;

      const bodyF2 = buildBodyPai_('PPH2', summaryF2, dados, diaC, dueDateF2);
      const resF2  = jiraRequest_('POST', '/rest/api/3/issue', bodyF2);
      if (!resF2.key) throw new Error('Erro ao criar PPH2: ' + JSON.stringify(resF2));

      resultado.chaves.f2 = resF2.key;
      resultado.chaves.f2_url = JIRA_BASE + '/browse/' + resF2.key;

      resultado.chaves.f2_subtasks = [];
      for (const st of stF2) {
        const r = criarSubtarefa_(resF2.key, st.nome, st.start, st.end);
        if (r.key) resultado.chaves.f2_subtasks.push(r.key);
      }
    }

    return { success: true, resultado };

  } catch (err) {
    return { success: false, erro: err.message };
  }
}

// ─── BUSCAR PRÓXIMO SERIAL (chamado via google.script.run) ────

function buscarProximoSerial() {
  try {
    const jql = 'project=' + JIRA_PROJECT +
                ' AND issuetype=Tarefa AND Departamento=PCP ORDER BY created DESC';
    const res = jiraRequest_('GET',
      '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) +
      '&maxResults=100&fields=summary,customfield_10537');

    let maxNum = 85; // último serial conhecido: S22000085 (GERDAU)

    if (res.issues) {
      for (const issue of res.issues) {
        // Tenta o campo customfield_10537 primeiro
        const v537 = String(issue.fields.customfield_10537 || '');
        const m537 = v537.match(/S?2?2?0*(\d+)/i);
        if (m537) {
          const n = parseInt(m537[1]);
          if (n > maxNum) maxNum = n;
        }
        // Fallback: extrai do summary
        const summary = issue.fields.summary || '';
        const ms = summary.match(/S22000(\d+)/i);
        if (ms) {
          const n = parseInt(ms[1]);
          if (n > maxNum) maxNum = n;
        }
      }
    }

    const next = maxNum + 1;
    return 'S22000' + String(next).padStart(3, '0');

  } catch (err) {
    return 'S22000086'; // fallback seguro
  }
}

// ─── BUILDERS ────────────────────────────────────────────────

function buildSummary_(fase, dados) {
  const dim    = dados.diametro + '"';
  const serial = dados.serial || 'S22000XXX';
  let   dest;

  if (dados.destino === 'estoque') {
    dest = '(Estoque - ' + serial + ')';
  } else if (dados.destino === 'interno') {
    dest = '(' + (dados.clienteNome || 'Agricef') + ')';
  } else {
    const num = dados.clienteNumero ? ' ' + dados.clienteNumero : '';
    dest = '(' + serial + ') ' + (dados.clienteNome || '') + num;
  }

  return 'P157 - CAMINHAO DE TUBOS HAULER ' + dim + ' - ' + dest + ' - ' + fase;
}

function buildBodyPai_(categoria, summary, dados, startDate, dueDate) {
  const baseline = dados.baselineDate || dueDate;
  const alvo     = dados.alvoDate     || dueDate;

  const fields = {
    project:           { key: JIRA_PROJECT },
    issuetype:         { name: 'Tarefa' },
    summary:           summary,
    customfield_10205: { value: categoria },
    customfield_10073: 'PCP',
    customfield_10038: dados.clienteNome || '',
    customfield_10139: { value: '157' },
    customfield_10015: startDate,
    duedate:           dueDate,
    customfield_10469: baseline,
    customfield_10470: alvo,
    customfield_10271: { value: dados.classificacao },
    customfield_10304: { value: dados.impactoFinanceiro },
    customfield_10337: { value: dados.alinhamento },
    customfield_10370: { value: dados.urgencia },
    customfield_10403: { value: dados.complexidade },
  };

  if (dados.serial) {
    fields.customfield_10537 = dados.serial;
  }

  return { fields };
}

function criarSubtarefa_(parentKey, nome, startDate, endDate) {
  return jiraRequest_('POST', '/rest/api/3/issue', {
    fields: {
      project:          { key: JIRA_PROJECT },
      issuetype:        { name: 'Subtarefa' },
      parent:           { key: parentKey },
      summary:          nome,
      customfield_10015: startDate,
      duedate:          endDate,
    }
  });
}

// ─── CÁLCULO DE DATAS ────────────────────────────────────────
// Estimativas conservadoras baseadas em dados reais dos 42 projetos.
// Parallelismo: Compras, Fab.Osti e Fab.Agricef iniciam juntos após Prep.

function calcularDatasF1_(startDate) {
  const d = function(n) { return addDias_(startDate, n); };
  return {
    prep:        { start: startDate, end: d(7)  },   // 1 semana
    compras:     { start: d(7),      end: d(35) },   // 4 semanas (paralelo)
    fabOsti:     { start: d(7),      end: d(49) },   // 6 semanas (paralelo)
    fabAgricef:  { start: d(7),      end: d(49) },   // 6 semanas (paralelo)
    preMontagem: { start: d(49),     end: d(63) },   // 2 semanas
  };
  // F1 total: 63 dias (~9 semanas) — mais conservador no campo Baseline
}

function calcularSubtasksF2_(diaC, dados) {
  const d = function(n) { return addDias_(diaC, n); };
  const sts = [];

  if (dados.alongamento) {
    // Variante: com alongamento de chassi (+2 semanas)
    sts.push({ nome: 'Alongamento Osti', start: diaC,   end: d(14) });
    sts.push({ nome: 'Montagens Osti',   start: d(14),  end: d(35) });
    sts.push({ nome: 'Montagens',        start: d(35),  end: d(49) });
    sts.push({ nome: 'Montagens 2',      start: d(49),  end: d(56) });
    sts.push({ nome: 'Testes',           start: d(56),  end: d(63) });

  } else if (dados.extensao) {
    // Variante: com fabricação de extensão
    sts.push({ nome: 'Fabricação Extensão', start: diaC,  end: d(14) });
    sts.push({ nome: 'Montagens Osti',      start: d(14), end: d(35) });
    sts.push({ nome: 'Montagens',           start: d(35), end: d(49) });
    sts.push({ nome: 'Testes',              start: d(49), end: d(56) });

  } else {
    // Padrão: 3 subtarefas
    sts.push({ nome: 'Montagens Osti', start: diaC,  end: d(21) });
    sts.push({ nome: 'Montagens',      start: d(21), end: d(35) });
    sts.push({ nome: 'Testes',         start: d(35), end: d(42) });
  }

  return sts;
}

// ─── JIRA REST API ────────────────────────────────────────────

function jiraRequest_(method, path, payload) {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('JIRA_EMAIL');
  const token = props.getProperty('JIRA_TOKEN');

  if (!email || !token) {
    throw new Error(
      'Credenciais Jira ausentes. Configure JIRA_EMAIL e JIRA_TOKEN ' +
      'em: Apps Script → Configurações do projeto → Script Properties.'
    );
  }

  const options = {
    method: method.toLowerCase(),
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(email + ':' + token),
      'Content-Type':  'application/json',
      'Accept':        'application/json',
    },
    muteHttpExceptions: true,
  };

  if (payload && ['post', 'put', 'patch'].includes(method.toLowerCase())) {
    options.payload = JSON.stringify(payload);
  }

  const resp = UrlFetchApp.fetch(JIRA_BASE + path, options);
  const text = resp.getContentText();

  try {
    return JSON.parse(text);
  } catch (_) {
    return { _raw: text, _status: resp.getResponseCode() };
  }
}

// ─── HELPER DE DATAS ─────────────────────────────────────────

function addDias_(dateStr, dias) {
  const d = new Date(dateStr + 'T12:00:00');
  d.setDate(d.getDate() + dias);
  return d.toISOString().split('T')[0];
}
