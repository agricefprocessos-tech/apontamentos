// Manual Dashboard AGRICEF — gerador docx
// Dependência: npm install -g docx

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, TableOfContents, LevelFormat,
  VerticalAlign, UnderlineType,
} = require("docx");
const fs = require("fs");

// ── Constantes de cor ─────────────────────────────────────────────────────────
const C_YELLOW   = "F0B429";
const C_ORANGE   = "FB923C";
const C_GREEN    = "4ADE80";
const C_BLUE     = "38BDF8";
const C_BGBOX    = "F3F4F6";
const C_DARK     = "1F2937";
const C_WHITE    = "FFFFFF";
const C_CREAM    = "FEFCE8";
const C_GRAY_TXT = "6B7280";

// ── Dimensões A4 com margens 2 cm (1134 DXA ≈ 2 cm) ─────────────────────────
const PAGE_W  = 11906;  // A4 width DXA
const PAGE_H  = 16838;  // A4 height DXA
const MARGIN  = 1134;   // ~2 cm
const CONTENT_W = PAGE_W - MARGIN * 2; // 9638 DXA

// ── Bordas helpers ────────────────────────────────────────────────────────────
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function cellBorder(color = "CCCCCC") {
  const b = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: b, bottom: b, left: b, right: b };
}

// ── Helper: parágrafo de espaçamento ─────────────────────────────────────────
function spacer(pt = 6) {
  return new Paragraph({ children: [new TextRun("")], spacing: { after: pt * 20 } });
}

// ── Helper: heading 1 ────────────────────────────────────────────────────────
function h1(text, addBreak = false) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: addBreak,
    children: [new TextRun({ text, font: "Calibri", size: 32, bold: true, color: C_DARK })],
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C_YELLOW, space: 4 } },
  });
}

// ── Helper: heading 2 ────────────────────────────────────────────────────────
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Calibri", size: 26, bold: true, color: C_YELLOW })],
    spacing: { before: 240, after: 80 },
  });
}

// ── Helper: heading 3 ────────────────────────────────────────────────────────
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Calibri", size: 24, bold: true, color: C_ORANGE })],
    spacing: { before: 180, after: 60 },
  });
}

// ── Helper: parágrafo normal ─────────────────────────────────────────────────
function para(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({
      text,
      font: "Calibri",
      size: opts.size || 22,
      bold: opts.bold || false,
      color: opts.color || C_DARK,
      italics: opts.italics || false,
    })],
    spacing: { after: opts.after !== undefined ? opts.after : 80, before: opts.before || 0 },
    alignment: opts.align || AlignmentType.LEFT,
  });
}

// ── Helper: bullet item ───────────────────────────────────────────────────────
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: [new TextRun({ text, font: "Calibri", size: 22, color: C_DARK })],
    spacing: { after: 60 },
  });
}

// ── Helper: box de destaque (tabela 1 célula) ─────────────────────────────────
function infoBox(label, lines) {
  const cellContent = [
    new Paragraph({
      children: [new TextRun({ text: label, font: "Calibri", size: 20, bold: true, color: C_YELLOW })],
      spacing: { after: 60 },
    }),
    ...lines.map(l =>
      new Paragraph({
        children: [new TextRun({ text: l, font: "Calibri", size: 20, color: C_DARK })],
        spacing: { after: 50 },
      })
    ),
  ];

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { fill: C_BGBOX, type: ShadingType.CLEAR },
            borders: {
              top:    { style: BorderStyle.NONE, size: 0, color: C_BGBOX },
              bottom: { style: BorderStyle.NONE, size: 0, color: C_BGBOX },
              right:  { style: BorderStyle.NONE, size: 0, color: C_BGBOX },
              left:   { style: BorderStyle.SINGLE, size: 16, color: C_YELLOW },
            },
            margins: { top: 120, bottom: 120, left: 200, right: 120 },
            children: cellContent,
          }),
        ],
      }),
    ],
  });
}

// ── Helper: par label+valor inline ───────────────────────────────────────────
function labelPara(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: label + ": ", font: "Calibri", size: 22, bold: true, color: C_DARK }),
      new TextRun({ text: value, font: "Calibri", size: 22, color: C_DARK }),
    ],
    spacing: { after: 80 },
  });
}

// ── CAPA ──────────────────────────────────────────────────────────────────────
function buildCoverPage() {
  return [
    // Espaço superior
    new Paragraph({ children: [new TextRun("")], spacing: { after: 3200 } }),

    // Linha decorativa superior
    new Paragraph({
      children: [new TextRun({ text: "", font: "Calibri" })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C_YELLOW } },
      spacing: { after: 240 },
    }),

    // Título principal
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "Dashboard AGRICEF",
        font: "Calibri",
        size: 72,
        bold: true,
        color: C_YELLOW,
      })],
      spacing: { after: 120 },
    }),

    // Subtítulo
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "Manual de Leitura e Análise",
        font: "Calibri",
        size: 36,
        color: C_DARK,
        bold: false,
      })],
      spacing: { after: 80 },
    }),

    // Linha decorativa abaixo do título
    new Paragraph({
      children: [new TextRun({ text: "", font: "Calibri" })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C_YELLOW } },
      spacing: { after: 240 },
    }),

    // Subtítulo descritivo
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "Gestão Visual da Produção do Chão de Fábrica",
        font: "Calibri",
        size: 28,
        color: C_GRAY_TXT,
        italics: true,
      })],
      spacing: { after: 400 },
    }),

    // Data
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "Maio 2026",
        font: "Calibri",
        size: 24,
        color: C_GRAY_TXT,
      })],
      spacing: { after: 200 },
    }),

    // Versão
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({
        text: "Versão 1.0",
        font: "Calibri",
        size: 20,
        color: C_GRAY_TXT,
      })],
      spacing: { after: 0 },
    }),

    // Quebra de página para o sumário
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ── SUMÁRIO ───────────────────────────────────────────────────────────────────
function buildTOC() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Sumário", font: "Calibri", size: 36, bold: true, color: C_DARK })],
      spacing: { after: 200 },
    }),
    new TableOfContents("", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ── SEÇÃO 1: INTRODUÇÃO ───────────────────────────────────────────────────────
function buildSection1() {
  return [
    h1("1. Introdução"),
    para(
      "O Dashboard AGRICEF é uma ferramenta de gestão visual da produção do chão de fábrica. " +
      "Consolida todos os apontamentos de horas dos operadores (aberturas, fechamentos, paradas, " +
      "retrabalhos) e transforma esses dados em indicadores estratégicos."
    ),
    spacer(8),
    para("Com o Dashboard, a liderança consegue responder imediatamente:", { bold: false }),
    spacer(4),
    bullet("Como está a eficiência da equipe"),
    bullet("Onde estão os gargalos e perdas"),
    bullet("Se as ordens estão sendo concluídas dentro do prazo estimado"),
    bullet("Qual é o ritmo de produção ao longo do tempo"),
    spacer(12),
  ];
}

// ── SEÇÃO 2: FILTROS ──────────────────────────────────────────────────────────
function buildSection2() {
  return [
    h1("2. Filtros"),
    para("Todos os gráficos e indicadores do dashboard respondem simultaneamente aos filtros aplicados. " +
         "Utilize-os para segmentar a análise conforme a necessidade."),
    spacer(8),
    labelPara("Funcionário",     "Filtra todos os gráficos para um operador específico"),
    labelPara("Nº de Série",     "Filtra para uma ordem/projeto específico"),
    labelPara("Operação",        "Filtra por tipo de operação (soldar, cortar, montar, etc.)"),
    labelPara("Código do Item",  "Filtra por código de peça específica"),
    labelPara("Data Início / Data Fim", "Define o período de análise"),
    spacer(8),
    infoBox("Dica de uso:", [
      "Para visão geral do mês, use apenas os filtros de data.",
      "Para investigar um problema específico, combine o filtro de funcionário com a data.",
    ]),
    spacer(12),
  ];
}

// ── SEÇÃO 3: KPIs ─────────────────────────────────────────────────────────────
function buildSection3() {
  const items = [
    {
      num: "3.1", title: "Apontamentos (Total no Período)",
      oQue: "Quantidade total de registros no período filtrado.",
      indica: "Volume de atividade registrada. Número baixo pode indicar sub-registro.",
      decidir: "Se baixo para o período, investigue se operadores estão apontando corretamente.",
    },
    {
      num: "3.2", title: "Funcionários Ativos",
      oQue: "Quantidade de operadores distintos com pelo menos um apontamento.",
      indica: "Tamanho efetivo da equipe que produziu no período.",
      decidir: "Se menor que o esperado, pode haver ausências ou problemas de registro.",
    },
    {
      num: "3.3", title: "Séries em Produção",
      oQue: "Quantidade de números de série distintos com apontamentos.",
      indica: "Quantos projetos/ordens estiveram ativos simultaneamente.",
      decidir: "Número muito alto indica excesso de WIP, diluindo o foco da equipe.",
    },
    {
      num: "3.4", title: "Horas Produtivas",
      oQue: "Soma do tempo efetivo em operações de produção (pares ABERTURA→FECHAMENTO).",
      indica: "O tempo que a equipe realmente produziu valor.",
      decidir: "Proporção saudável é pelo menos 70% do total.",
    },
    {
      num: "3.5", title: "Horas Parada",
      oQue: "Soma do tempo em paradas (Set-up, Manutenção, Falta de material, etc.).",
      indica: "Tempo perdido por causas externas à produção.",
      decidir: "Paradas recorrentes do mesmo tipo exigem ação corretiva estrutural.",
    },
    {
      num: "3.6", title: "Horas Retrabalho",
      oQue: "Soma do tempo gasto em retrabalho (INÍCIO→TÉRMINO DE RETRABALHO).",
      indica: "Tempo perdido por erros e não conformidades.",
      decidir: "Retrabalho acima de 5% do total produtivo é sinal de alerta.",
    },
    {
      num: "3.7", title: "Eficiência Geral",
      oQue: "Percentual de tempo produtivo em relação ao total.",
      formula: "Horas Produtivas ÷ (Produtivas + Parada + Retrabalho + Ausências) × 100",
      referencias: [
        "Acima de 75%: equipe em boa performance",
        "Entre 60% e 75%: atenção às causas de perda",
        "Abaixo de 60%: investigar urgentemente",
      ],
    },
    {
      num: "3.8", title: "Itens Produzidos",
      oQue: "Soma das quantidades realizadas informadas nos fechamentos.",
      indica: "Volume físico produzido no período.",
      decidir: "",
    },
    {
      num: "3.9", title: "Horas Ausência",
      oQue: "Estimativa de horas perdidas por dias sem apontamento (8h por dia sem registro).",
      indica: "",
      decidir: "Ausências sistemáticas do mesmo operador em dias específicos indicam padrão.",
    },
    {
      num: "3.10", title: "OEE (Disponibilidade × Performance × Qualidade)",
      oQue: "Indicador mundial de eficiência industrial.",
      referencias: [
        "Classe Mundial: acima de 85%",
        "Aceitável: acima de 65%",
      ],
      decidir: "OEE baixo com disponibilidade alta = problema de qualidade ou ritmo. Disponibilidade baixa = muitas paradas.",
    },
    {
      num: "3.11", title: "Taxa de Retrabalho",
      oQue: "Percentual de operações fechadas que geraram retrabalho registrado.",
      decidir: "Acima de 10% é crítico. Acione análise de causa raiz (5 Porquês).",
    },
    {
      num: "3.12", title: "Backlog Aberto",
      oQue: "Número de aberturas sem fechamento correspondente.",
      decidir: "Use o painel Backlog para identificar cada ordem e o botão Fechar para registrar o fechamento direto pela dashboard.",
    },
    {
      num: "3.13", title: "WIP Médio (Work In Progress)",
      oQue: "Número médio de operações abertas simultaneamente. Calculado pela Lei de Little: WIP = Lead Time × Throughput.",
      indica: "Quantas ordens a equipe carrega em paralelo em média.",
      decidir: "WIP alto com Lead Time alto = gargalo. Trabalhe com WIP limitado (método Kanban).",
    },
    {
      num: "3.14", title: "Throughput Médio",
      oQue: "Média de operações fechadas por mês sobre o histórico completo.",
      decidir: "Se Throughput Médio é 40 ops/mês e precisa entregar 60, é necessário reforço ou revisão de prazo.",
    },
    {
      num: "3.15", title: "Coeficiente de Variação",
      oQue: "Desvio padrão do throughput mensal dividido pela média (%).",
      indica: "Consistência e previsibilidade da equipe.",
      referencias: [
        "Abaixo de 20%: muito consistente, produção previsível",
        "Entre 20% e 40%: variabilidade moderada, investigar sazonalidade",
        "Acima de 40%: produção imprevisível, eliminar causas de variação",
      ],
    },
    {
      num: "3.16", title: "Lead Time Médio",
      oQue: "Tempo médio em horas entre abertura e fechamento de uma operação.",
      decidir: "Se consistentemente maior que o estimado, revise estimativas ou investigue causas de demora.",
    },
  ];

  const nodes = [h1("3. KPIs — Cartões de Resumo")];
  nodes.push(para(
    "Os cartões de resumo apresentam os indicadores-chave do período filtrado em uma única tela. " +
    "Cada cartão é projetado para uma leitura rápida e uma decisão imediata."
  ));
  nodes.push(spacer(8));

  for (const item of items) {
    nodes.push(h2(`${item.num} ${item.title}`));
    nodes.push(labelPara("O que é", item.oQue));
    if (item.indica) nodes.push(labelPara("O que indica", item.indica));
    if (item.formula) nodes.push(labelPara("Fórmula", item.formula));
    if (item.referencias) {
      nodes.push(para("Referências:", { bold: true }));
      item.referencias.forEach(r => nodes.push(bullet(r)));
    }
    if (item.decidir) {
      nodes.push(infoBox("Como decidir:", [item.decidir]));
    }
    nodes.push(spacer(8));
  }

  return nodes;
}

// ── SEÇÃO 4: GRÁFICOS E TABELAS ───────────────────────────────────────────────
function buildSection4() {
  const graficos = [
    {
      num: "4.1", title: "Eficiência por Funcionário",
      oQue: "Tabela ou gráfico com horas produtivas, paradas, retrabalho e EE% por operador.",
      analisar: ["EE% muito abaixo da média", "Muitas horas de parada", "Alto retrabalho"],
      decisoes: ["Redistribuir tarefas", "Identificar quem precisa de mentoria", "Reconhecer melhores performers"],
    },
    {
      num: "4.2", title: "Tempo por Operação",
      oQue: "Barras horizontais com total de horas por tipo de operação, eficiência e meta configurável.",
      analisar: ["Operações que consomem mais horas", "Operações abaixo da meta"],
      decisoes: ["Investir em equipamentos e treinamento", "Definir e monitorar metas por operação"],
    },
    {
      num: "4.3", title: "Top Códigos de Item",
      oQue: "Ranking dos itens com maior tempo de operação no período.",
      analisar: ["Itens que concentram mais horas (itens gargalo)"],
      decisoes: ["Priorizar melhorias nos itens mais consumidores", "Revisar estimativas"],
    },
    {
      num: "4.4", title: "Produção por Nº de Série",
      oQue: "Horas por número de série com detalhe de operações e itens.",
      analisar: ["Séries com muitas horas", "Séries com operações em aberto"],
      decisoes: ["Priorizar fechamento das séries críticas", "Alocar mais operadores nas atrasadas"],
    },
    {
      num: "4.5", title: "Paradas e Retrabalhos",
      oQue: "Barras com tempo perdido por tipo de parada e tipo de retrabalho.",
      analisar: ["Tipo de parada mais frequente", "Set-up longo", "Falta de Material frequente"],
      decisoes: ["Plano de ação para causa raiz", "Manutenção preventiva", "Revisar planejamento de materiais"],
    },
    {
      num: "4.6", title: "Ranking de Produtividade",
      oQue: "Lista ordenada dos funcionários por horas produtivas.",
      analisar: ["Grande diferença entre 1° e último (desbalanceamento)", "Constância no topo ou no final"],
      decisoes: ["Programas de reconhecimento", "Mentoria entre operadores"],
    },
    {
      num: "4.7", title: "Últimos Apontamentos",
      oQue: "Feed com registros mais recentes (operador, tipo, operação, série).",
      analisar: ["Apontamentos muito antigos", "Muitas aberturas sem fechamento recente"],
      decisoes: [],
    },
    {
      num: "4.8", title: "Previsto x Realizado",
      oQue: "Comparação entre horas estimadas e realmente gastas por código de item.",
      analisar: ["Itens com realizado muito maior que previsto", "Itens onde realizado é menor"],
      decisoes: ["Atualizar tabela de estimativas", "Usar desvios para planejar prazos das próximas ordens"],
    },
    {
      num: "4.9", title: "Assiduidade e Ausências",
      oQue: "Calendário heatmap por funcionário — dias com registro (verde) vs dias sem registro (vermelho).",
      analisar: ["Funcionários com muitos dias sem registro", "Padrões de ausência (toda segunda, toda sexta)"],
      decisoes: ["Conversa individual com alta ausência", "Revisão do processo de registro"],
    },
    {
      num: "4.10", title: "OEE e Tempo de Ciclo",
      oQue: "OEE e Lead Time por série de produção.",
      analisar: ["Séries com OEE muito baixo", "Lead Time muito longo entre operações"],
      decisoes: [],
    },
    {
      num: "4.11", title: "Tendência Semanal",
      oQue: "Linha do tempo semanal com evolução da eficiência.",
      analisar: ["Tendência de subida (melhora) ou queda (investigar)", "Picos e vales sazonais"],
      decisoes: [
        "Queda consistente por 3+ semanas = reunião de análise de causa",
        "Usar picos como referência de capacidade máxima",
      ],
    },
    {
      num: "4.12", title: "Heatmap por Hora do Dia",
      oQue: "Mapa de calor com intensidade de apontamentos por hora.",
      analisar: [
        "Horários sem apontamento no meio da jornada",
        "Concentração no final do turno",
      ],
      decisoes: [
        "Orientar registro em tempo real",
        "Alocar atividades de menor complexidade nos horários de menor produtividade",
      ],
    },
    {
      num: "4.13", title: "Backlog — Ordens em Aberto",
      oQue: "Tabela de todas as aberturas sem fechamento, calculado do histórico completo (cobre sistemas antigo e novo).",
      colunas: "Funcionário, Operação, Cód. Item, Série, Qtd. Planejada, Abertura, Botão Fechar.",
      analisar: [
        "Ordens abertas há mais de 1 dia",
        "Muitas ordens do mesmo operador",
        "Ordens de operações críticas",
      ],
      fechar: [
        "1. Clique em Fechar na linha da ordem",
        "2. Informe quantidade realizada e observação",
        "3. Clique em Confirmar Fechamento",
        "4. A lista atualiza automaticamente",
      ],
    },
    {
      num: "4.14", title: "Produtividade por Dia da Semana",
      oQue: "Barras com média de horas produtivas por dia da semana.",
      analisar: ["Dias com produtividade consistentemente baixa", "Comparar pares de dias"],
      decisoes: ["Programar atividades administrativas nos dias historicamente menos produtivos"],
    },
    {
      num: "4.15", title: "Auditoria e Validação",
      oQue: "Ferramenta de verificação da qualidade dos dados com 3 abas:",
      abas: [
        "Pares Calculados: lista de pares ABERTURA→FECHAMENTO com duração",
        "Por Funcionário: análise de pares por operador",
        "Inconsistências: detecta aberturas sem fechamento, fechamentos sem abertura, durações suspeitas",
      ],
      analisar: ["Pares com duração >8h", "Alta contagem de inconsistências"],
      decisoes: ["Usar mensalmente para limpar histórico", "Treinar operadores com mais inconsistências"],
    },
    {
      num: "4.16", title: "Análise de Desempenho Individual",
      oQue: "Análise aprofundada por operador com velocidade comparada aos pares, consistência e taxa de retrabalho.",
      analisar: [
        "Operador muito mais lento que os pares",
        "Muito mais rápido (verificar qualidade)",
        "Alta taxa de retrabalho individual",
      ],
      decisoes: ["Plano de desenvolvimento individual", "Documentar boas práticas dos mais rápidos"],
    },
  ];

  const nodes = [h1("4. Seções — Gráficos e Tabelas", true)];
  nodes.push(para(
    "Esta seção descreve cada painel do dashboard, o que ele apresenta, o que deve ser analisado " +
    "e as decisões que podem ser tomadas a partir dos dados."
  ));
  nodes.push(spacer(8));

  for (const g of graficos) {
    nodes.push(h2(`${g.num} ${g.title}`));
    nodes.push(labelPara("O que é", g.oQue));
    if (g.colunas)  nodes.push(labelPara("Colunas", g.colunas));
    if (g.abas) {
      nodes.push(para("Abas disponíveis:", { bold: true }));
      g.abas.forEach(a => nodes.push(bullet(a)));
    }
    if (g.analisar && g.analisar.length) {
      nodes.push(infoBox("O que analisar:", g.analisar));
      nodes.push(spacer(4));
    }
    if (g.fechar) {
      nodes.push(para("Como usar o botão Fechar:", { bold: true }));
      g.fechar.forEach(f => nodes.push(bullet(f)));
    }
    if (g.decisoes && g.decisoes.length) {
      nodes.push(infoBox("Decisões:", g.decisoes));
    }
    nodes.push(spacer(10));
  }

  return nodes;
}

// ── SEÇÃO 5: NOVOS INDICADORES ESTRATÉGICOS ──────────────────────────────────
function buildSection5() {
  const nodes = [h1("5. Novos Indicadores Estratégicos", true)];

  // 5.1
  nodes.push(h2("5.1 Panorama de Produção Anual"));

  nodes.push(h3("Modo Throughput + Entradas"));
  nodes.push(para(
    "Barras duplas por mês. Verde = operações fechadas (Throughput). " +
    "Azul = novas aberturas (Entradas). Tooltip mostra saldo do mês."
  ));
  nodes.push(infoBox("O que analisar:", [
    "Meses com entradas > fechamentos (backlog crescendo)",
    "Fechamentos > entradas (reduzindo backlog)",
    "Saldo consistentemente negativo (risco de acúmulo)",
  ]));
  nodes.push(spacer(6));
  nodes.push(infoBox("Decisões:", [
    "Pico de entradas = planejar reforço",
    "Identificar sazonalidade para planejar férias e manutenções",
  ]));
  nodes.push(spacer(10));

  nodes.push(h3("Modo CFD — Diagrama de Fluxo Cumulativo"));
  nodes.push(para(
    "Três linhas acumuladas no tempo: Aberturas acumuladas (azul), Fechamentos acumulados (verde), " +
    "WIP = área amarela entre as duas curvas."
  ));
  nodes.push(para("Como ler o CFD:", { bold: true }));
  nodes.push(bullet("Linhas paralelas e próximas: produção fluindo bem, pouco WIP"));
  nodes.push(bullet("Área amarela crescendo: WIP aumentando — equipe abre mais do que fecha"));
  nodes.push(bullet("Área amarela constante: produção em equilíbrio"));
  nodes.push(bullet("Linhas se afastando abruptamente: evento de interrupção ou crise"));
  nodes.push(bullet("Curvatura descendente da linha verde: queda no ritmo de fechamento"));
  nodes.push(spacer(6));
  nodes.push(infoBox("Decisões:", [
    "Área crescendo por 2+ semanas = revisar carga de trabalho",
    "Usar o CFD como termômetro em reuniões semanais",
  ]));
  nodes.push(spacer(10));

  // 5.2
  nodes.push(h2("5.2 Dentro / Fora do Prazo por Operação"));
  nodes.push(para(
    "Barras horizontais empilhadas por operação. Verde = no prazo (≤115% do estimado). " +
    "Vermelho = fora do prazo (>115%). Card lateral mostra percentual global."
  ));
  nodes.push(labelPara("Critério", "Operação \"no prazo\" se horas reais não ultrapassam 115% das horas estimadas."));
  nodes.push(infoBox("O que analisar:", [
    "Operações com alto percentual fora do prazo",
    "Taxa global abaixo de 70%",
  ]));
  nodes.push(spacer(6));
  nodes.push(infoBox("Decisões:", [
    "Atualizar estimativas para itens que consistentemente ultrapassam 115%",
    "Investigar se excesso é por complexidade real ou ineficiência",
  ]));
  nodes.push(spacer(10));

  // 5.3
  nodes.push(h2("5.3 Tabela de Projetos por Série"));
  nodes.push(para(
    "Tabela de todos os Nºs de série com status, operações fechadas/abertas, itens distintos, " +
    "horas totais, data início e último fechamento."
  ));
  nodes.push(para("Status possíveis:", { bold: true }));
  nodes.push(bullet("EM ANDAMENTO: tem operação em aberto"));
  nodes.push(bullet("CONCLUIDO: todas as operações fechadas"));
  nodes.push(bullet("ABERTO: aberturas sem fechamentos ainda"));
  nodes.push(spacer(6));
  nodes.push(infoBox("O que analisar:", [
    "Séries Em Andamento há muito tempo (projetos travados)",
    "Muitas ops abertas vs fechadas (risco de atraso)",
    "Horas muito acima do esperado",
  ]));
  nodes.push(spacer(6));
  nodes.push(infoBox("Decisões:", [
    "Usar em reuniões semanais de acompanhamento",
    "Priorizar séries com maior proporção de ops abertas",
    "Identificar concluídas para liberar capacidade",
  ]));
  nodes.push(spacer(12));

  return nodes;
}

// ── SEÇÃO 6: FLUXO DE ANÁLISE ─────────────────────────────────────────────────
function buildSection6() {
  const nodes = [h1("6. Fluxo de Análise Recomendado", true)];
  nodes.push(para(
    "Seguir uma rotina estruturada garante que os dados do dashboard sejam usados para gerar " +
    "decisões consistentes e melhoria contínua."
  ));
  nodes.push(spacer(8));

  nodes.push(h2("6.1 Toda Segunda-Feira (15 minutos)"));
  nodes.push(bullet("Verifique o Backlog Aberto — feche ou acione operadores com ordens esquecidas"));
  nodes.push(bullet("Confira a Eficiência Geral e compare com a semana anterior"));
  nodes.push(bullet("Olhe o Throughput vs Entradas — o saldo está positivo?"));
  nodes.push(bullet("Identifique o maior tipo de parada da semana anterior"));
  nodes.push(bullet("Verifique a Tabela de Projetos — alguma série crítica em risco?"));
  nodes.push(spacer(10));

  nodes.push(h2("6.2 Mensalmente"));
  nodes.push(bullet("Execute a Auditoria & Validação para garantir qualidade dos dados"));
  nodes.push(bullet("Analise a Tendência Semanal do mês completo"));
  nodes.push(bullet("Compare o Coeficiente de Variação com o mês anterior"));
  nodes.push(bullet("Revise estimativas usando o gráfico Previsto × Realizado"));
  nodes.push(spacer(12));

  return nodes;
}

// ── SEÇÃO 7: GLOSSÁRIO ────────────────────────────────────────────────────────
function buildSection7() {
  const terms = [
    ["ABERTURA",                  "Registro de início de uma operação de produção"],
    ["FECHAMENTO",                "Registro de conclusão de uma operação"],
    ["BACKLOG",                   "Ordens abertas sem fechamento correspondente"],
    ["WIP (Work In Progress)",    "Quantidade de trabalho em curso simultaneamente"],
    ["THROUGHPUT",                "Velocidade de entrega — operações concluídas por período"],
    ["LEAD TIME",                 "Tempo total entre início e fim de uma operação"],
    ["CFD (Cumulative Flow Diagram)", "Gráfico acumulativo que mostra o fluxo de produção ao longo do tempo"],
    ["OEE (Overall Equipment Effectiveness)", "Indicador mundial de eficiência industrial"],
    ["EE%",                       "Eficiência do Funcionário — proporção do tempo em atividade produtiva"],
    ["Coeficiente de Variação",   "Medida de consistência — quanto menor, mais previsível a produção"],
    ["Lei de Little",             "WIP = Lead Time x Throughput — lei fundamental do fluxo de produção"],
    ["Poka-yoke",                 "Mecanismo de prevenção de erros — no sistema impede dupla abertura simultânea"],
  ];

  const headerBorder = cellBorder(C_YELLOW);
  const rowBorder    = cellBorder("CCCCCC");

  const col1 = Math.round(CONTENT_W * 0.35);
  const col2 = CONTENT_W - col1;

  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        width: { size: col1, type: WidthType.DXA },
        shading: { fill: C_YELLOW, type: ShadingType.CLEAR },
        borders: headerBorder,
        margins: { top: 100, bottom: 100, left: 120, right: 80 },
        children: [new Paragraph({
          children: [new TextRun({ text: "Termo", font: "Calibri", size: 22, bold: true, color: "000000" })],
        })],
      }),
      new TableCell({
        width: { size: col2, type: WidthType.DXA },
        shading: { fill: C_YELLOW, type: ShadingType.CLEAR },
        borders: headerBorder,
        margins: { top: 100, bottom: 100, left: 120, right: 80 },
        children: [new Paragraph({
          children: [new TextRun({ text: "Definição", font: "Calibri", size: 22, bold: true, color: "000000" })],
        })],
      }),
    ],
  });

  const dataRows = terms.map((row, idx) => {
    const fill = idx % 2 === 0 ? C_CREAM : C_WHITE;
    return new TableRow({
      children: [
        new TableCell({
          width: { size: col1, type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR },
          borders: rowBorder,
          margins: { top: 80, bottom: 80, left: 120, right: 80 },
          children: [new Paragraph({
            children: [new TextRun({ text: row[0], font: "Calibri", size: 20, bold: true, color: C_DARK })],
          })],
        }),
        new TableCell({
          width: { size: col2, type: WidthType.DXA },
          shading: { fill, type: ShadingType.CLEAR },
          borders: rowBorder,
          margins: { top: 80, bottom: 80, left: 120, right: 80 },
          children: [new Paragraph({
            children: [new TextRun({ text: row[1], font: "Calibri", size: 20, color: C_DARK })],
          })],
        }),
      ],
    });
  });

  return [
    h1("7. Glossário", true),
    para("Definições dos termos e siglas utilizados no Dashboard AGRICEF."),
    spacer(8),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [col1, col2],
      rows: [headerRow, ...dataRows],
    }),
    spacer(12),
  ];
}

// ── HEADER / FOOTER ───────────────────────────────────────────────────────────
function buildHeader() {
  return new Header({
    children: [
      new Paragraph({
        children: [new TextRun({
          text: "Manual de Leitura e Análise — Dashboard AGRICEF",
          font: "Calibri",
          size: 18,
          color: C_GRAY_TXT,
        })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C_YELLOW, space: 4 } },
        spacing: { after: 80 },
      }),
    ],
  });
}

function buildFooter() {
  return new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: "AGRICEF — Confidencial",
            font: "Calibri",
            size: 16,
            color: C_GRAY_TXT,
          }),
          new TextRun({
            children: ["\t", PageNumber.CURRENT],
            font: "Calibri",
            size: 16,
            color: C_GRAY_TXT,
          }),
        ],
        tabStops: [{ type: "right", position: CONTENT_W }],
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD", space: 4 } },
        spacing: { before: 80 },
      }),
    ],
  });
}

// ── MONTAR DOCUMENTO ──────────────────────────────────────────────────────────
const children = [
  ...buildCoverPage(),
  ...buildTOC(),
  ...buildSection1(),
  ...buildSection2(),
  ...buildSection3(),
  ...buildSection4(),
  ...buildSection5(),
  ...buildSection6(),
  ...buildSection7(),
];

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 22 },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Calibri", size: 32, bold: true, color: C_DARK },
        paragraph: {
          spacing: { before: 360, after: 120 },
          outlineLevel: 0,
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Calibri", size: 26, bold: true, color: C_YELLOW },
        paragraph: {
          spacing: { before: 240, after: 80 },
          outlineLevel: 1,
        },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Calibri", size: 24, bold: true, color: C_ORANGE },
        paragraph: {
          spacing: { before: 180, after: 60 },
          outlineLevel: 2,
        },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 720, hanging: 360 } },
              run: { font: "Calibri" },
            },
          },
          {
            level: 1,
            format: LevelFormat.BULLET,
            text: "◦",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 1080, hanging: 360 } },
              run: { font: "Calibri" },
            },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      headers: { default: buildHeader() },
      footers: { default: buildFooter() },
      children,
    },
  ],
});

const outputPath = "C:\\agricef-apontamento\\dasboard\\Manual_Dashboard_AGRICEF.docx";

Packer.toBuffer(doc)
  .then(buffer => {
    fs.writeFileSync(outputPath, buffer);
    const stats = fs.statSync(outputPath);
    console.log("Documento gerado com sucesso!");
    console.log("Caminho: " + outputPath);
    console.log("Tamanho: " + (stats.size / 1024).toFixed(1) + " KB");
  })
  .catch(err => {
    console.error("Erro ao gerar documento:", err);
    process.exit(1);
  });
