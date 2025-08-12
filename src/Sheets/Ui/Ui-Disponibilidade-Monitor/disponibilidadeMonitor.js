function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Visitas")
    .addItem("Buscar monitor disponível", "showAvailabilitySidebar")
    .addItem("Gerenciar Horário Monitores", "abrirSidebarGerenciarHorarios")
    .addToUi();
}

function showAvailabilitySidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar_Disponibilidade").setTitle(
    "Agendamento Jardim Botânico UFSM"
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

function getDiaSemanaEPeriodo(dataStr, horarioStr) {
  const data = new Date(dataStr);
  const diasSemana = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  const diaSemana = diasSemana[data.getDay()];

  if (data.getDay() === 0 || data.getDay() === 6) {
    return {
      diaSemana: null,
      periodo: null,
      mensagem: "Agendamentos disponíveis apenas de segunda a sexta.",
    };
  }

  const periodo =
    {
      "A - 8:30 às 10:00": "Manhã",
      "B - 10:30 às 12:00": "Manhã",
      "C - 13:30 às 15:00": "Tarde",
      "D - 15:30 às 17:00": "Tarde",
    }[horarioStr.trim()] || "";

  return { diaSemana, periodo };
}

function buscarMonitorDisponivel(data, horario) {
  const { diaSemana, periodo } = getDiaSemanaEPeriodo(data, horario);
  if (!diaSemana || !periodo) {
    return { status: "erro", mensagem: "Dia ou período inválido" };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Horário Monitores");
  const dados = sheet.getDataRange().getValues();
  const cabecalho = dados[0];
  const colDiaSemana = cabecalho.indexOf(diaSemana);

  if (colDiaSemana === -1) {
    return { status: "erro", mensagem: `Coluna ${diaSemana} não encontrada.` };
  }

  const monitoresDisponiveis = dados
    .slice(1)
    .filter((linha) => linha[colDiaSemana] === periodo || linha[colDiaSemana] === "Integral")
    .map(([nome, email]) => ({ nome, email }));

  if (monitoresDisponiveis.length > 0) {
    return {
      status: "ok",
      monitores: monitoresDisponiveis,
    };
  } else {
    return {
      status: "vazio",
      mensagem: "Nenhum monitor disponível para este horário.",
    };
  }
}

function verificarMonitorSidebar() {
  try {
    const dados = obterDadosVisitaSelecionado();
    const { startTime, endTime } = ajustarHorario(dados.dataVisitante, dados.horarioVisitante);
    const disponivel = verificarDisponibilidade(startTime, endTime);

    if (!disponivel) {
      return {
        status: "ocupado",
        mensagem: "Já existe um evento agendado nesse horário.",
        detalhesVisita: {
          data: Utilities.formatDate(new Date(dados.dataVisitante), Session.getScriptTimeZone(), "dd/MM/yyyy"),
          horario: dados.horarioVisitante,
          escola: dados.escolaVisitante,
          responsavel: dados.responsavelVisitante,
        },
      };
    }

    const monitoresInfo = buscarMonitorDisponivel(dados.dataVisitante, dados.horarioVisitante);

    return {
      status: monitoresInfo.status,
      monitores: monitoresInfo.monitores || [],
      mensagem: monitoresInfo.mensagem || "",
      detalhesVisita: {
        data: Utilities.formatDate(new Date(dados.dataVisitante), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        horario: dados.horarioVisitante,
        escola: dados.escolaVisitante,
        responsavel: dados.responsavelVisitante,
      },
    };
  } catch (erro) {
    return {
      status: "erro",
      mensagem: erro.toString(),
    };
  }
}

function obterDadosVisitaSelecionado() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  const dados = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  return {
    responsavelVisitante: dados[4],
    horarioVisitante: dados[12],
    emailVisitante: dados[1],
    escolaVisitante: dados[3],
    dataVisitante: dados[10],

    telefoneVisitante: dados[2],

    municipio: dados[5],
    endereco: dados[6],
    tipoInstituicao: dados[7],
    nivel: dados[8],
    anoSerie: dados[9],
    numVisitantes: dados[11],
    atividades: dados[13],
    opcaoChuva: dados[14],
    conteudoSala: dados[15],
    piquenique: dados[16],
    termoResponsabilidade: dados[17],
  };
}

function confirmarAgendamentoPeloSidebar(nomeMonitor, emailMonitor) {
  const dados = obterDadosVisitaSelecionado();

  const horarioCompleto = dados.horarioVisitante;
  const horarioSomente = horarioCompleto.split(" - ")[1] || horarioCompleto;

  Logger.log(`Confirmar agendamento para: 
  Responsavel: ${dados.responsavelVisitante}, 
  Email Visitante: ${dados.emailVisitante}, 
  Horario Visita: ${dados.horarioVisitante}, 
  Escola Visitante: ${dados.escolaVisitante}, 
  Data Visita: ${dados.dataVisitante}`);

  const { startTime, endTime } = ajustarHorario(dados.dataVisitante, dados.horarioVisitante);

  const isAvailable = verificarDisponibilidade(startTime, endTime);
  Logger.log(`Disponibilidade: ${isAvailable}`);

  if (!isAvailable) {
    return "Horário indisponível.";
  } else {
    const description = `Visitante: ${dados.responsavelVisitante}\nE-mail: ${dados.emailVisitante}\nHorário: ${horarioSomente}`;

    inserirEventoAgenda({
      name: dados.responsavelVisitante,
      email: dados.emailVisitante,
      startTime,
      endTime,
      description,
    });
    Logger.log("funfou inserção!");

    dados.monitorDestinado = nomeMonitor;
    dados.emailMonitor = emailMonitor;
    salvarVisitaConfirmadaParaGuiaVisitasConfirmadas(dados, horarioSomente);

    enviarEmailVisita("VISITA_CONFIRMADA", dados);

    return "Evento inserido na agenda no Google Calendar e e-mail de confirmação enviado!";
  }
}

function confirmarAgendamentoPorBotaoSidebar(nomeMonitor, emailMonitor) {
  const dados = obterDadosVisitaSelecionado();
  dados.monitorDestinado = nomeMonitor;
  dados.emailMonitor = emailMonitor;
  return processarAgendamentoSidebar("confirmar", dados);
}

function recusarAgendamentoPorBotaoSidebar() {
  const dados = obterDadosVisitaSelecionado();
  return processarAgendamentoSidebar("recusar", dados);
}

function salvarVisitaConfirmadaParaGuiaVisitasConfirmadas(dados, horarioSomente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaConfirmadas = ss.getSheetByName("Visitas Confirmadas");

  planilhaConfirmadas.appendRow([
    dados.dataVisitante,
    horarioSomente,
    dados.responsavelVisitante,
    dados.emailVisitante,
    dados.telefoneVisitante,
    dados.monitorDestinado || "",
    dados.emailMonitor || "",
    dados.escolaVisitante,
    dados.municipio || "",
    dados.endereco || "",
    dados.tipoInstituicao || "",
    dados.nivel || "",
    dados.anoSerie || "",
    dados.numVisitantes || "",
    dados.atividades || "",
    dados.opcaoChuva || "",
    dados.conteudoSala || "",
    dados.piquenique || "",
    dados.termoResponsabilidade || "",
  ]);
}

function processarAgendamentoSidebar(acao, dados) {
  const horarioCompleto = dados.horarioVisitante;
  const horarioSomente = horarioCompleto.split(" - ")[1] || horarioCompleto;

  const { startTime, endTime } = ajustarHorario(dados.dataVisitante, dados.horarioVisitante);

  if (acao === "confirmar") {
    const isAvailable = verificarDisponibilidade(startTime, endTime);
    if (!isAvailable) return "Horário indisponível.";

    inserirEventoAgenda({
      name: dados.responsavelVisitante,
      email: dados.emailVisitante,
      startTime,
      endTime,
      description: `Visitante: ${dados.responsavelVisitante}\nE-mail: ${dados.emailVisitante}\nHorário: ${horarioSomente}`,
    });

    salvarVisitaConfirmadaParaGuiaVisitasConfirmadas(dados, horarioSomente);

    enviarEmailVisita("VISITA_CONFIRMADA", dados);
    return "Agendamento confirmado e e-mail de aviso enviado.";
  }

  if (acao === "recusar") {
    enviarEmailVisita("VISITA_RECUSADA", dados);
    return "Solicitação recusada e e-mail de aviso enviado.";
  }
  return "Ação desconhecida.";
}

function test_ObterDadosVisitaSelecionado() {
  const dados = obterDadosVisitaSelecionado();
  Logger.log(dados);
}

function test_confirmarAgendamentoPeloSidebar() {
  const nomeMonitorTeste = "João da Silva";
  const emailMonitorTeste = "joao.silva@exemplo.com";

  const dadosVisita = obterDadosVisitaSelecionado();
  Logger.log("Dados da visita selecionada:");
  Logger.log(dadosVisita);

  const resultado = confirmarAgendamentoPeloSidebar(nomeMonitorTeste, emailMonitorTeste);

  Logger.log("Resultado da confirmação:");
  Logger.log(resultado);
}
