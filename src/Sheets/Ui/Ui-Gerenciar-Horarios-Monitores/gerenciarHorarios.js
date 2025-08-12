function abrirSidebarGerenciarHorarios() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar_GerenciarHorarios").setTitle(
    "Gerenciar Horários dos Monitores"
  );
  SpreadsheetApp.getUi().showSidebar(html);
}

function acessaAbaNaPlanilha() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = planilha.getSheetByName("Horário Monitores");
  if (!sheet) {
    throw new Error("Aba 'Horário Monitores' não encontrada.");
  }
  return sheet;
}

function getMonitores() {
  const sheet = acessaAbaNaPlanilha();

  const dados = sheet.getDataRange().getValues();
  const monitores = [];

  for (let i = 1; i < dados.length; i++) {
    const nome = dados[i][0];
    const numLinhaRealPlanilha = i + 1;
    if (nome) {
      monitores.push({ nomeMonitorAchado: nome, linha: numLinhaRealPlanilha });
    }
  }

  return monitores;
}

function atualizarNomeEmailMonitor(nomeOriginal, novoNome, novoEmail) {
  const sheet = acessaAbaNaPlanilha();

  const COLUNA_A_NOME = 1;
  const COLUNA_B_EMAIL = 2;
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0] === nomeOriginal) {
      const numLinhaRealPlanilha = i + 1;
      sheet.getRange(numLinhaRealPlanilha, COLUNA_A_NOME).setValue(novoNome);
      sheet.getRange(numLinhaRealPlanilha, COLUNA_B_EMAIL).setValue(novoEmail);
      return;
    }
  }

  throw new Error("Monitor não encontrado: " + nomeOriginal);
}

function adicionarMonitor(nomeInserir, emailInserir) {
  const sheet = acessaAbaNaPlanilha();

  if (!nomeInserir || !emailInserir) {
    throw new Error("Nome e Email do monitor são obrigatórios");
  }

  const ultimaLinhaPreenchida = sheet.getLastRow();
  const intervaloDeDados = sheet.getDataRange();
  const totalDeColunas = intervaloDeDados.getNumColumns();

  sheet.insertRowAfter(ultimaLinhaPreenchida);
  const linhaDoNovoMonitor = ultimaLinhaPreenchida + 1;

  sheet.getRange(linhaDoNovoMonitor, 1).setValue(nomeInserir);
  sheet.getRange(linhaDoNovoMonitor, 2).setValue(emailInserir);

  const intervaloParaCopiarFormatacao = sheet.getRange(ultimaLinhaPreenchida, 1, 1, totalDeColunas);
  const intervaloParaAplicarFormato = sheet.getRange(linhaDoNovoMonitor, 1, 1, totalDeColunas);

  intervaloParaCopiarFormatacao.copyTo(intervaloParaAplicarFormato, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  if (totalDeColunas > 2) {
    const colunasAdicionais = totalDeColunas - 2;
    sheet.getRange(linhaDoNovoMonitor, 3, 1, colunasAdicionais).clearContent();
  }

  return `Monitor ${nomeInserir} adicionado com sucesso na linha ${linhaDoNovoMonitor}`;
}

function removerMonitor(nomeRemover) {
  if (!nomeRemover || typeof nomeRemover !== "string") {
    throw new Error("Nome do monitor é obrigatório e deve ser uma string");
  }

  const planilha = acessaAbaNaPlanilha();
  const monitores = getMonitores();

  const monitorParaRemover = monitores.find((m) => m.nomeMonitorAchado === nomeRemover);

  if (!monitorParaRemover) {
    throw new Error(`Monitor "${nomeRemover}" não encontrado na planilha`);
  }

  planilha.deleteRow(monitorParaRemover.linha);

  return `Monitor "${nomeRemover}" foi removido com sucesso`;
}

function inserirNovaLinhaNaPlanilha(planilha) {
  const ultimaLinha = planilha.getLastRow();
  planilha.insertRowAfter(ultimaLinha);
  return ultimaLinha + 1;
}

function copiarFormatacaoDaLinhaAnterior(planilha, linha) {
  const totalColunas = planilha.getDataRange().getNumColumns();
  const linhaAnterior = linha - 1;

  planilha
    .getRange(linhaAnterior, 1, 1, totalColunas)
    .copyTo(planilha.getRange(linha, 1, 1, totalColunas), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function testeGetMonitores() {
  try {
    const monitores = getMonitores();
    Logger.log(`Quantidade de monitores encontrados: ${monitores.length}`);

    if (monitores.length > 0) {
      monitores.forEach((m, index) => {
        Logger.log(`#${index + 1} - Nome: ${m.nome} | Linha: ${m.linha}`);
      });
    } else {
      Logger.log("Nenhum monitor encontrado na planilha.");
    }
  } catch (erro) {
    Logger.log("❌ Erro ao executar getMonitores: " + erro.message);
  }
}

function testeAtualizarHorarioMonitor() {
  try {
    const nome = "Lucas"; // Nome do monitor (como está na planilha)
    const dia = "quarta"; // Pode ser: segunda, terca, quarta, quinta, sexta
    const novoHorario = "Integral"; // Pode ser: Manhã, Tarde, Integral, etc.

    atualizarHorarioMonitorNoSheet(nome, dia, novoHorario);
    Logger.log(`✅ Horário atualizado com sucesso para ${nome} em ${dia}: ${novoHorario}`);
  } catch (erro) {
    Logger.log("❌ Erro ao atualizar horário: " + erro.message);
  }
}
