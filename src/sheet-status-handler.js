function onFormSubmit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();

  // Defina a coluna da qual você quer adicionar o Status (exemplo: coluna F)
  var statusColumn = 13; // Coluna F

  // Preenche o status automaticamente como "Pendente"
  var statusCell = sheet.getRange(row, statusColumn);

  // Verifica se a célula da coluna Status está vazia e preenche com "Pendente"
  if (!statusCell.getValue()) {
    statusCell.setValue("Pendente");
  }

  // Adiciona a validação de dados na coluna Status, se ainda não existir
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendente", "Confirmar", "Recusar"], true)
    .setAllowInvalid(false)
    .build();

  // Aplica a validação de dados para toda a coluna Status (exceto o cabeçalho)
  var lastRow = sheet.getLastRow();
  sheet.getRange(2, statusColumn, lastRow - 1, 1).setDataValidation(validation);
}

function createTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const editedColumn = range.getColumn();
  const editedRow = range.getRow();

  const statusColumn = 13; // número da coluna "Status" (ex: F = 6)
  const emailColumn = 4; // número da coluna onde está o e-mail (ex: C = 3)

  // Verifica se a edição foi na coluna "Status"
  if (editedColumn === statusColumn) {
    const newStatus = range.getValue();
    const email = sheet.getRange(editedRow, emailColumn).getValue();
    const name = sheet.getRange(editedRow, 2).getValue(); // Nome na coluna 2

    if (newStatus === "Pendente") {
      Logger.log(
        `Status permanece como "Pendente" para ${name}. Nenhuma ação necessária.`
      );
      return;
    }

    // Confirmado: Envia e-mail de confirmação
    if (newStatus === "Confirmar") {
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "Confirmação de sua visita ao Jardim Botânico",
          body: `Olá ${name},\n\nTemos o prazer de informar que sua visita foi confirmada!\n\nAtenciosamente,\nEquipe do Jardim Botânico`,
        });
        Logger.log(`E-mail de confirmação enviado para ${email}.`);
      } else {
        Logger.log(
          `Não foi possível enviar o e-mail de confirmação: endereço de e-mail ausente.`
        );
      }
    }

    // Recusar: Envia e-mail de recusa
    if (newStatus === "Recusar") {
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "Atualização do status de sua visita",
          body: `Olá ${name},\n\nLamentamos informar que sua solicitação de visita foi recusada.\n\nAtenciosamente,\nEquipe do Jardim Botânico`,
        });
        Logger.log(`E-mail de recusa enviado para ${email}.`);
      } else {
        Logger.log(
          `Não foi possível enviar o e-mail de recusa: endereço de e-mail ausente.`
        );
      }
    }
  }
}
