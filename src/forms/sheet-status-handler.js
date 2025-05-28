function onFormSubmit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();

  var statusColumn = 13;

  var statusCell = sheet.getRange(row, statusColumn);
  if (!statusCell.getValue()) {
    statusCell.setValue("Pendente");
  }

  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Pendente", "Confirmar", "Recusar"], true)
    .setAllowInvalid(false)
    .build();

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

  const statusColumn = 13; 
  const emailColumn = 4; 

  if (editedColumn === statusColumn) {
    const newStatus = range.getValue();
    const email = sheet.getRange(editedRow, emailColumn).getValue();
    const name = sheet.getRange(editedRow, 2).getValue(); 

    if (newStatus === "Pendente") {
      Logger.log(
        `Status permanece como "Pendente" para ${name}. Nenhuma ação necessária.`
      );
      return;
    }

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