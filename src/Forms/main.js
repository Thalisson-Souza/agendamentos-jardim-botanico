const SHEET_NAME = "Respostas ao formulário 1";

function onFormSubmit(e) {
  try {
    const formData = FormParser.parseFormResponse(e);

    const horarioSomente = formData.visitTime.split(" - ")[1] || formData.visitTime;

    const parsedDate = parseDate(formData.visitDate);
    const { startTime, endTime } = ajustarHorario(parsedDate, formData.visitTime);

    const isAvailable = verificarDisponibilidade(startTime, endTime);

    const destinationId = FormApp.getActiveForm().getDestinationId();
    const sheet = SpreadsheetApp.openById(destinationId).getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (!isAvailable) {
      enviarEmail({
        template: EMAIL_TEMPLATES.VISITA_RECUSADA,
        recipientEmail: formData.email,
        params: {
          nome: formData.name,
          dataVisita: formData.visitDate,
          horario: horarioSomente,
        },
      });
      sheet.deleteRow(lastRow);
      Logger.log(`Solicitação recusada: horário ocupado para ${formData.email}`);
      Logger.log(`id da planilha "${SHEET_NAME}": ${destinationId}.`);
      Logger.log(
        `OBS: Linha ${lastRow} apagada do sheets guia ${SHEET_NAME}, pois a inserção é feita automatica pelo forms.`
      );
    } else {
      Logger.log(
        `Solicitação recebida para ${formData.email} - horário livre inserindo no Sheets, para aguardar controle manual.`
      );
    }
  } catch (error) {
    Logger.log("Erro no processamento do formulário: " + error.toString());
  }
}

function testOnFormSubmit() {
  const e = {
    namedValues: {
      visitDate: ["2025-08-04"],
      visitTime: ["B - 10:30 às 12:00"],
      email: ["thalisson12sza@gmail.com"],
      name: ["Teste Responsável"],
    },
  };

  FormParser.parseFormResponse = function (event) {
    return {
      visitDate: event.namedValues.visitDate[0],
      visitTime: event.namedValues.visitTime[0],
      email: event.namedValues.email[0],
      name: event.namedValues.name[0],
    };
  };

  onFormSubmit(e);
}
