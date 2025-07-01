/**
 * Função principal acionada pelo envio do formulário.
 * @param {GoogleAppsScript.Events.Forms.FormResponse} e */
function onFormSubmit(e) {
  try {
    const formData = FormParser.parseFormResponse(e); 
    Logger.log("Dados do formulário processados: " + JSON.stringify(formData));

    const result = BusinessLogic.processarSolicitacaoVisita(formData); 
    Logger.log("Resultado da lógica de negócio: " + JSON.stringify(result.status));

    switch (result.status) {
      case 'CONFIRMED':
        CalendarService.inserirEventoAgenda(result.calendarEventDetails); 
        EmailService.enviarEmail(result.emailDetails); 
        Logger.log("Visita agendada com sucesso e e-mail de confirmação enviado.");
        break;
      case 'REJECTED_CONFLICT':
        EmailService.enviarEmail(result.emailDetails);
        Logger.log("Visita recusada por conflito de horário e e-mail de recusa enviado.");
        break;
      case 'INCOMPLETE_DATA':
        Logger.log("Dados incompletos, nenhuma ação tomada.");
        break;
      default:
        Logger.log("Status de resultado desconhecido.");
    }

  } catch (error) {
    Logger.log("Erro geral no processamento do formulário: " + error.toString());
  }
}
