/** Processa uma solicitação de visita, verificando disponibilidade e determinando o resultado.*/
function processarSolicitacaoVisita(formData) {
  const { email, name, visitDate, visitTime } = formData;

  if (!email || !name || !visitDate || !visitTime) {
    Logger.log("Dados do formulário incompletos. Não é possível prosseguir.");
    return { status: 'INCOMPLETE_DATA', emailRecipient: email };
  }

  const parsedDate = parseDate(visitDate);
  const { startTime, endTime } = ajustarHorario(parsedDate, visitTime);

  const isAvailable = verificarDisponibilidade(startTime, endTime);
  Logger.log(`Disponibilidade: ${isAvailable}`);

  if (!isAvailable) {
    return {
      status: 'REJECTED_CONFLICT',
      emailDetails: {
        template: EMAIL_TEMPLATES.VISITA_RECUSADA,
        recipientEmail: email,
        params: { nome: name, dataVisita: visitDate, horario: visitTime }
      }
    };
  } else {
    const description = `Visitante: ${name}\nE-mail: ${email}\nHorário: ${visitTime}`;
    return {
      status: 'CONFIRMED',
      calendarEventDetails: {
        name: name,
        email: email,
        startTime: startTime,
        endTime: endTime,
        description: description
      },
      emailDetails: {
        template: EMAIL_TEMPLATES.VISITA_CONFIRMADA,
        recipientEmail: email,
        params: { nome: name, dataVisita: visitDate, horario: visitTime }
      }
    };
  }
}
