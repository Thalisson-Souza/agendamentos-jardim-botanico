function onFormSubmit(e) {
  const responses = e.response.getItemResponses();
  let email = '';
  let name = '';
  let dataVisita = '';
  let horario = '';

  // Extrai dados do formulário
  responses.forEach(function(response) {
    const title = response.getItem().getTitle();
    const answer = response.getResponse();

    if (title === 'E-mail de contato') {
      email = answer;
    }
    if (title === 'Nome completo') {
      name = answer;
    }
    if (title === 'Data preferida para a visita') {
      dataVisita = answer; // Formato DD/MM/AAAA
    }
    if (title === 'Horário preferido') {
      horario = answer; // "Manhã" ou "Tarde"
    }
  });

  try {
    // Insere o evento no Google Agenda
    inserirEventoAgenda(name, email, dataVisita, horario);

    // Envia e-mail de confirmação
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: 'Confirmação da visita ao Jardim',
        body: `Olá ${name},\n\nSua visita foi agendada para:\nData: ${dataVisita}\nHorário: ${horario}\n\nAtenciosamente,\nEquipe do Jardim`
      });
    }

  } catch (error) {
    Logger.log("Erro: " + error.toString());
    // Opcional: enviar e-mail de erro para você
    MailApp.sendEmail("seu-email@exemplo.com", "Erro no agendamento", error.toString());
  }
}

// Função para inserir evento no Google Agenda
function inserirEventoAgenda(name, email, dataVisita, horario) {
  try {
    // Converte a data para o formato do Calendar
    const dataConvertida = new Date(dataVisita.replace(/(\d{2})\/(\d{2})\/(\d{4})/, "$3-$2-$1"));

    // Define horários (Manhã: 9h-12h, Tarde: 14h-17h)
    const startTime = new Date(dataConvertida);
    const endTime = new Date(dataConvertida);

    if (horario === "Manhã") {
      startTime.setHours(9, 0, 0);
      endTime.setHours(12, 0, 0);
    } else {
      startTime.setHours(14, 0, 0);
      endTime.setHours(17, 0, 0);
    }

    // Cria o evento no Google Agenda
    const calendar = CalendarApp.getDefaultCalendar(); // Usa o calendário padrão
    calendar.createEvent(
      `Visita agendada - ${name}`, // Título do evento
      startTime,
      endTime,
      {
        description: `Visitante: ${name}\nE-mail: ${email}\nHorário: ${horario}`,
        guests: email // Opcional: convida o e-mail do visitante
      }
    );

    Logger.log(`Evento criado para ${name} em ${dataVisita}, ${horario}`);
  } catch (error) {
    throw new Error("Erro ao inserir evento na agenda: " + error.toString());
  }
}
