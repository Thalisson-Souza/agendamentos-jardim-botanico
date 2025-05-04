function onFormSubmit(e) {
  const responses = e.response.getItemResponses();
  let email = "";
  let name = "";
  let dataVisita = "";
  let horario = "";

  // Extrai dados do formulário
  responses.forEach(function (response) {
    const title = response.getItem().getTitle();
    const answer = response.getResponse();

    if (title === "E-mail de contato") {
      email = answer;
    }
    if (title === "Nome completo responsável") {
      name = answer;
    }
    if (title === "Data preferida para a visita") {
      dataVisita = answer; // Formato DD/MM/AAAA
    }
    if (title === "Horário preferido") {
      horario = answer; // "Manhã" ou "Tarde"
    }
  });

  try {
    const disponivel = verificarDisponibilidade(dataVisita, horario);
    if (!disponivel) {
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "Conflito de Horário para Visita ao Jardim Botânico",
          body: `Olá ${name},
          
Infelizmente, a data e o horário solicitados para a visita ao Jardim Botânico já estão ocupados. Por favor, escolha outro horário e tente novamente.

Atenciosamente,  
Equipe do Jardim Botânico`,
        });

        Logger.log(`E-mail de conflito enviado para ${email}.`);
      }
    } else {
      inserirEventoAgenda(name, email, dataVisita, horario);

      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: "Solicitação de Visita Recebida",
          body: `Olá ${name},
    
Recebemos sua solicitação de visita ao Jardim Botânico para:

📅 Data: ${dataVisita}  
🕒 Período: ${horario.toLowerCase()}

Em breve, você receberá a confirmação de sua solicitação.

Atenciosamente,  
Equipe do Jardim Botânico`,
        });
        Logger.log(`E-mail de confirmação enviado para ${email}.`);
      }
    }
  } catch (error) {
    Logger.log("Erro: " + error.toString());
    // Opcional: enviar e-mail de erro para você
    MailApp.sendEmail(
      "seu-email@exemplo.com","Erro no agendamento", error.toString()
    );
  }
}

// Função para inserir evento no Google Agenda
function inserirEventoAgenda(name, email, dataVisita, horario) {
  try {
    // Converte a data para o formato do Calendar
    const dataConvertida = new Date(
      dataVisita.replace(/(\d{2})\/(\d{2})\/(\d{4})/, "$3-$2-$1")
    );

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
        guests: email, // Opcional: convida o e-mail do visitante
      }
    );

    Logger.log(`Evento criado para ${name} em ${dataVisita}, ${horario}`);
  } catch (error) {
    throw new Error("Erro ao inserir evento na agenda: " + error.toString());
  }
}

// Função para verificar se a data e o horário estão ocupados no calendário
function verificarDisponibilidade(dataVisita, horario) {
  try {
    // Converte a data para o formato do Calendar
    const dataConvertida = new Date(
      dataVisita.replace(/(\d{2})\/(\d{2})\/(\d{4})/, "$3-$2-$1")
    );

    // Define os horários (Manhã: 9h-12h, Tarde: 14h-17h)
    const startTime = new Date(dataConvertida);
    const endTime = new Date(dataConvertida);

    if (horario === "Manhã") {
      startTime.setHours(9, 0, 0);
      endTime.setHours(12, 0, 0);
    } else {
      startTime.setHours(14, 0, 0);
      endTime.setHours(17, 0, 0);
    }

    // Acessa o calendário e verifica se já existe um evento nesse horário
    const calendar = CalendarApp.getDefaultCalendar(); // Usa o calendário padrão
    const events = calendar.getEvents(startTime, endTime);

    if (events.length > 0) {
      // Caso haja eventos, retorna um erro ou mensagem de conflito
      return false;
    } else {
      // Caso não haja eventos, retorna verdadeiro
      return true;
    }
  } catch (error) {
    Logger.log("Erro na verificação de disponibilidade: " + error.toString());
    return false;
  }
}
