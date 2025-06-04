function onFormSubmit(e) {
  const responses = e.response.getItemResponses();
  let email = "";
  let name = "";
  let dataVisita = "";
  let horario = "";

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
      dataVisita = answer; //DD/MM/AAAA
    }
    if (title === "Horário preferido") {
      horario = answer; 
    }
  });

  try {
    const disponivel = verificarDisponibilidade(dataVisita, horario);
    if (!disponivel) {
      if (email) {
        enviarEmailConflito(email, name);
      }
    }else {
      inserirEventoAgenda(name, email, dataVisita, horario);

      if(email){
        enviarEmailSolicitacao(email, name, dataVisita, horario);
      }
    }
    }catch (error){
      Logger.log("Erro: " + error.toString());
    }
}

function enviarEmailConflito(email, name) {
  const subject = "Conflito de Horário para Visita ao Jardim Botânico";
  const body = `Olá ${name},
          
Infelizmente, a data e o horário solicitados para a visita ao Jardim Botânico já estão ocupados. Por favor, escolha outro horário e tente novamente.

Atenciosamente,  
Equipe do Jardim Botânico`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
  });

  Logger.log(`E-mail de conflito enviado para ${email}.`);
}

function enviarEmailSolicitacao(email, name, dataVisita, horario) {
  const subject = "Solicitação de Visita Recebida";
  const body = `Olá ${name},
    
Recebemos sua solicitação de visita ao Jardim Botânico para:

📅 Data: ${dataVisita}  
🕒 Período: ${horario.toLowerCase()}

Em breve, você receberá a confirmação de sua solicitação.

Atenciosamente,  
Equipe do Jardim Botânico`;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
  });

  Logger.log(`E-mail de confirmação enviado para ${email}.`);
}

function inserirEventoAgenda(name, email, dataVisita, horario) {
  try {
    const dataConvertida = new Date(dataVisita.replace(/(\d{2})\/(\d{2})\/(\d{4})/, "$3-$2-$1"));
    dataConvertida.setHours(dataConvertida.getHours() + 3);

    const startTime = new Date(dataConvertida);
    const endTime = new Date(dataConvertida);

    if (horario === "Manhã 8:30 - 10:00") {
      startTime.setHours(8, 30, 0);
      endTime.setHours(10, 0, 0);
    } else if(horario === "Manhã 10:30 - 12:00"){
      startTime.setHours(10, 30, 0);
      endTime.setHours(12, 0, 0);
    }else if(horario === "Tarde 13:30 - 15:00"){
      startTime.setHours(13, 30, 0);
      endTime.setHours(15, 0, 0);
    }else if(horario === "Tarde 15:30 - 17:00"){
      startTime.setHours(15, 30, 0);
      endTime.setHours(17, 0, 0);
    }

    const calendar = CalendarApp.getDefaultCalendar(); // Calendar padrão
    calendar.createEvent(
      `Visita agendada - ${name}`,
      startTime,
      endTime,
      {
        description: `Visitante: ${name}\nE-mail: ${email}\nHorário: ${horario}`,
        guests: email, 
        timeZone: "America/Sao_Paulo",
      }
    );
    Logger.log(`Evento criado para ${name} em ${dataVisita}, ${horario}`);
  } catch (error) {
    throw new Error("Erro ao inserir evento na agenda: " + error.toString());
  }
}

function verificarDisponibilidade(dataVisita, horario) {
  try {
    const dataConvertida = new Date(dataVisita.replace(/(\d{2})\/(\d{2})\/(\d{4})/, "$3-$2-$1"));
    dataConvertida.setHours(dataConvertida.getHours() + 3);

    const startTime = new Date(dataConvertida);
    const endTime = new Date(dataConvertida);

    if (horario === "Manhã 8:30 - 10:00") {
      startTime.setHours(8, 30, 0);
      endTime.setHours(10, 0, 0);
    } else if(horario === "Manhã 10:30 - 12:00"){
      startTime.setHours(10, 30, 0);
      endTime.setHours(12, 0, 0);
    }else if(horario === "Tarde 13:30 - 15:00"){
      startTime.setHours(13, 30, 0);
      endTime.setHours(15, 0, 0);
    }else if(horario === "Tarde 15:30 - 17:00"){
      startTime.setHours(15, 30, 0);
      endTime.setHours(17, 0, 0);
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startTime, endTime);

    if (events.length > 0) {
      return false;
    } else {
      return true;
    }
  } catch (error) {
    Logger.log("Erro na verificação de disponibilidade: " + error.toString());
    return false;
  }
}