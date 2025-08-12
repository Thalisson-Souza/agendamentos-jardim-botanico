const CALENDAR_ID = "thalissongamer65@gmail.com";

function getCalendarService() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    throw new Error(`Calendário com ID ${CALENDAR_ID} não encontrado.`);
  } else {
    Logger.log("Calendário encontrado!");
  }
  return calendar;
}

/** Verifica a disponibilidade de um horário no calendário. */
function verificarDisponibilidade(startTime, endTime) {
  try {
    const calendar = getCalendarService();
    const scriptTimeZone = Session.getScriptTimeZone();
    const calendarTimeZone = calendar.getTimeZone();

    if (scriptTimeZone !== calendarTimeZone) {
      Logger.log(`Aviso: timezones diferentes - Script: ${scriptTimeZone}, Calendário: ${calendarTimeZone}`);
    }

    const events = calendar.getEvents(startTime, endTime);
    Logger.log(
      `Eventos encontrados entre ${Utilities.formatDate(
        startTime,
        calendarTimeZone,
        "dd/MM/yyyy HH:mm:ss"
      )} e ${Utilities.formatDate(endTime, calendarTimeZone, "dd/MM/yyyy HH:mm:ss")}: ${events.length}`
    );

    return events.length === 0;
  } catch (error) {
    Logger.log("Erro na verificação de disponibilidade: " + error.toString());
    return false;
  }
}

function inserirEventoAgenda({ name, email, startTime, endTime, description }) {
  try {
    const calendar = getCalendarService();
    const calendarTimeZone = calendar.getTimeZone();

    Logger.log("startTime (tipo): " + typeof startTime + " | " + startTime);
    Logger.log("endTime (tipo): " + typeof endTime + " | " + endTime);

    if (!(startTime instanceof Date) || isNaN(startTime.getTime())) {
      throw new Error("startTime inválido!");
    }
    if (!(endTime instanceof Date) || isNaN(endTime.getTime())) {
      throw new Error("endTime inválido!");
    }

    calendar.createEvent(`Visita agendada - ${name}`, startTime, endTime, {
      description: description,
      guests: email,
      timeZone: calendarTimeZone,
    });
    Logger.log(
      `Evento criado para ${name} em ${Utilities.formatDate(startTime, calendarTimeZone, "dd/MM/yyyy HH:mm:ss")}`
    );
  } catch (error) {
    throw new Error("Erro ao inserir evento na agenda: " + error.toString());
  }
}
