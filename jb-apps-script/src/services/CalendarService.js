/**   Obtém o serviço de calendário pelo ID. */
function getCalendarService() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    throw new Error(`Calendário com ID ${CALENDAR_ID} não encontrado.`);
  }
  return calendar;
}

function verificarDisponibilidade(startTime, endTime) {
  try {
    const calendar = getCalendarService();
    const scriptTimeZone = Session.getScriptTimeZone();
    const calendarTimeZone = calendar.getTimeZone();

    if (scriptTimeZone !== calendarTimeZone) {
      Logger.log(`Aviso: timezones diferentes - Script: ${scriptTimeZone}, Calendário: ${calendarTimeZone}`);
    }

    const events = calendar.getEvents(startTime, endTime);
    Logger.log(`Eventos encontrados entre ${Utilities.formatDate(startTime, calendarTimeZone, "dd/MM/yyyy HH:mm:ss")} e ${Utilities.formatDate(endTime, calendarTimeZone, "dd/MM/yyyy HH:mm:ss")}: ${events.length}`);

    return events.length === 0;
  } catch (error) {
    Logger.log("Erro na verificação de disponibilidade: " + error.toString());
    return false;
  }
}

function inserirEventoAgenda(eventDetails) {
  const { name, email, startTime, endTime, description } = eventDetails;
  try {
    const calendar = getCalendarService();
    const calendarTimeZone = calendar.getTimeZone();

    calendar.createEvent(
      `Visita agendada - ${name}`,
      startTime,
      endTime,
      {
        description: description,
        guests: email,
        timeZone: calendarTimeZone,
      }
    );
    Logger.log(`Evento criado para ${name} em ${Utilities.formatDate(startTime, calendarTimeZone, "dd/MM/yyyy HH:mm:ss")}`);
  } catch (error) {
    throw new Error("Erro ao inserir evento na agenda: " + error.toString());
  }
}
