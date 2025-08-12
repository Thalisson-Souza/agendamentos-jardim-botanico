const CALENDAR_ID = "thalissongamer65@gmail.com";

function verificarDisponibilidade(startTime, endTime) {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) throw new Error("Calendário não encontrado.");

    const events = calendar.getEvents(startTime, endTime);
    return events.length === 0;
  } catch (error) {
    Logger.log("Erro ao verificar disponibilidade: " + error.toString());
    return false;
  }
}
