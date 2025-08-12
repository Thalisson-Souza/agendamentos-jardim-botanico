function parseDate(dateStr) {
  const [year, month, day] = dateStr.split("-");
  return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
}

function ajustarHorario(date, timeSlot) {
  const startTime = new Date(date);
  const endTime = new Date(date);

  switch (timeSlot) {
    case "A - 8h30 às 10:00":
      startTime.setHours(8, 30);
      endTime.setHours(10, 0);
      break;
    case "B - 10:30 às 12:00":
      startTime.setHours(10, 30);
      endTime.setHours(12, 0);
      break;
    case "C - 13h30 às 15:00":
      startTime.setHours(13, 30);
      endTime.setHours(15, 0);
      break;
    case "D - 15:30 às 17:00":
      startTime.setHours(15, 30);
      endTime.setHours(17, 0);
      break;
    default:
      throw new Error("Horário inválido: " + timeSlot);
  }

  return { startTime, endTime };
}
