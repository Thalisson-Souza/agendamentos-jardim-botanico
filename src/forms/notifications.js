function enviarAviso48Horas() {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const limiteInferior = new Date(now.getTime() + 48 * 60 * 60 * 1000); // Daqui a 48 horas
    const limiteSuperior = new Date(limiteInferior.getTime() + 1 * 60 * 60 * 1000); // Intervalo de 1 hora após 48 horas

    Logger.log("Buscando eventos entre " + limiteInferior + " e " + limiteSuperior);

    const events = calendar.getEvents(limiteInferior, limiteSuperior);

    if (events.length === 0) {
      Logger.log("Nenhum evento encontrado para aviso.");
      return;
    }

    events.forEach(event => {
        const guests = event.getGuestList();
      if (guests.length === 0) {
        Logger.log(`O evento "${event.getTitle()}" não possui convidados.`);
        return;
      }

      guests.forEach(guest => {
        const email = guest.getEmail();
        const attendingStatus = guest.getGuestStatus();

        if (attendingStatus === CalendarApp.GuestStatus.NO) {
          Logger.log(`O convidado ${email} recusou o evento "${event.getTitle()}".`);
          return;
        }

        MailApp.sendEmail({
          to: email,
          subject: `Aviso: Evento "${event.getTitle()}" em 48 horas`,
          body: `Olá,

Gostaríamos de lembrar que o evento "${event.getTitle()}" está programado para daqui a 48 horas.

📅 Data do evento: ${event.getStartTime().toLocaleDateString("pt-BR")}
🕒 Horário do evento: ${event.getStartTime().toLocaleTimeString("pt-BR", { hour: '2-digit', minute: '2-digit' })}

Atenciosamente,
Equipe do Jardim Botânico`
        });

        Logger.log(`Email de aviso enviado para: ${email} sobre o evento "${event.getTitle()}".`);
      });
    });

  } catch (error) {
    Logger.log("Erro na função de aviso: " + error.toString());
  }
}

function enviarEmailFeedback() {
  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  
  const limiteInferior = new Date(now.getTime() - 2 * 60 * 60 * 1000); // 2hr atrás
  const limiteSuperior = now;
  
  const events = calendar.getEvents(limiteInferior, limiteSuperior);
  
  events.forEach(event => {
    if (event.getEndTime() <= now) {
      const guests = event.getGuestList();
      guests.forEach(guest => {
        const email = guest.getEmail();
        
        MailApp.sendEmail({
          to: email,
          subject: `Queremos ouvir você! Feedback sobre o evento "${event.getTitle()}"`,
          body: `Olá,

Obrigado por participar do evento "${event.getTitle()}" no Jardim Botânico.

Gostaríamos de saber como foi sua experiência. Por favor, responda ao formulário de feedback no link abaixo:

https://forms.gle/fH98Y4R1vkCCVyKf7

Agradecemos sua colaboração!

Atenciosamente,
Equipe do Jardim Botânico`
        });
        Logger.log(`Email de feedback enviado para ${email} sobre o evento "${event.getTitle()}".`);
      });
    }
  });
}