/** Envia uma notificação por e-mail usando um template predefinido. */
function enviarEmail(emailDetails) {
  const { template, recipientEmail, params } = emailDetails;
  try {
    if (!template) {
      throw new Error(`Template de e-mail não fornecido ou inválido.`);
    }

    const { subject, body } = fillEmailTemplate(template, params);

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: body.trim(),
    });

    Logger.log(`E-mail enviado para ${recipientEmail} usando o template (objeto)`);
  } catch (error) {
    Logger.log(`Erro ao enviar e-mail: ${error}`);
  }
}
