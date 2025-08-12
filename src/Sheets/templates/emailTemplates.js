const EMAIL_TEMPLATES = {
  VISITA_CONFIRMADA: {
    subject: "Visita Confirmada",
    body6indented: ({ nome, dataVisita, horario }) => `
      Prezado(a) ${nome},

      Recebemos sua solicitação de visita ao Jardim Botânico da UFSM para o dia ${dataVisita}, ${horario}. A data e o horário estão confirmados.

      Informações importantes:

      ● Comunicação: Todo o contato oficial será feito exclusivamente pelo e-mail jardimbotanico@ufsm.br. Em caso de dúvidas, estamos disponíveis também pelo WhatsApp: (55) 99193-8183.
      ● Tolerância de Atraso: A tolerância máxima para atrasos é de 30 minutos. Após esse período, a visita guiada poderá ser comprometida e as atividades poderão ser readequadas.
      ● Cancelamentos ou Imprevistos: Se precisar cancelar ou alterar a visita, por favor, informe-nos o mais rápido possível através do e-mail jardimbotanicoufsm@ufsm.br ou via WhatsApp.

      Orientações para o Professor/Responsável:
      - Roupas e calçados confortáveis (calça e tênis).
      - Protetor solar e repelente.
      - Garrafinha de água.

      Durante a visita, não é permitido:
      - Andar distante do monitor de educação ambiental e do grupo.
      - Subir em plantas, arrancá-las ou danificá-las de qualquer forma.
      - Colher ou ingerir frutos sem a autorização do monitor.
      - Deixar lixo pelo caminho.
      - Fazer brincadeiras de cunho preconceituoso ou violento.

      Aguardamos sua visita e estamos à disposição para qualquer esclarecimento adicional.

      Atenciosamente,  
      Equipe do Jardim Botânico da UFSM`,
  },

  VISITA_RECUSADA: {
    subject: "Necessidade de Remarcação de Data",
    body6indented: ({ nome, dataVisita, horario }) => `
      Prezado(a) ${nome},

      Recebemos um pedido de visita ao Jardim Botânico no dia ${dataVisita}, ${horario}. No entanto, neste dia e horário já temos outra escola agendada.

      Sugerimos que escolha outra data observando o Calendário de Visitas que se encontra no site do Jardim Botânico em "Agendamento de Visita" e atentando para o prazo mínimo de 7 dias úteis de antecedência na solicitação.

      Link: https://www.ufsm.br/orgaos-suplementares/jardim-botanico/agendamentos

      Após escolher nova data, por favor preencha o formulário novamente.

      Agradecemos seu interesse e aguardamos vocês!

      Atenciosamente,  
      Equipe do Jardim Botânico da UFSM`,
  },

  LEMBRETE_VISITA: {
    subject: "Lembrete: Visita Agendada ao Jardim Botânico",
    body6indented: ({ nome, dataVisita, horario }) => `
      Prezado(a) ${nome},

      Lembramos que a sua visita está agendada para o dia ${dataVisita} às ${horario}. Nossa equipe está preparada para atendê-los!

      Caso seja necessário cancelar por algum imprevisto, comunique rapidamente pelo e-mail jardimbotanico@ufsm.br com o assunto “cancelamento”. Para remarcação, basta realizar nova solicitação pelo formulário de visitações.

      Atenciosamente,  
      Equipe do Jardim Botânico da UFSM`,
  },

  POS_VISITA: {
    subject: "Avaliação da Visitação ao JBSM",
    body6indented: () => `
      Olá,

      Agradecemos pela visita e esperamos que a experiência tenha sido positiva. Para melhorar o nosso atendimento, solicitamos o preenchimento da avaliação pelo formulário: https://docs.google.com/forms/d/1mivpbSZA7R42DwLxIXZHtsF3W8vxpA1XvH6rVpQokvM/edit. São só 5 minutinhos!

      Obrigada!  
      Equipe do Jardim Botânico da UFSM`,
  },
};

const SEND_TEST_EMAIL = false;
const DEFAULT_EMAIL_RECIPIENT = "thalisson12sza@gmail.com";

const EMAIL_TEST_CASES = {
  VISITA_CONFIRMADA: {
    recipient: DEFAULT_EMAIL_RECIPIENT,
    template: EMAIL_TEMPLATES.VISITA_CONFIRMADA,
    params: {
      nome: "Joana",
      dataVisita: "12/06/2025",
      horario: "Manhã",
    },
  },
  VISITA_RECUSADA: {
    recipient: DEFAULT_EMAIL_RECIPIENT,
    template: EMAIL_TEMPLATES.VISITA_RECUSADA,
    params: {
      nome: "Joana",
      dataVisita: "13/06/2025",
      horario: "Tarde",
    },
  },
  LEMBRETE_VISITA: {
    recipient: DEFAULT_EMAIL_RECIPIENT,
    template: EMAIL_TEMPLATES.LEMBRETE_VISITA,
    params: {
      nome: "Joana",
      dataVisita: "14/06/2025",
      horario: "Manhã",
    },
  },
  POS_VISITA: {
    recipient: DEFAULT_EMAIL_RECIPIENT,
    template: EMAIL_TEMPLATES.POS_VISITA,
    params: {},
  },
};

function fillEmailTemplate(template, params) {
  const subject = template.subject;
  const body = template.body6indented(params).replace(/^ {6}/gm, "");
  return { subject, body };
}

function testEmailTemplate(testCase) {
  const { recipient, template, params } = EMAIL_TEST_CASES[testCase];
  const { subject, body } = fillEmailTemplate(template, params);

  Logger.log("Sending test email...");
  Logger.log("To: " + recipient);
  Logger.log("Subject: " + subject);
  Logger.log("Body:\n" + body);

  if (SEND_TEST_EMAIL) {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
    });
  } else {
    Logger.log("🚫 Test email not sent. Set SEND_TEST_EMAIL = true to proceed.");
  }
}

function testAllTemplates() {
  Object.keys(EMAIL_TEST_CASES).forEach((testCase) => {
    Logger.log("Testing template: " + testCase);
    testEmailTemplate(testCase);
  });
}

function test_VISITA_CONFIRMADA() {
  testEmailTemplate("VISITA_CONFIRMADA");
}
