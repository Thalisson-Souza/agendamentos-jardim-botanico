const FIELD_ID_MAP = {
  name: 1318079111,
  visitDate: 967253933,
  visitTime: 1008086976,
};

const FormParser = {
  parseFormResponse: function (e) {
    const responses = e.response.getItemResponses();
    const formData = {
      name: null,
      visitDate: null,
      visitTime: null,
      email: e.response.getRespondentEmail(),
    };

    responses.forEach(function (itemResponse) {
      const itemId = itemResponse.getItem().getId();
      const answer = itemResponse.getResponse();

      if (itemId === FIELD_ID_MAP.name) formData.name = answer;
      else if (itemId === FIELD_ID_MAP.visitDate) formData.visitDate = answer;
      else if (itemId === FIELD_ID_MAP.visitTime) formData.visitTime = answer;
    });

    return formData;
  },
};
