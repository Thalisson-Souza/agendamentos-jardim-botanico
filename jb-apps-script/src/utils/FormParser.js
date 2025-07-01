/**  Extrai e estrutura os dados de um evento de envio de formulário,
 * @param {GoogleAppsScript.Events.Forms.FormResponse} e - 
 * @returns {Object} */
function parseFormResponse(e) {
  const responses = e.response.getItemResponses();
  const formData = {};

  for (const key in FIELD_ID_MAP) {
    formData[key] = null;
  }

  responses.forEach(function (itemResponse) {
    const itemId = itemResponse.getItem().getId(); 
    const answer = itemResponse.getResponse();     

    if (itemId === FIELD_ID_MAP.email) {
      formData.email = answer;
    } else if (itemId === FIELD_ID_MAP.name) {
      formData.name = answer;
    } else if (itemId === FIELD_ID_MAP.visitDate) {
      formData.visitDate = answer;
    } else if (itemId === FIELD_ID_MAP.visitTime) {
      formData.visitTime = answer;
    }
  });

  return formData;
}