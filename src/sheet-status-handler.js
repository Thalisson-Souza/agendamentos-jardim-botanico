function onFormSubmit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();
  
  // Defina a coluna da qual você quer adicionar o Status (exemplo: coluna F)
  var statusColumn = 6; // Coluna F
  
  // Preenche o status automaticamente como "Pendente"
  var statusCell = sheet.getRange(row, statusColumn);
  
  // Verifica se a célula da coluna Status está vazia e preenche com "Pendente"
  if (!statusCell.getValue()) {
    statusCell.setValue("Pendente");
  }

  // Adiciona a validação de dados na coluna Status, se ainda não existir
  var validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendente', 'Confirmado', 'Recusado'], true)
    .setAllowInvalid(false)
    .build();
  
  // Aplica a validação de dados para toda a coluna Status (exceto o cabeçalho)
  var lastRow = sheet.getLastRow();
  sheet.getRange(2, statusColumn, lastRow - 1, 1).setDataValidation(validation);
}

function createTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}
