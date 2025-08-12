function logIDs() {
  const form = FormApp.openByUrl("https://docs.google.com/forms/d/1K_rq5fhjAoBcrrUYnvbpqrwiOs-VWfJROKMIgio4huM/edit");

  const items = form.getItems();
  Logger.log("IDs dos Itens do Formulário:");
  items.forEach((item) => {
    Logger.log(`Título: "${item.getTitle()}" | ID: "${item.getId()}"`);
  });
}
