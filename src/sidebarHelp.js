function sidebarHelp() {
  const generator = HtmlService.createHtmlOutputFromFile('help')
    .setTitle('Banco de preguntas')
    .setWidth(550);
  SpreadsheetApp.getUi().showSidebar(generator);
}

export default sidebarHelp;
