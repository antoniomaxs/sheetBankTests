function sidebarTemplate() {
  const sidebar = HtmlService.createHtmlOutputFromFile('template')
    .setTitle('Banco de preguntas')
    .setWidth(550);
  SpreadsheetApp.getUi().showSidebar(sidebar);
}

export default sidebarTemplate;
