function sidebarFeedback() {
  const generator = HtmlService.createHtmlOutputFromFile('feedback')
    .setTitle('Banco de preguntas')
    .setWidth(550);
  SpreadsheetApp.getUi().showSidebar(generator);
}

export default sidebarFeedback;
