function sidebarExam() {
  const generator = HtmlService.createHtmlOutputFromFile('generador')
    .setTitle('Banco de preguntas')
    .setWidth(550);
  SpreadsheetApp.getUi().showSidebar(generator);
}

export default sidebarExam;
