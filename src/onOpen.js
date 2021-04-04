const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu(' ⦿ Banco de preguntas')
    .addItem('1. Crear plantilla', 'sidebarTemplate')
    .addItem('2. Crear examen', 'sidebarExam')
    .addItem('3. Crear feedback', 'sidebarFeedback')
    .addSeparator()
    .addItem('+ Info', 'sidebarHelp')
    .addToUi();
};

export default onOpen;
