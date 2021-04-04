const applyTemplate = (name) => {
  let nSheet;
  let spreadsheet;
  let sheet;
  let currentCell;
  const maxRows = 100;

  if (
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() === `#Bank#${name}`
  ) {
    SpreadsheetApp.getUi().alert(
      'No se puede crear otra plantilla con el mismo nombre. Cambie el nombre de la nueva hoja en el panel de Crear Plantilla.'
    );
  } else {
    nSheet = SpreadsheetApp.getActiveSpreadsheet()
      .insertSheet()
      .setName(`#Bank# ${name}`);
    spreadsheet = SpreadsheetApp.getActive();
    sheet = spreadsheet.getActiveSheet();
    sheet.getRange(1, 1, maxRows, sheet.getMaxColumns()).activate();
    spreadsheet.getActiveRangeList().setFontFamily('Roboto Condensed').setFontSize(11);

    // Bloqueo de cabecera
    spreadsheet.getRange('1:4').activate();
    spreadsheet.getActiveRangeList().setFontWeight('bold');
    spreadsheet.getActiveSheet().setFrozenRows(4);

    // Datos
    // Fila 1
    spreadsheet.getRange('A1').activate();
    spreadsheet.getCurrentCell().setValue('Nº Filas');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('B1').activate();
    spreadsheet.getCurrentCell().setValue('100');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    spreadsheet.getRange('C1').activate();
    spreadsheet.getCurrentCell().setValue('URL Form');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('E1').activate();
    spreadsheet.getCurrentCell().setValue('URL Feedback');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');

    // Datos
    // Fila 2
    spreadsheet.getRange('A2').activate();
    spreadsheet.getCurrentCell().setValue('Título');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('C2').activate();
    spreadsheet.getCurrentCell().setValue('Mostrar progreso');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('D2').activate();
    spreadsheet.getCurrentCell().setValue('NO');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    spreadsheet.getRange('E2').activate();
    spreadsheet.getCurrentCell().setValue('Limitar 1 respuesta');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getCurrentCell().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    spreadsheet.getRange('F2').activate();
    spreadsheet.getCurrentCell().setValue('SI');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    // spreadsheet.getRange('G2').activate();
    // spreadsheet.getCurrentCell().setValue('Publicar calificaciones');
    // spreadsheet.getCurrentCell().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    // spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    // spreadsheet.getRange('H2').activate();
    // spreadsheet.getCurrentCell().setValue('INMEDIATA');
    // spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    spreadsheet.getRange('I2').activate();
    spreadsheet.getCurrentCell().setValue('Mensaje confirmación');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('J2').activate();
    spreadsheet.getCurrentCell().setValue('Gracias por tus respuestas!');

    // Fila 3
    spreadsheet.getRange('A3').activate();
    spreadsheet.getCurrentCell().setValue('Descripción');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('C3').activate();
    spreadsheet.getCurrentCell().setValue('Recopilar emails');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('D3').activate();
    spreadsheet.getCurrentCell().setValue('NO');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    spreadsheet.getRange('E3').activate();
    spreadsheet.getCurrentCell().setValue('Aleatorizar preguntas');
    spreadsheet.getCurrentCell().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('F3').activate();
    spreadsheet.getCurrentCell().setValue('NO');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    spreadsheet.getRange('G3').activate();
    spreadsheet.getCurrentCell().setValue('Mostrar respuestas');
    spreadsheet.getCurrentCell().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('H3').activate();
    spreadsheet.getCurrentCell().setValue('NO');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');
    spreadsheet.getRange('I3').activate();
    spreadsheet.getCurrentCell().setValue('Separador opciones');
    spreadsheet.getCurrentCell().setHorizontalAlignment('right');
    spreadsheet.getRange('J3').activate();
    spreadsheet.getCurrentCell().setValue(';');
    spreadsheet.getCurrentCell().setHorizontalAlignment('center');

    // Fila 4
    spreadsheet.getRange('A4').activate();
    spreadsheet.getCurrentCell().setValue('Tipo');
    spreadsheet.getRange('B4').activate();
    spreadsheet.getCurrentCell().setValue('Título');
    spreadsheet.getRange('C4').activate();
    spreadsheet.getCurrentCell().setValue('Descripción');
    spreadsheet.getRange('D4').activate();
    spreadsheet.getCurrentCell().setValue('URL');
    spreadsheet.getRange('E4').activate();
    spreadsheet.getCurrentCell().setValue('Puntuación');
    spreadsheet.getRange('F4').activate();
    spreadsheet.getCurrentCell().setValue('Nivel');
    spreadsheet.getRange('G4').activate();
    spreadsheet.getCurrentCell().setValue('¿Activa?');
    spreadsheet.getRange('H4').activate();
    spreadsheet.getCurrentCell().setValue('¿LLave?');
    spreadsheet.getRange('I4').activate();
    spreadsheet.getCurrentCell().setValue('Respuestas');
    spreadsheet.getRange('J4').activate();
    spreadsheet.getCurrentCell().setValue('Soluciones');

    // Fondos
    // Fila 1, 2 y 3
    spreadsheet.getRange('A1:A3').activate();
    spreadsheet.getActiveRangeList().setBackground('#26a69a'); // verde(teal)
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');
    spreadsheet.getRange('C1:C3').activate();
    spreadsheet.getActiveRangeList().setBackground('#26a69a'); // verde(teal)
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');
    spreadsheet.getRange('E1:E3').activate();
    spreadsheet.getActiveRangeList().setBackground('#26a69a'); // verde(teal)
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');
    spreadsheet.getRange('G3').activate();
    spreadsheet.getActiveRangeList().setBackground('#26a69a'); // verde(teal)
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');
    spreadsheet.getRange('I2:I3').activate();
    spreadsheet.getActiveRangeList().setBackground('#26a69a'); // verde(teal)
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');

    // Fila 4
    spreadsheet.getRange('A4:F4').activate();
    spreadsheet.getActiveRangeList().setBackground('#9c27b0'); // deep-purple
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');
    spreadsheet.getRange('G4:H4').activate();
    spreadsheet.getActiveRangeList().setBackground('#F44336'); // red
    spreadsheet.getRange('I4').activate();
    spreadsheet.getActiveRangeList().setBackground('#9c27b0'); // deep-purple
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');
    spreadsheet.getRange('J4').activate();
    spreadsheet.getActiveRangeList().setBackground('#2979FF '); // blue
    spreadsheet.getActiveRangeList().setFontColor('#ffffff');

    // Tamanios
    spreadsheet.getActiveSheet().setColumnWidth(2, 260);
    spreadsheet.getActiveSheet().setColumnWidth(3, 193);
    spreadsheet.getActiveSheet().setColumnWidth(4, 138);
    spreadsheet.getActiveSheet().setColumnWidth(9, 300);
    spreadsheet.getActiveSheet().setColumnWidth(10, 300);

    // Validacion de datos (select)
    spreadsheet.getRange('A5').activate();
    spreadsheet
      .getRange('A5')
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInList(
            [
              'RESP_MULT',
              'SELEC_MULT',
              'VERD_FALSO',
              'TEXTO_CORTO',
              'TEXTO_LARGO',
              'HORA',
              'FECHA',
              'ESCALA',
              'VIDEO',
            ],
            true
          )
          .build()
      );
    spreadsheet.getRange(`A6:A${maxRows}`).activate();
    // currentCell = spreadsheet.getCurrentCell();
    // spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    // currentCell.activateAsCurrentCell();
    spreadsheet
      .getRange('A5')
      .copyTo(
        spreadsheet.getActiveRange(),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    spreadsheet.getRange('F5').activate();
    spreadsheet
      .getRange('F5')
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInList(
            ['Nivel 1(Bajo)', 'Nivel 2(Medio)', 'Nivel 3(Alto)', 'Nivel 4(Experto)'],
            true
          )
          .build()
      );
    spreadsheet.getRange(`F6:F${maxRows}`).activate();
    /* currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell(); */
    spreadsheet
      .getRange('F5')
      .copyTo(
        spreadsheet.getActiveRange(),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    spreadsheet.getRange('G5').activate();
    spreadsheet
      .getRange('G5')
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInList(['SI', 'NO'], true)
          .build()
      );
    spreadsheet.getRange(`G6:G${maxRows}`).activate();
    /* currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell(); */
    spreadsheet
      .getRange('G5')
      .copyTo(
        spreadsheet.getActiveRange(),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    spreadsheet.getRange('H5').activate();
    spreadsheet
      .getRange('H5')
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInList(['SI', 'NO'], true)
          .build()
      );
    spreadsheet.getRange(`H6:H${maxRows}`).activate();
    /* currentCell = spreadsheet.getCurrentCell();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell(); */
    spreadsheet
      .getRange('H5')
      .copyTo(
        spreadsheet.getActiveRange(),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    // Valores por defecto
    spreadsheet.getRange(`F5:F${maxRows}`).setValue('Nivel 2(Medio)');
    spreadsheet.getRange(`G5:G${maxRows}`).setValue('SI');
    spreadsheet.getRange(`H5:H${maxRows}`).setValue('NO');

    // Otros ajustes
    spreadsheet
      .getRange(`I5:F${maxRows}`)
      .setWrapStrategies(SpreadsheetApp.WrapStrategy.WRAP);
    spreadsheet
      .getRange(`J5:F${maxRows}`)
      .setWrapStrategies(SpreadsheetApp.WrapStrategy.WRAP);
  }
};

export default applyTemplate;
