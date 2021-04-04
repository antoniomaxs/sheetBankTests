function addItemsToForm(form) {
  let item;
  // Pregunta 1
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Está realizando esta encuesta antes de conocer la nota del examen?')
    .setChoices([item.createChoice('Sí'), item.createChoice('No')])
    .setRequired(true);

  // Pregunta 2
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Cuál era su situación en la asignatura antes de realizar el examen?')
    .setChoices([
      item.createChoice('Mala'),
      item.createChoice('Regular'),
      item.createChoice('Buena'),
    ])
    .setRequired(true);

  // Pregunta 3
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Cómo cree que se ha desarrollado el examen?')
    .setChoices([
      item.createChoice('Muy mal'),
      item.createChoice('Mal'),
      item.createChoice('Bien'),
      item.createChoice('Muy bien'),
    ])
    .setRequired(true);

  // Pregunta 4
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Cree que las preguntas estaban correctamente expresadas?')
    .setChoices([
      item.createChoice('Sí, todas.'),
      item.createChoice('Sí, casi todas.'),
      item.createChoice('No, con algunas excepciones.'),
      item.createChoice('No, nignuna.'),
    ])
    .setRequired(true);

  // Pregunta 5
  item = form.addMultipleChoiceItem();
  item
    .setTitle(
      '¿Todas las preguntas estaban dentro del temario explicado por el profesor? En caso de responder negativamente puede indicar que preguntas le han parecido estar fuera del temario en la casilla final de observaciones.'
    )
    .setChoices([
      item.createChoice('Sí, todas.'),
      item.createChoice('Sí, casi todas.'),
      item.createChoice('No, con algunas excepciones.'),
      item.createChoice('No, ninguna.'),
    ])
    .setRequired(true);

  // Pregunta 6
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Cree que la dificultad de las preguntas ha sido adecuada?')
    .setChoices([
      item.createChoice('Sí, todas.'),
      item.createChoice('Sí, casi todas.'),
      item.createChoice('No, con algunas excepciones.'),
      item.createChoice('No, ninguna.'),
    ])
    .setRequired(true);

  // Pregunta 7
  item = form.addMultipleChoiceItem();
  item
    .setTitle(
      '¿Cree que los contenidos expuestos en la evaluación corresponden a los conceptos más importantes del tema?'
    )
    .setChoices([
      item.createChoice('Sí, todos.'),
      item.createChoice('Sí, casi todos.'),
      item.createChoice('No, con algunas excepciones.'),
      item.createChoice('No, ninguno.'),
    ])
    .setRequired(true);

  // Pregunta 8
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Cree que la ponderación de las preguntas ha sido adecuada?')
    .setChoices([item.createChoice('Sí'), item.createChoice('No')])
    .setRequired(true);

  // Pregunta 9
  item = form.addMultipleChoiceItem();
  item
    .setTitle('La dificultad general del examen es:')
    .setChoices([
      item.createChoice('Baja'),
      item.createChoice('Media'),
      item.createChoice('Alta'),
      item.createChoice('Muy alta'),
    ])
    .setRequired(true);

  // Pregunta 10
  item = form.addMultipleChoiceItem();
  item
    .setTitle('¿Cuál es su opinión sobre las evaluaciones mediante cuestionarios online?')
    .setChoices([
      item.createChoice('Muy buena'),
      item.createChoice('Buena'),
      item.createChoice('Regular'),
      item.createChoice('Mala'),
      item.createChoice('Muy mala'),
    ])
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle('Observaciones')
    .setHelpText(
      'En esta casilla podrá indicar cualquier observación que considere oportuna sobre la evaluación o su realización.'
    );
}

const createFeedback = () => {
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const headValues = spreadsheet.getRange('A1:J3').getValues();

  const formTitle = headValues[1][1];
  Logger.log(`Título del formulario: ${formTitle}`);
  const formDescription = headValues[2][1];
  Logger.log(`Descripción del formulario: ${formDescription}`);

  if (formTitle === '' || formDescription === '') {
    SpreadsheetApp.getUi().alert(
      'Los datos de título y descripción del formulario son obligatorios. Revise los datos en la hoja'
    );
  } else {
    const form = FormApp.create(`[Feedback] ${formTitle}`)
      .setTitle(`(Feedback) ${formTitle}`)
      .setDescription(`(Feedback) ${formDescription}`)
      .setIsQuiz(false)
      .setAllowResponseEdits(false)
      // .setAcceptingResponses(true)
      .setCollectEmail(false)
      .setConfirmationMessage('Gracias por sus respuetas!')
      .setLimitOneResponsePerUser(true)
      .setPublishingSummary(true)
      .setShuffleQuestions(false)
      .setProgressBar(true);

    addItemsToForm(form);

    Logger.log(
      `El formulario feedback ha sido creado correctamente con el id ${form.getId()}`
    );

    spreadsheet.getRange('F1').activate();
    spreadsheet
      .getCurrentCell()
      .setValue(`https://docs.google.com/forms/d/${form.getId()}`);
    // return form.getId().toString();
  }
};

export default createFeedback;
