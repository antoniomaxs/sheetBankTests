function getNumLastActiveRow(dataRows) {
  if (dataRows.isBlank()) return -1;
  let lastIndex = -1;
  dataRows.getDisplayValues().forEach(function (row, index) {
    if (row[0] !== '') {
      lastIndex = index;
    }
  });
  return lastIndex + 1;
}

function addItemOptions(item, respuestas, soluciones) {
  const choices = [];
  let respuestaCorrecta = false;
  for (let k = 0; k < respuestas.length; k += 1) {
    respuestaCorrecta = false;
    for (let j = 0; j < soluciones.length; j += 1) {
      if (respuestas[k].toString().trim() === soluciones[j].toString().trim()) {
        respuestaCorrecta = true;
      }
    }
    choices.push(item.createChoice(respuestas[k], respuestaCorrecta));
  }
  item.setChoices(choices);
  return item;
}

function addItemsToForm(values, form, selectedItemsRownum, separator) {
  let linea;
  let tipoPregunta;
  let titulo;
  let descripcion;
  let url;
  let puntuacion;
  // let nivel;
  // let obligatoria;
  // let llave;
  let strRespuestas;
  let respuestas = [];
  let strSoluciones;
  let soluciones = [];
  let item;
  for (let i = 0; i < selectedItemsRownum.length; i += 1) {
    linea = selectedItemsRownum[i];
    if (values[linea] !== 'undefined') {
      tipoPregunta = values[linea][0];
      titulo = values[linea][1];
      Logger.log(`Cargando datos de la línea: ${linea} con titulo ${titulo}`);
      descripcion = values[linea][2];
      url = values[linea][3];
      puntuacion = values[linea][4];
      // nivel = values[linea][5];
      // activa = values[linea][6];
      // llave = values[linea][7];
      strRespuestas = values[linea][8];
      Logger.log(`strRespuestas: ${strRespuestas}`);
      respuestas = strRespuestas.split(separator);
      strSoluciones = values[linea][9];
      Logger.log(`strSoluciones: ${strSoluciones}`);
      soluciones = strSoluciones.split(separator);

      Logger.log(`tipo de pregunta: ${tipoPregunta}`);
      if (url !== '') {
        Logger.log(`detectada URL: ${url}`);
        if (url.indexOf('[VIDEO]') >= 0) {
          Logger.log(`detectado VIDEO`);
          item = form.addVideoItem().setTitle(`[Vídeo] ${titulo}`);
          item.setVideoUrl(url.substring(7, url.length).trim());
        } else if (url.indexOf('[IMAGEN]') >= 0) {
          Logger.log(`detectada IMAGEN`);
          const img = DriveApp.getFileById(url.substring(8, url.length).trim());
          item = form.addImageItem().setTitle(`[Imagen] ${titulo}`);
          item.setImage(img);
        }
      }

      switch (tipoPregunta) {
        case 'RESP_MULT':
          item = form.addMultipleChoiceItem().setTitle(titulo).setHelpText(descripcion);
          item = addItemOptions(item, respuestas, soluciones);
          // if(url !== '') item.setVideoUrl(url);
          break;
        case 'SELEC_MULT':
          item = form.addCheckboxItem().setTitle(titulo).setHelpText(descripcion);
          item = addItemOptions(item, respuestas, soluciones);
          // if(url !== '') item.setVideoUrl(url);
          break;
        case 'VERD_FALSO':
          item = form.addMultipleChoiceItem().setTitle(titulo).setHelpText(descripcion);
          item = addItemOptions(item, respuestas, soluciones);
          break;
        case 'TEXTO_CORTO':
          item = form.addTextItem().setTitle(titulo).setHelpText(descripcion);
          break;
        case 'TEXTO_LARGO':
          item = form.addParagraphTextItem().setTitle(titulo).setHelpText(descripcion);
          break;
        case 'HORA':
          item = form.addTimeItem().setTitle(titulo).setHelpText(descripcion);
          break;
        case 'FECHA':
          item = form.addDateItem().setTitle(titulo).setHelpText(descripcion);
          break;
        case 'ESCALA':
          item = form.addTimeItem().setTitle(titulo).setHelpText(descripcion);
          break;
        /* case 'VIDEO':
          item = form.addVideoItem().setTitle(titulo).setHelpText(descripcion);
          item.setVideoUrl(url);
          break; */
        default:
          item = null;
          SpreadsheetApp.getUi().alert(
            'Algún tipo de pregunta no es correcta. Revísela, por favor'
          );
          break;
      }

      item.setRequired(true);
      item.setPoints(puntuacion);
    } else {
      Logger.log(`undefined en : ${values[linea]}`);
    }
  }
}

function getLevelFilterItems(numq, numFilas, values, levelsq, alreadyIncluded) {
  const items = [];
  let level;
  for (let i = 0; i < numFilas; i += 1) {
    level = values[i][5];
    for (let j = 0; j < levelsq.length; j += 1) {
      if (level === levelsq[j] && alreadyIncluded.indexOf(j) === -1) {
        Logger.log(
          `Detectada pregunta ${i} con un nivel de los del filtro y no está ya en el test.`
        );
        items.push(i);
        break;
      }
    }
    if (items.length === numq) break;
  }
  return items;
}

function getAleaItems(numq, numFilas, alreadyIncluded) {
  const items = [];
  let tmpAlea;

  // Calculos los números aleatorios para las preguntas
  while (items.length < numq) {
    tmpAlea = Math.round(Math.random() * (numFilas - 1));
    if (
      items.indexOf(tmpAlea) === -1 &&
      (alreadyIncluded === undefined ||
        (alreadyIncluded !== undefined && alreadyIncluded.indexOf(tmpAlea) === -1))
    ) {
      items.push(tmpAlea);
    }
  }
  Logger.log(`Items encontrados ${items}`);
  return items;
}

function getKeyItems(numq, numFilas, values, levelsq) {
  const items = [];
  let isLlave = false;
  let level;
  for (let i = 0; i < numFilas; i += 1) {
    // if (values[i] !== undefined) {
    isLlave = values[i][7];
    if (isLlave === 'SI') {
      if (levelsq === '') {
        Logger.log(
          `Detectada pregunta llave en la posición ${i} (filtrado de niveles desactivado)`
        );
        items.push(i);
      } else {
        level = values[i][5];
        for (let j = 0; j < levelsq.length; j += 1) {
          if (level === levelsq[j]) {
            Logger.log(
              `Detectada pregunta ${i} con un nivel de los del filtro y no está ya en el test.`
            );
            items.push(i);
          }
        }
      }
    }
  }
  if (items.length > numq) {
    const aleaItems = [];
    let tmpAlea;
    while (aleaItems.length < numq) {
      tmpAlea = Math.round(Math.random() * (items.length - 1));
      if (!aleaItems.includes(items[tmpAlea])) {
        aleaItems.push(items[tmpAlea]);
      }
    }
    Logger.log(`KeyItems aleatorios encontrados ${aleaItems}`);
    return aleaItems;
  }
  Logger.log(`KeyItems encontrados ${items}`);
  return items;
}

function calculateItems(numq, numFilas, keysq, levelsq, values) {
  let items = [];
  Logger.log(`Valores introducidos: keysq:${keysq}, levelsq:${levelsq}, numq:${numq}`);

  if (keysq !== '' && keysq === true) {
    Logger.log('Se ha indicado que se deben de incluir todas las preguntas llave');
    items = getKeyItems(numq, numFilas, values, levelsq);
  }

  // if ((keysq === '' || keysq === 'false') || items.length < numq) {
  if (items.length < numq) {
    Logger.log(
      `El sistema ha podido añadir ${items.length} preguntas llave o no se ha marcado la opción y el total son ${numq}`
    );
    if (levelsq === '') {
      const aleaItems = getAleaItems(numq - items.length, numFilas, items);
      Logger.log(`Se han calculado ${aleaItems.length} preguntas aleatorias.`);
      items = items.concat(aleaItems);
    } else {
      Logger.log(
        `Se ha activado el filtrado de preguntas por niveles, niveles del filtro ${levelsq.length}`
      );
      const filterItems = getLevelFilterItems(
        numq - items.length,
        numFilas,
        values,
        levelsq,
        items
      );
      Logger.log(
        `Se han calculado ${filterItems.length} preguntas filtradas por niveles.`
      );
      items = items.concat(filterItems);
    }
  } else {
    Logger.log('El test ha sido completado con preguntas de tipo llave por completo.');
  }

  Logger.log(`Nº Filas seleccionadas: ${items.length}`);
  return items;
}

const createForm = (data) => {
  const { keysq, levelsq, numq } = data;
  Logger.log(`Valores introducidos: keysq:${keysq}, levelsq:${levelsq}, numq:${numq}`);
  const spreadsheet = SpreadsheetApp.getActiveSheet();
  const headValues = spreadsheet.getRange('A1:J3').getValues();
  let numFilas;

  const lastDataRow = headValues[0][1];
  if (numFilas !== '') {
    numFilas = getNumLastActiveRow(spreadsheet.getRange(`B5:B${lastDataRow + 5}`));
  } else {
    numFilas = getNumLastActiveRow(spreadsheet.getRange('B5:B105'));
  }
  const firstDataRow = numFilas + 4;
  Logger.log(
    `Nº de filas de la hoja: ${numFilas} y última fila de datos: ${firstDataRow} por lo tanto los datos están comprendidos en el rango A5:J${firstDataRow}`
  );
  const values = spreadsheet.getRange(`A5:J${firstDataRow}`).getValues();

  let dataValidated = false;
  if (numq === '') {
    SpreadsheetApp.getUi().alert(
      'El número de preguntas del exámen es un valor obligatorio.'
    );
  } else if (numq > numFilas) {
    SpreadsheetApp.getUi().alert(
      'El número de preguntas no puede ser mayor que el númeo de preguntas que existen.'
    );
  } else {
    Logger.log(`Datos iniciales válidos. creando un exámen de ${numq} preguntas...`);
    dataValidated = true;
  }

  const formTitle = headValues[1][1];
  Logger.log(`Título del formulario: ${formTitle}`);
  const formDescription = headValues[2][1];
  Logger.log(`Descripción del formulario: ${formDescription}`);

  if (formTitle === '' || formDescription === '') {
    SpreadsheetApp.getUi().alert(
      'Los datos de título y descripción del formulario son obligatorios. Revise los datos en la hoja'
    );
    dataValidated = false;
  }

  if (dataValidated) {
    // Calculos los items según los datos introducidos.
    const selectedItemsRownum = calculateItems(numq, numFilas, keysq, levelsq, values);

    if (selectedItemsRownum.length !== parseInt(numq, 10)) {
      Logger.log(
        `Se han podido recuperar  ${selectedItemsRownum.length} preguntas y el total necesario son ${numq}`
      );
      SpreadsheetApp.getUi().alert(
        'Con los filtros establecidos no se han podido obtener el número de preguntas indicado, repase las preguntas o modifique los filtros'
      );
    } else {
      // Leemos las opciones del formulario
      const showProgress = headValues[1][3];
      const limitOneAnswer = headValues[1][5];
      // const publishMarks = headValues[1][7];
      const confirmationMessage = headValues[1][9];

      const collectEmails = headValues[2][3];
      const alatorizeQuestions = headValues[2][5];
      const showAnswers = headValues[2][7];
      const optionSeparator = headValues[2][9];

      // Creamos el form y aniadimos las preguntas
      const form = FormApp.create(formTitle)
        .setTitle(formTitle)
        .setDescription(formDescription)
        .setIsQuiz(true)
        .setAllowResponseEdits(false)
        // .setAcceptingResponses(true)
        .setCollectEmail(collectEmails === 'SI')
        .setConfirmationMessage(confirmationMessage)
        .setLimitOneResponsePerUser(limitOneAnswer === 'SI')
        .setPublishingSummary(showAnswers === 'SI')
        .setShuffleQuestions(alatorizeQuestions === 'SI')
        .setProgressBar(showProgress === 'SI');

      addItemsToForm(values, form, selectedItemsRownum, optionSeparator);

      Logger.log(`El formulario ha sido creado correctamente con el id ${form.getId()}`);
      spreadsheet.getRange('D1').activate();
      spreadsheet
        .getCurrentCell()
        .setValue(`https://docs.google.com/forms/d/${form.getId()}`);
    }
  } else {
    Logger.log('Los datos no han podido ser validados.');
  }
};

export default createForm;
