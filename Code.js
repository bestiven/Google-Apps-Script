function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tanque Volumen')
    .addItem('Calcular Volumen', 'mostrarBarraLateral')
    .addToUi();
}

function mostrarBarraLateral() {
  var html = HtmlService.createHtmlOutputFromFile('Formulario')
      .setTitle('Calcular Volumen')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function calcularVolumen(tipo, altura) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName;
  var sheet;
  var range;
  var values;
  var rowFound = -1; // Usar -1 para indicar que no se encontró
  var minDifference = Infinity; // Inicializar con un valor grande
  var closestValue;
  var closestRow;
  var adjacentValue;

  // Determinar el nombre de la hoja basado en el tipo y verificar el límite de altura
  if (tipo == 1) {
    sheetName = 'Tanque1';
    if (altura > 2165) {
      return 'La altura ingresada para Tanque 1 es mayor a 2165 y no está dentro de los valores permitidos.';
    }
  } else if (tipo == 2) {
    sheetName = 'Tanque2';
    if (altura > 1505) {
      return 'La altura ingresada para Tanque 2 es mayor a 1505 y no está dentro de los valores permitidos.';
    }
  } else if (tipo == 3) {
    sheetName = 'Tanque3';
    if (altura > 2130) {
      return 'La altura ingresada para Tanque 3 es mayor a 2130 y no está dentro de los valores permitidos.';
    }
  } else {
    return 'Tipo de tanque no válido';
  }

  // Obtener la hoja correspondiente
  sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return 'Hoja no encontrada: ' + sheetName;
  }

  // Definir el rango a buscar (fila 13, columna 1 y 2)
  range = sheet.getRange(13, 1, sheet.getLastRow() - 12, 2); // Obtener desde la fila 13 hacia abajo en las columnas 1 y 2
  values = range.getValues();

  // Buscar el valor exacto o el más cercano
  for (var i = 0; i < values.length; i++) {
    var currentValue = values[i][0];
    var currentAdjacentValue = values[i][1];
    var difference = Math.abs(currentValue - altura);
    
    if (currentValue == altura) {
      rowFound = i + 13; // Fila en la hoja de cálculo
      return { row: rowFound, adjacentValue: currentAdjacentValue }; // Retorna la fila y el valor adyacente
    }
    
    // Actualizar el valor más cercano
    if (difference < minDifference) {
      minDifference = difference;
      closestValue = currentValue;
      closestRow = i + 13; // Fila en la hoja de cálculo
      adjacentValue = currentAdjacentValue; // Valor en la columna adyacente
    }
  }

  if (closestRow !== undefined) {
    return { row: closestRow, adjacentValue: adjacentValue }; // Retorna la fila y el valor adyacente del valor más cercano
  } else {
    return 'No se encontraron valores';
  }
}

function obtenerDatosTanque(tipo, altura) {
  var result = calcularVolumen(tipo, altura);
  
  if (typeof result === 'string') {
    return result; // Retorna el mensaje de error o no encontrado
  }
  
  var mensaje = '';

  if (result.row !== undefined) {
    mensaje = + result.adjacentValue + ' Galones';
  } else {
    mensaje = 'No se encontró el volumen para el tanque ' + tipo + ' con altura ' + altura + ' mm.';
  }
  
  return mensaje;
}







