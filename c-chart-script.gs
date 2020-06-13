function analysis() {
  //Agregamos en una variable la hoja donde se van a realizar los análisis.
  var analysisSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AnalisisSheet");
  //Agregamos en una variable la hoja donde se encuentran los datos.
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  
  //Limpiamos las celdas, porque podría haber datos anteriores.
  clearData(analysisSheet);
  
  /* 
  Para el análisis de gráficas C, solo necesitamos contar la cantidad de quejas en un día, por lo que necesitamos tomar los datos
  de la fila de la fecha cuando ocurre la queja. A partir de esos datos se puede contar la cantidad de quejas en ese día.
  La fila que necesitamos es desde la D2 hasta la última fila de la última fecha ingresada (por eso se utiliza la función getLastRow()).
  
  NOTA: esta notación no es la más eficiente en términos de funcionabilidad porque getLastRow() devuelve el valor de la última fila de aquella columna
  que tenga la mayor cantidad de filas (porque esa sería la última de toda la hoja de cálculo; de todas las columnas).
  Este enfoque funciona únicamente porque la última fila va a ser la misma para todas las columnas de la hoja; es una matriz cuadrada.
  */
  
  //Se almacena el rango de la columna de los datos de las fechas de quejas.
  var datesColumnRange = "D2:D" + dataSheet.getLastRow();
  //Se almacena el rango de las columna de los datos del tipo de quejas.
  var complainColumnRange = "F2:D" + dataSheet.getLastRow();
  
  //Almacenamos en una variable un array con las fechas que se dieron las quejas.
  var data = dataSheet.getRange(datesColumnRange).getValues();
  //Almacenamos en una variable un array con los tipos de quejas que se dieron.
  var complains = dataSheet.getRange(complainColumnRange).getValues();
  
  /*
  El análisis se va a realizar según la fechas que quiera analizar el cliente, por lo que se va a analizar las fechas de un rango de fechas dado.
  */
  //Se almacena la fecha desde donde se quiere empezar el rango. El cliente lo debe ingresar desde la hoja de análisis.
  var dateFrom = new Date(Date.parse(analysisSheet.getRange(4, 2).getValue()));
  //Se almacena la fecha hasta donde se quiere empezar el rango. El cliente lo debe ingresar desde la hoja de análisis.
  var dateTo = new Date(Date.parse(analysisSheet.getRange(5, 2).getValue()));
  
  //Se verifica que el cliente haya ingresado el rango que quiere analizar.
  if (dateFrom && dateTo) {
    //El cliente sí ingresó el rango de fechas.
    
    //Ahora se almacena fechas que sí cumplen con el rango dado.
    var filteredDates = checkRange(data, complains, dateFrom, dateTo);
    
    //Verificamos que si hubo quejas en el rango dado.
    if (filteredDates.length > 0) {
      //Sí hubo quejas. 
      
      //Se crean las fechas que se encuentran dentro del rango a revisar
      var datesToCheck = getDates(dateFrom, dateTo);
      //Ahora se almacena las fechas en las que se dieron quejas y la cantidad de quejas. Se analizan aquellas fechas que cumplieron con las condiciones.
      var datesComplains = datesCount(filteredDates);
      
      //Ahora se agregan las fechas con quejas encontradas encontra las fechas que están dentro del rango.
      var analysisDates = completeDates(datesComplains, datesToCheck);
      
      //Se agregan a la hoja de análisis las fechas de queja y su cantidad.
      analysisSheet.getRange(9, 1, analysisDates.length, 2).setValues(analysisDates);
      
      //La última fila de los datos.
      var dataRange = 8 + analysisDates.length;
      
      //Se copia el dato de LCI en los datos.
      analysisSheet.getRange("C9:C" + dataRange).setValue("=$H$6");
      //Se copia el dato de la media en los datos.
      analysisSheet.getRange("D9:D" + dataRange).setValue("=$H$4");
      //Se copia el dato de LCS en los datos.
      analysisSheet.getRange("E9:E" + dataRange).setValue("=$H$5");
    }
    else {
      //No hubo quejas.
      analysisSheet.getRange(7, 1).setValue("NO SE ENCONTRARON QUEJAS DENTRO EL RANGO ESTABLECIDO.");
    }
    
  }
  else {
    //El cliente no ingresó el rango de fechas.
    analysisSheet.getRange(4, 2).setBackground('red');
    analysisSheet.getRange(5, 2).setBackground('red');
    analysisSheet.getRange(4, 3).setValue("Debe agregar el rango de fechas para el análisis.");
  }
}

/*
Esta función recorre las fechas que se tomaron de la hoja de quejas.
1. Verifica que la fecha se encuentre dentro del rango dado.
2. Agrega aquellas fechas que sí cumplen con la condición a un array.
3. Devuelve el array con las fechas que cumplen la condición.
*/
function checkRange(data, complains, dateFrom, dateTo) {
  //Array donde se van a almacenar aquellas fechas que cumplen con las condiciones
  let filteredDates = [];
  //Se realiza un recorrido entre los datos
  data.forEach(function(datum) {
    //Se verifica que la fecha que se está verificando cumple las condiciones.
    if(Date.parse(datum) >= dateFrom && Date.parse(datum) <= dateTo) {
      //Si cumple las funciones, se almacena en la variable que se creó.
      filteredDates.push(Date.parse(datum));
    }
  });
  
  /*for(let i = 0; i < data.length; i++) {
    if(Date.parse(data[i]) >= dateFrom && Date.parse(data[i]) <= dateTo) {
      
    }
  }*/
  
  //Se devuelven aquellas fechas que cumplen con la condición del rango de fecha.
  return filteredDates;
}

/*
Esta función recorre las fechas que cumplen las condiciones del rango de fecha, para proceder a obtener la cantidad de quejas que se dieron en las fechas filtradas.
*/
function datesCount(filteredDates) {
  //Variable donde se va a almacenar la fecha cuando se dieron quejas y la cantidad de quejas que se dieron.
  let data = [];
  //Bucle que recorre las fechas filtradas
  for(let i = 0; i < filteredDates.length; i++) {
    /*
    Se crea un contador para llevar la cantidad de quejas para la fecha que se está analizando. Se inicializa en 1 porque se debe tomar en cuenta la fecha que se está analizando.
    Si dicha fecha se está analizando es porque ese día hubo al menos una queja; es decir, ella misma.
    */
    let counter = 1;
    //Bucle donde se analizan las fechas posteriores a la fecha que se está analizando.
    for(let j = i+1; j < filteredDates.length; j++) {
      //Se verifica que si la fechas siguientes son iguales a la que se está analizando. Si sí, es porque hay más de una queja.
      if(filteredDates[i] == filteredDates[j]) {
        //Se aumenta el contador de las quejas, porque se encontro una coincidencia.
        ++counter;
        //Se elimina la fecha que se comparó desde la variable en donde están todas las fechas que se están analizando. Se elimina porque ya fue contada como una queja de la fecha
        //que se está analizando.
        filteredDates.splice(j, 1);
        //Como se elimina la fecha que cumple la condición, el recorrido debe empezar desde el mismo índice que se estaba analizando, porque, como se eliminó la fecha que se estaba
        //comparando, una nueva fecha se corrió a ese índice donde estaba la fecha que se acaba de eliminar.
        --j;
      }
    }
    //Se agrega el objeto del día de la fecha en análisis y la cantidad de quejas que se encontraron en la variable de los datos oficiales.
    data.push([new Date(filteredDates[i]), counter]);
  }
  
  //Ordenamos las fechas.
  data.sort(sortFunction);
  
  //Se devuelven las fechas con sus cantidades de quejas.
  return data;
}

/*Taken from: https://stackoverflow.com/questions/16096872/how-to-sort-2-dimensional-array-by-column-value/16097058#16097058*/
function sortFunction(a, b){
  if (a[0] === b[0]) {
    return 0;
  }
  else {
    return (a[0] < b[0]) ? -1 : 1;
  }
}

/*
Función que compara las fechas con quejas encontradas con las fechas generadas en el rango.
Este enfoque funciona solamente porque las fechas con quejas están ordenadas.
*/
function completeDates(foundDates, rangeDates) {
  //Variable donde se va a almacenar la fecha cuando se dieron quejas y la cantidad de quejas que se dieron.
  let data = [];
  
  //Bucle donde se compara las fechas con quejas contra las fechas generadas en el rango.
  for (let  i = 0; i < rangeDates.length; i++) {
    //Verificamos si todavía quedan fechas con quejas para comparar contra las fechas del rango.
    if(i <= foundDates.length - 1) {
      //Verificamos si la fecha generada es igual a la fecha con quejas encontrada
      if(rangeDates[i].getTime() == foundDates[i][0].getTime()) {
        //Se agrega el objeto del día de la fecha en análisis y la cantidad de quejas que se encontraron en la variable de los datos oficiales.
        data.push([new Date(rangeDates[i]), foundDates[i][1]]);
        //Eliminamos la fecha con quejas de la búsqueda.
        foundDates.splice(i, 1);        
      }
      else {
        //Si no hay coincidencias, significa que en ese día no hubo quejas, por lo que se agrega cero (0) quejas.
        //Se agrega el objeto del día de la fecha en análisis y la cantidad de quejas.
        data.push([new Date(rangeDates[i]), 0]);
      }
      //Eliminamos del array la fecha generada, porque ya no se tiene que comparar.
      rangeDates.splice(i, 1);
      //Se tiene que mantener en el mismo lugar.
      --i;
    }
    //Si ya no hay con qué comparar, significa que ya se han agregado todas las fechas con quejas a las fechas generadas.
    else
      break;
  }
  
  //Se verifica si quedaron fechas generadas sin quejas. Puede que suceda si se llega a un punto donde ya se verificaron las fechas con quejas.
  if(rangeDates.length > 0) {
    for (let  i = 0; i < rangeDates.length; i++) {
      let complainDate = [new Date(rangeDates[i]), 0];
      data.push(complainDate);
    }
  }
  return data;
}

/*Taken from: https://stackoverflow.com/questions/4413590/javascript-get-array-of-dates-between-2-dates*/
function getDates(startDate, stopDate) {
  var dateArray = new Array();
  var currentDate = startDate;
  while (currentDate <= stopDate) {
    dateArray.push(new Date(currentDate));
    currentDate = addDays(currentDate, 1);
  }
  return dateArray;
}

/*Taken and modified from https://stackoverflow.com/questions/4413590/javascript-get-array-of-dates-between-2-dates*/
function addDays(dat, days) {
  dat.setDate(dat.getDate() + days);
  return dat;
}

/*Clears data from the sheet*/
function clearData(sheet) {
  sheet.getRange("A9:E").clearContent();
  sheet.getRange(4, 2).setBackground('white');
  sheet.getRange(5, 2).setBackground('white');
  sheet.getRange(4, 3).clearContent();
  sheet.getRange(7, 1).clearContent();
}