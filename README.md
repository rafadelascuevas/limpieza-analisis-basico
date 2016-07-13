# Limpieza y análisis de datos

## Limpieza
===========

### Hoja de cálculo
-------------------

Vamos a trabajar con un caso real, modificado para poder aplicar varias técnicas de limpieza. 
Es rara la vez que nos darán una base de datos limpia y estructurada. La mayoría de veces estará fragmentada, mal formateada, con errores... Para poder realizar un buen analisis, tenemos que unificarla, estructurarla y limpiarla.

En este caso tenemos tres archivos Excel: `tarjetas_01.xlsx`,`tarjetas_02.xlsx` y `tarjetas_03.xlsx` con los registros de las tarjetas *black* de ejecutivos de Cajamadrid. Tenemos que unir los tres. Cada archivo tiene varias sub-hojas, lo que dificulta un poco la tarea.

Los materiales están en `/datasets/hoja_calculo_tarjetas_black/`. 

En este directorio hay varias carpetas numeradas. Si te pierdes en alguno de los pasos, puedes ir a la carpeta siguiente y coger el dataset ya tratado.

¡Empecemos!

- Cogemos el archivo `/datasets/hoja_calculo_tarjetas_black/01_originales/tarjetas_01.xlsx` y lo subimos a **Google Drive**. 

- Una vez en Drive, lo abrimos con **Google Spreadsheets**.

- ¡Problema! Hay varias sub-hojas. Para analizar los datos cómodamente necesitamos una sola tabla que contenga todos los datos.

Google Spreadsheets no tiene opción de juntar todas las sub-hojas. Pero con un poco de **Javascript** podemos sacar el contenido en varios archivos csv y luego juntarlos.

Vamos a *Herramientas --> Editor de Secuencias de comandos...*

Insertamos el siguiente script, que nos servirá para guardar todas las hojas en CSVs. Hay que pegarlo justo a continuación de `function myFunction() {

```javascript
/*
 * script to export data in all sheets in the current spreadsheet as individual csv files
 * files will be named according to the name of the sheet
 * author: Michael Derazon
*/

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "export as csv files", functionName: "saveAsCSV"}];
  ss.addMenu("csv", csvMenuEntries);
};

function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append ".csv" extension to the sheet name
    fileName = sheet.getName() + ".csv";
    // convert all available sheet data to csv format
    var csvFile = convertRangeToCsvFile_(fileName, sheet);
    // create a file in the Docs List with the given name and the csv data
    folder.createFile(fileName, csvFile);
  }
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
```

Le damos a **Ejecutar**

Nos pide que guardemos y nombremos el proyecto. Lo llamaremos, por ejemplo, "hojas_a_csv"

Es posible que haya que ejecutar de nuevo. Nos pedirá que demos permisos al script.

Una vez ejecutado, volvemos a la pestaña de la hoja de cálculo. Debería aparecer un nuevo menú llamado **csv**. Si el menú no aparece, actualizamos la página. Si el problema persiste probamos a cerrar y abrir de nuevo el navegador.

Seleccionamos **csv-->Export as csv files**. Puede tardar un poco.

Un pop up nos avisa del destino de los archivos: una carpeta en el directorio raíz de Google Drive.

¡Bien! Ya tenemos nuestros archivos csv. Vamos a la carpeta de salida en Drive, seleccionamos y botón derecho *--> Descargar*.

Nos sandrá un zip con todos los archivos. Descompriminos en nuestro escritorio y le damos un nombre más reconocible a la carpeta.

Vamos a unirlos con **Talend Open Studio for Big Data**.

### Talend Open Studio for Big Data
-----------------------------------

### Open Refine
---------------

## Análisis
===========

### Estructurar
---------------

#### Hoja de cálculo

### Preguntar a los datos
-------------------------
#### Hoja de cálculo

##### Tablas dinámicas

#### Base de datos

Para pasar de formato access de Microsoft a mac: primero, descargar [ActualOCB](https://www.macupdate.com/app/mac/20360/actual-odbc-driver-for-access/download)

