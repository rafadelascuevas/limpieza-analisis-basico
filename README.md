# Taller de limpieza y análisis de datos

Vamos a trabajar con un caso real –Las tarjetas *black* de Cajamadrid– modificado para poder aplicar varias técnicas de limpieza.

Es bastante común encontrarse con *datasets* fragmentados, mal formateados, con errores... Para poder realizar un buen analisis, antes tenemos que **unificar**, **limpiar** y **estructurar**.

En este caso tenemos tres archivos Excel: `tarjetas_01.xlsx`,`tarjetas_02.xlsx` y `tarjetas_03.xlsx` que contienen los registros de las tarjetas *black* de los ejecutivos de Cajamadrid. Tenemos que unir los tres. Cada archivo tiene varias sub-hojas, lo que dificulta un poco la tarea.

Los archivos están en `/datasets/hoja_calculo_tarjetas_black/`. 

En este directorio hay varias carpetas numeradas. Si te pierdes en alguno de los pasos, puedes ir a la carpeta siguiente y coger el dataset ya tratado.

Necesitaremos:

- [Cuenta de Google](https://accounts.google.com)
- [Open Refine](http://openrefine.org/)

Descargas opcionales:

- [Talend Open Studio for Big Data](https://www.talend.com/download/talend-open-studio)
- [Microsoft Excel](https://products.office.com/es-es/excel)
- [PostgreSQL](https://www.postgresql.org/download/)

¡Empecemos!

## 1. Unificar

### 1.1. Hoja de cálculo

- Cogemos el archivo `/datasets/hoja_calculo_tarjetas_black/01_originales/tarjetas_01.xlsx` y lo subimos a [Google Drive](https://www.google.com/intl/es_es/drive/). 

- Una vez en Drive, lo abrimos con **Google Spreadsheets** ![Hojas de cálculo de Google](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/spreadsheet-icon.png "Hojas de cálculo de Google")

- ¡Problema! Hay varias sub-hojas. Para analizar los datos cómodamente necesitamos una sola tabla que contenga todos los datos.
Google Spreadsheets no tiene opción de juntar todas las sub-hojas. Pero con un poco de **Javascript** podemos sacar el contenido en varios archivos csv y luego juntarlos.

- Vamos a *Herramientas --> Editor de Secuencias de comandos...*

- Le damos nombre al script, por ejemplo "hojas_a_csv". Borramos todo e insertamos el siguiente código, que nos servirá para guardar todas las hojas en CSV.

```javascript
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

- Le damos a **Guardar** y ejecutamos la función **onOpen**

- Una vez ejecutado, volvemos a la pestaña de la hoja de cálculo. Debería aparecer un nuevo menú llamado **csv**.

- Seleccionamos **csv-->Export as csv files**. Puede tardar un poco. Un pop up nos avisa del destino de los archivos: una carpeta en el directorio raíz de Google Drive.

¡Bien! Ya tenemos nuestros archivos csv. 

- Vamos a la carpeta de salida en Drive, seleccionamos y botón derecho **--> Descargar**. Nos sandrá un zip con todos los archivos.

- Descompriminos en nuestro escritorio y le damos un nombre más reconocible a la carpeta.

Repetir el proceso con los archivos `tarjetas_02.xlsx` y `tarjetas_03.xlsx`.

Tenemos 64 archivos `.csv` Vamos a unirlos con **Talend Open Studio for Big Data**.

### 1.2. Talend Open Studio for Big Data

Talend es una herramienta muy potente que permite trabajar con datos a través de una interfaz gráfica, sin tener que escribir código. Pero tiene una curva de aprendizaje bastante pronunciada y la cantidad de opciones puede ser abrumadora.

Utiliza **componentes** que se van añadiendo a una mesa de trabajo y se conectan entre sí a modo de diagrama. Esto hace que sea muy flexible a la hora de afrontar un problema. Podemos crear nuestros propios **trabajos** personalizados.

En este caso vamos a crear un **trabajo** que nos permita extraer todos los archivos csv de un directorio y fundirlos en uno.

- Abrir Talend **--> Create a new project**

- **Create a new... --> Job** (o "Trabajo")

- En la pestaña derecha, ***Palette** buscar el componente **tFileList_1**. Arrastrarlo a la mesa de trabajo.

- Hacer lo mismo con los componentes **tFileInputDelimited_1**, **tUnite_1**, **tLogRow_1** y **tFileOutputDelimited_1**

- Conectar de izquierda a derecha todos los componentes. Botón derecho en el primer componente **-->Fila-->Iterate** y pinchamos con el botón izquierdo en el siguiente componente.

Vamos a definir las opciones de entrada y salida de nuestros csv.

- Pinchar en **tFileList_1** y a continuación, abajo, en la pestaña **Component**. Seleccionar el directorio donde tenemos todos los archivos csv.

- Pinchar en **tFileInputDelimited_1** y a continuación, abajo, en la pestaña **Component**. En **Nombre de Archivo/flujo** pulsamos **Ctrl-Espacio** y en el desplegable seleccionamos **tFileList.CURRENT_FILEPATH**

- Como **Separador de Campo** ponemos **","**

- Vamos a definir las columnas que tienen nuestros archivos csv. Pinchamos en **Edit Shcema** y en el signo **+** 9 veces para agregar 9 columnas. Nombramos cada columna con el nombre que tienen en los csv. en **Tipo** (de dato) dejamos todos como **string**. Al aceptar nos preguntará si propaga los cambios al resto de componentes. Le decimos que sí.

- Pinchar en **tLogRow_1** y a continuación, abajo, en la pestaña **Component**. En **Mode** Seleccionar **Table**.

- Pinchar en **tFileOutputDelimited_1** y a continuación, abajo, en la pestaña **Component**. Como **Separador de Campo** ponemos **";"**. Hay que fijarse en **Nombre de Archivo**. En esa ruta aparecerá nuestro archivo csv unificado. Por defecto se llama `out.csv`.

- Finalmente vamos a la pestaña **Run** y ejecutamos el trabajo con **Run**

Ya tenemos nuestro archivo único, `out.csv`. Vamos a limpiarlo con **Open Refine**.

## 2. Limpiar

### 2.1. Open Refine

## 3. Analizar

### 3.1. Estructurar

#### Hoja de cálculo

### 3.2. Preguntar a los datos

#### Hoja de cálculo

##### Tablas dinámicas

#### Base de datos

Para pasar de formato access de Microsoft a mac: primero, descargar [ActualOCB](https://www.macupdate.com/app/mac/20360/actual-odbc-driver-for-access/download)

