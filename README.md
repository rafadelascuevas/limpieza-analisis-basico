# Taller de limpieza y análisis de datos

Un conjunto de técnicas y consejos básicos para periodistas de datos que se enfrentan por primera vez a *datasets* sin estructurar.

Vamos a trabajar con un caso real –Las tarjetas *black* de Cajamadrid– modificado para poder aplicar varias técnicas de limpieza.

Es bastante común encontrarse con *datasets* fragmentados, mal formateados, con errores... Para poder realizar un buen analisis, antes tenemos que **unificar**, **limpiar** y **estructurar**.

En este caso tenemos tres archivos Excel: `tarjetas_01.xlsx`,`tarjetas_02.xlsx` y `tarjetas_03.xlsx` que contienen los registros de las tarjetas *black* de los ejecutivos de Cajamadrid. Tenemos que unir los tres. Cada archivo tiene varias sub-hojas, lo que dificulta un poco la tarea.

Los archivos están en [`/datasets/hoja_calculo_tarjetas_black/`](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/datasets/hoja_calculo_tarjetas_black/). 

En este directorio hay varias carpetas numeradas. Si te pierdes en alguno de los pasos, puedes ir a la carpeta siguiente y coger el dataset ya tratado.

Necesitaremos:

- [Cuenta de Google](https://accounts.google.com)
- [Open Refine](http://openrefine.org/)

Descargas opcionales:

- [Talend Open Studio for Big Data](https://www.talend.com/download/talend-open-studio) (gratuito)
- [Microsoft Excel](https://products.office.com/es-es/excel) (de pago)
- [PostgreSQL](https://www.postgresql.org/download/) (gratuito)

¡Empecemos!

## 1. Unificar

### 1.1. Hoja de cálculo

- Bajamos el archivo `/datasets/hoja_calculo_tarjetas_black/01_originales/tarjetas_01.xlsx` y lo subimos a [Google Drive](https://www.google.com/intl/es_es/drive/). 

- Una vez en Drive, lo abrimos con **Google Spreadsheets** ![Hojas de cálculo de Google](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/spreadsheet-icon.png "Hojas de cálculo de Google")

- ¡Problema! Hay varias sub-hojas. Para analizar los datos cómodamente necesitamos una sola tabla que contenga todos los datos.
Google Spreadsheets no tiene opción de juntar todas las sub-hojas. Pero con un poco de **Javascript** podemos sacar el contenido en varios archivos csv y luego juntarlos.

- Vamos a **Herramientas --> Editor de Secuencias de comandos...**

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
  // crea un directorio con el nombre de la hoja
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // añade la extension ".csv" al nombre de la hoja
    fileName = sheet.getName() + ".csv";
    // convierte todos los datos disponibles de la hoja a formato csv
    var csvFile = convertRangeToCsvFile_(fileName, sheet);
    // crea un archivo con el nombre y los datos csv
    folder.createFile(fileName, csvFile);
  }
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

function convertRangeToCsvFile_(csvFileName, sheet) {
  // toma el rango disponible de datos en la hoha de calculo
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // hace un loop en el rango de los datos y crea un string con los datos csv
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // une las columnas de cada fila
        // añade retorno de carro al final de cada fila, excepto la ultima
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

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_01.png "Inicio de Talend")

- **Create a new... --> Job** (o "Trabajo")

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_02.png "Talend nuevo trabajo")

- En la pestaña derecha, **Palette** buscar el componente **tFileList**. Arrastrarlo a la mesa de trabajo.

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_03.png "Talend buscar componentes")

- Hacer lo mismo con los componentes **tFileInputDelimited**, **tUnite**, **tLogRow** y **tFileOutputDelimited**

- Conectar de izquierda a derecha todos los componentes. Botón derecho en el primer componente **-->Fila-->Iterate** y pinchamos con el botón izquierdo en el siguiente componente.

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_05.png "Talend conectar componentes1")
![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_06.png "Talend conectar componentes2")

Vamos a definir las opciones de entrada y salida de nuestros csv.

- Pinchar en **tFileList_1** y a continuación, abajo, en la pestaña **Component**. Seleccionar el directorio donde tenemos todos los archivos csv.

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_07.png "Talend opciones tFileList")

- Pinchar en **tFileInputDelimited_1** y a continuación, abajo, en la pestaña **Component**. En **Nombre de Archivo/flujo** pulsamos **Ctrl-Espacio** y en el desplegable seleccionamos **tFileList.CURRENT_FILEPATH**

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_08.png "Talend opciones tFileInputDelimited")

- Como **Separador de Campo** ponemos **","**

- Marcamos la casilla **opciones csv**. Es importante, porque si no definimos bien los separadores de fila y columna de nuestros csv, se pueden perder datos.

- Vamos a definir las columnas que tienen nuestros archivos csv. Pinchamos en **Edit Schema** y en el signo **+** 9 veces para agregar 9 columnas. Nombramos cada columna con el nombre que tienen en los csv. en **Tipo** (de dato) dejamos todos como **string**. Al aceptar nos preguntará si propagamos los cambios al resto de componentes. Le decimos que sí.

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_09.png "Talend agregar columnas")

- Pinchar en **tLogRow_1** y a continuación, abajo, en la pestaña **Component**. En **Mode** Seleccionar **Table**.

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_10.png "Talend opciones tLogRow")

- Pinchar en **tFileOutputDelimited_1** y a continuación, abajo, en la pestaña **Component**. Como **Separador de Campo** ponemos **";"**. Hay que fijarse en **Nombre de Archivo**. En esa ruta se guardará nuestro archivo csv resultante. Por defecto se llamará `out.csv`.

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_11.png "Talend opciones tFileOutputDelimited")

- Finalmente vamos a la pestaña **Run** y ejecutamos el trabajo con el botón **Run**

![Talend](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/talend_12.png "Talend ejecutar trabajo")

Ya tenemos nuestro archivo único, `out.csv`. Vamos a limpiarlo con **Open Refine**.

## 2. Limpiar

### 2.1. Open Refine

Al ejecutar Refine no se abre ninguna ventana. Lo que hace Refine es montar un servidor local en el puerto 3333. Así que vamos a abrir una pestaña del navegador con la dirección `http://127.0.0.1:3333/`.

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_01.png "Refine direccion")

- Vamos a la pestaña **Create Project** y seleccionamos nuestro archivo, `out.csv`.

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_02.png "Refine nuevo proyecto")

- Refine previsualiza la tabla. Tenemos que seleccionar la codificación de caracteres: **UTF-8**. Si tenemos un falso encabezamiento de columnas, hay que seleccionar **Ignore first 1 line(s) at beginning of file**.

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_022.png "Refine previsualizacion tabla")

Para poder analizar nuestra tabla necesitamos que los **elementos que se repiten** (por ejemplo, todas las entradas con un nombre y apellidos) sean **exactamente iguales**. En nuestra tabla hay espacios de más que no vemos, o puntos al final de un nombre. Estos caracteres extra harán que más tarde, al agrupar los elementos para su análisis, aparezcan grupos distintos con entradas que deberían ser iguales.

- Vamos a comprobarlo. Vamos a la columna **NOMBRE**. Pinchamos en el encabezamiento y seleccionamos **Facet --> Text facet**.

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_023.png "Refine text facet")

A la izquierda aparecen todas las entradas **distintas** de la columna **NOMBRE**. Al lado de cada una aparece el número de filas que tiene cada una. Por ejemplo: Hay 448 filas con el **NOMBRE** **"ACOSTA CUBERO, JOSE"**.

Podemos ver que, por ejemplo **"CAFRANGA CAVESTANY, MARIA CARMEN"** aparece dos veces. No debería ser así, vamos a arreglarlo.

Además, tenemos una columna con dos tipos de entradas juntas: **Nombre de comercio** y **Actividad**. Para saber, por ejemplo, cuanto dinero gastaba un asesor en un comercio determinado, tenemos que separar esos conceptos.

¡Empecemos!

- Lo primero es limpiar las celdas de espacios no deseados. Vamos a la columna **NOMBRE**. Pinchamos en el encabezamiento y seleccionamos **Edit cells --> Common transforms --> Trim leading and trailing whitespace**. Repetimos el proceso en todas las columnas.

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_03.png "Refine limpiar espacios")

- Ahora vamos a editar masivamente los nombres en la columna **NOMBRE** para unificar los que no son iguales. Para ello usaremos **Edit cells --> Cluster and edit**

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_04.png "Refine Cluster and edit")

- Cluster & edit reconoce los patrones de caracteres parecidos y sugiere fusiones. Comprobamos que las sugerencias son correctas. Pinchamos en **Select all** y a continuación en **Merge selected & re-cluster**. Si todo va bien, en un segundo reconocimiento todo está limpio. Refine utiliza distintos métodos de reconocimiento de patrones. El que viene por defecto nos sirve en este caso, pero a veces el patrón es más complejo.

- Ahora vamos a separar esa columna con dos tipos de entradas juntas. Pinchamos en el encabezado **Nombre_Comercio_y_Actividad** y seleccionamos **Edit column --> Split into several columns**

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_05.png "Refine separar columnas1")

- Tenemos varias opciones para separar la columna. En nuestro caso, hay un patrón de caracteres que separa un concepto de otro: tres guiones bajos. Así que vamos a poner tres guiones bajos en el campo **Separator**. 

![Refine](https://github.com/rafadelascuevas/limpieza-analisis-basico/blob/master/img/refine_06.png "Refine separar columnas2")

- Podemos aplicar más transformaciones. Por ejemplo, separar las columnas de los apellidos y el nombre por la coma. Alternar el nombre de las columnas, sustituir caracteres de forma masiva...

- Cuando nuestra tabla esté limpia, exportamos. Vamos a guardar en formato Excel para poder trabajar indistintamente en Excel o Google Spreadsheets. **Export --> Excel**


## 3. Analizar

### 3.1. Estructurar

Para realizar un buen análisis es muy importante que los datos estén bien estratificados. 

Lo principal: "Pensar en vertical". Cada columna aporta una capa de conocimiento. Cuanto más concreto, mejor.

Por ejemplo: Tenemos una columna con la fecha completa de cada transacción. Pero si queremos un análisis más profundo, podemos separar esa columna en días, meses y años. 

#### Hoja de cálculo

Volvemos a Google Drive. Subimos nuestro archivo limpio y lo abrimos con Spreadsheets.

- Buena práctica 1: Añadimos al nombre de nuestro archivo **_ORIGINAL** y creamos una copia con el añadido **_TRABAJO**. Así, si comentemos errores podemos volver a la hoja original. 

- Buena práctica 2: Añadimos una columna con el encabezado **ID**. Ponemos un 1 en la primera fila, y rellenamos la serie automáticamente.

- Vamos a crear columnas con rangos de fecha que nos será útiles para analizar los días, meses y horas en los que se hacía mayor uso de las tarjetas. Creamos columnas llamadas **DIA**, **MES**, **AÑO** y **RANGO HORA**

- La función `=TEXT` (o `=TEXTO` si estamos en Excel y tenemos una versión en castellano) convierne un número a texto con un formato predefinido. En la columna **MES** escribimos `=TEXT(B2;"MMMM")`. Estamos diciendo que tome el valor de la celda B2 y de ahí pase el mes a texto, en formato largo: "MMMM".

- Hacemos algo similar en la columna **AÑO**: `=TEXT(B2;"AAAA")`

- Y en la columna **DIA**: `=TEXT(B2;"DD")`

- Además vamos a sacar las horas. Pero los datos no están en formato de hora *hh:mm:ss* por lo que tenemos que sustituir el punto `.` por dos puntos `:`. Seleccionamos la columna entera y damos a **Editar --> Buscar y sustituir**.

- Ahora sí, En la columna **RANGO HORA** escribimos: `=TEXT(F2;"hh")`

### 3.2. Preguntar a los datos

- Para *interrogar* a los datos usamos tablas dinámicas. Hay que empezar por lo general (¿cuanto dinero se han gastado en total?, ¿quién ha gastado más dinero?) a lo concreto (¿Quién hizo el cargo más grande? ¿En qué días se usaron más las tarjetas?)

#### Hoja de cálculo

##### Tablas dinámicas

- Colocamos en cursor en la casilla **A1**. **Vamos al menú Datos --> Tabla dinámica**

- En el apartado **Valores** seleccionamos **IMPORTE**. A continuación, vamos probando conjuntos de datos en los apartados **Filas** y **columnas**

#### Base de datos

Excel tiene un límite de 1,048,576 filas y 16,384 columnas. Para analizar bases de datos grandes tenemos que utilizar herramientas de tratamiento de bases de datos. Existen varios lenguajes de bases de datos relacionales. El más extendido es SQL, así que vamos a utilizar una herramienta open source gratuita: [PostgreSQL](https://www.postgresql.org/download/).

##### PostgreSQL

SQL es un lenguaje declarativo para hacer consultas y modificaciones en bases de datos. Si tenemos tiempo, practicaremos unas *queries* a la base de datos del ministerio de Agricultura que agrupa todas las playas de España. El archivo está en `datasets/sql_playas/playas.sql`