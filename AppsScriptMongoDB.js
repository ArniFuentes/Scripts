function cargarDatos() {

    // ESTABLECIENDO HOJA DEL SHEETS
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

    // IMPORTANDO LOS DATOS
    var getData = UrlFetchApp.fetch(
        'https://direccionEndpointMongoDB'
    ).getContentText();

    // Arreglo con todos los datos
    var arrayData = JSON.parse(getData);

    try {
        // Quitando los filtros existentes en la hoja Data.
        hoja.getFilter().remove();
    } catch (error) {

    }

    // BORRANDO LOS DATOS DESDE LA FILA 3 HACIA ABAJO
    var rangoBorrar = hoja.getRange(3, 1, hoja.getLastRow() - 2, hoja.getLastColumn());
    rangoBorrar.clearContent();

    // CREAR UN ARRAY PARA QUE SE GUARDEN TODOS LOS DATOS DE FORMA INDIVIDUAL
    const miArreglito = [];

    for (var i = 0; i < arrayData.length; i++) {

        miArreglito.push(
            arrayData[i]['clave1'],
            arrayData[i]['clave2'],
            arrayData[i]['clave3'],
            arrayData[i]['clave4'],
            arrayData[i]['clave5'],
            arrayData[i]['clave6'],
            arrayData[i]['clave7'],
            arrayData[i]['clave8'],
            arrayData[i]['clave9'],
            arrayData[i]['clave10'],
            arrayData[i]['clave11'],
            arrayData[i]['clave12'],
            arrayData[i]['clave13'],
            arrayData[i]['clave14'],
            arrayData[i]['clave15'],
            arrayData[i]['clave16'],
            arrayData[i]['clave17'],
            arrayData[i]['clave18']
        );
    };

    // CREANDO UN ARRAY QUE CONTIENE ARRAYS DE 18 ELEMENTOS
    var arregloFinal = [];

    for (var i = 0; i <= miArreglito.length; i += 18) {
        arregloFinal.push(miArreglito.slice(i, i + 18));
    }

    // ELIMINANDO EL ULTIMO ARRAY VACIO
    arregloFinal.pop();

    // // COPIANDO EL ARREGLO EN EL RANGO
    hoja.getRange('A2:R' + '' + (arregloFinal.length + 1)).setValues(arregloFinal);

    // COPIANDO LAS FUNCIONES DE LA SEGUNDA FILA HASTA EL FINAL DE LA TABLA
    hoja.getRange('s2:be2').copyTo(hoja.getRange('s3:be' + hoja.getLastRow()));

}