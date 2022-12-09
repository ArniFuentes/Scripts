function enviarCorreo(proveedor, destinatarios, asunto) {

    try {

        // hoja con la tabla dinámica y la nomenclatura
        var sheet = SpreadsheetApp.getActive().getSheetByName(proveedor);

        // rango de la tabla dinámica (arreglo de arreglos)
        var rango = sheet.getRange(11, 1, sheet.getLastRow(), 10).getValues();

        // declaración de la variable contenedora del body 
        var body = '';

        body += '<p>' + sheet.getRange('A1').getValue() + '</p>';
        body += '<table>';

        body += "<tr>";
        body += "<td><b>BL</b></td>";
        body += "<td>Backlog o pedidos pendientes.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>Q</b></td>";
        body += "<td>Q de pedidos que debían salir el día.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>D</b></td>";
        body += "<td>Despachos hechos hasta hoy.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>E</b></td>";
        body += "<td>Visitas exitosas.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>F</b></td>";
        body += "<td>Visitas Fallidas.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>SOT</b></td>";
        body += "<td>(Ship on Time) Asignado el día de entrega teorica, en el caso de los SLA ";
        body += "en horas, 1 hora antes del cumplimiento del SLA.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>OTD</b></td>";
        body += "<td>(On Time Delivery) Visitados en o antes del la fecha pactada con el cliente.</td>";
        body += "</tr>";

        body += "<tr>";
        body += "<td><b>OTIF</b></td>";
        body += "<td>(On Time In Full Delivery) Todas las entregas entregadas en o antes";
        body += "de la fecha pactada con el cliente.</td>";
        body += "</tr>";

        body += "</table>";

        // salto de linea
        body += "<br>";

        // Encabezados de la tabla
        body += "<table style='border:1px solid #dddddd;border-collapse:collapse;text-align:center'";
        body += "border = 1 cellpadding = 5>";
        body += "<tr style='background-color: #4b5796; color: white'>";
        body += "<th>Dia entrega Teo</th>";
        body += "<th>Servicio</th>";
        body += "<th>BL Acum</th>";
        body += "<th>Q</th>";
        body += "<th>D</th>";
        body += "<th>E</th>";
        body += "<th>F</th>";
        body += "<th>SOT</th>";
        body += "<th>OTD</th>";
        body += "<th>OTIF</th>";
        body += "</tr>";

        // Insertando los registros de la tabla (rango.length es 140)
        for (var i = 1; i < rango.length - 10; i++) {

            // si es el último registro, que le ponga letra negra remarcada
            if (i == rango.length - 11) {

                body += "<tr style='background-color: #f2f2f2'>";

                if (String(rango[i][0]).match('Total') != null) {
                    body += "<td><b>" + rango[i][0] + "</b></td>";
                } else if (String(rango[i][0]).match('Suma') != null) {
                    body += "<td><b>" + rango[i][0] + "</b></td>";

                } else {
                    body += "<td><b>" + Utilities.formatDate(
                        new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy"
                    ) + "</b></td>";
                }

                body += "<td><b>" + rango[i][1] + "</b></td>";
                body += "<td><b>" + (rango[i][2]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][3]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][4]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][5]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][6]).toFixed(0) + "</b></td>";
                body += "<td><b>" + (rango[i][7] * 100).toFixed(2) + '%' + "</b></td>";
                body += "<td><b>" + (rango[i][8] * 100).toFixed(2) + '%' + "</b></td>";
                body += "<td><b>" + (rango[i][9] * 100).toFixed(2) + '%' + "</b></td>";

                body += "</tr>";

            } else {

                // Si match devuelve el string 'Total' entonces que la fila tenga un color mas oscuro
                if (String(rango[i][0]).match('Total') != null) {

                    body += "<tr style='background-color: #f2f2f2'>";

                    if (String(rango[i][0]).match('Total') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else if (String(rango[i][0]).match('Suma') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else {
                        body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    }

                    // body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    body += "<td>" + rango[i][1] + "</td>";
                    body += "<td>" + (rango[i][2]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][3]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][4]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][5]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][6]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][7] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][8] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][9] * 100).toFixed(2) + '%' + "</td>";

                    body += "</tr>";

                } else {

                    body += "<tr>";

                    if (String(rango[i][0]).match('Total') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else if (String(rango[i][0]).match('Suma') != null) {
                        body += "<td>" + rango[i][0] + "</td>";
                    } else {
                        body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    }

                    // body += "<td>" + Utilities.formatDate(new Date(rango[i][0]), "GMT+1", "dd/MM/yyyy") + "</td>";
                    body += "<td>" + rango[i][1] + "</td>";
                    body += "<td>" + (rango[i][2]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][3]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][4]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][5]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][6]).toFixed(0) + "</td>";
                    body += "<td>" + (rango[i][7] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][8] * 100).toFixed(2) + '%' + "</td>";
                    body += "<td>" + (rango[i][9] * 100).toFixed(2) + '%' + "</td>";

                    body += "</tr>";

                }

            }

        }

        // final del tag table 
        body += "</table>";

        // salto de linea
        body += "<br>";

        // // Link a la data
        // body += 'Link: ' + linkData;

        MailApp.sendEmail(
            {
                to: destinatarios,
                subject: asunto,
                htmlBody: body,
                replyTo: 'destinatario@rayoapp.com',
                cc: 'destinatario@rayoapp.com'
            }
        )
    } catch (error) {
        console.log(error);
    }

}