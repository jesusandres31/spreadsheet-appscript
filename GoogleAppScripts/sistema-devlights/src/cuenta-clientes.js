function cuentaClientes() {
  /**
   * CONSTANTS
   */
  const SHEET = {
    CUENTA_CLIENTES: 'CUENTA_CLIENTES',
    FACTURACION: 'FACTURACION',
  };

  /**
   * INITIALIZATIONS
   */
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var cuentaClientesSheet = sheet.getSheetByName(SHEET.CUENTA_CLIENTES);
  var facturacionSheet = sheet.getSheetByName(SHEET.FACTURACION);

  /**
   * MAIN APP FUNCTION
   */
  (function () {
    // Limpiar la tabla CUENTA_CLIENTES
    limpiarTabla(cuentaClientesSheet);

    // Obtener datos de la hoja FACTURACION
    var dataRange = facturacionSheet.getDataRange();
    var data = dataRange.getValues();

    // Crear objeto para almacenar totales por cliente y mes
    var totalsByClientAndMonth = {};

    // Recorrer filas de la hoja FACTURACION
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Extraer información relevante
      var cliente = row[2];
      var importe = parseFloat(
        String(row[3]).replace('$', '').replace(',', '')
      );
      var fechaPago = new Date(row[6]);

      // Obtener mes y año de la fecha de pago
      var mes = fechaPago.getMonth() + 1; // Meses en JavaScript van de 0 a 11
      var año = fechaPago.getFullYear();

      // Inicializar objeto para el cliente si no existe
      if (!totalsByClientAndMonth[cliente]) {
        totalsByClientAndMonth[cliente] = {};
      }

      // Inicializar todos los meses con 0 si no existe
      for (var j = 1; j <= 12; j++) {
        if (!totalsByClientAndMonth[cliente][j]) {
          totalsByClientAndMonth[cliente][j] = 0;
        }
      }

      // Sumar importe al total correspondiente
      totalsByClientAndMonth[cliente][mes] += importe;
    }

    // Escribir totales en la hoja CUENTA_CLIENTES
    escribirTotalesEnCuentaClientes(totalsByClientAndMonth);
  })();

  /**
   * FEATURES
   */
  // Función para escribir los totales en la hoja CUENTA_CLIENTES
  function escribirTotalesEnCuentaClientes(totalsByClientAndMonth) {
    // Obtener las columnas de la hoja CUENTA_CLIENTES
    var columnas = cuentaClientesSheet.getRange(1, 1, 1, 13).getValues()[0];

    // Recorrer el objeto totalsByClientAndMonth
    for (var cliente in totalsByClientAndMonth) {
      if (totalsByClientAndMonth.hasOwnProperty(cliente)) {
        // Encontrar la fila correspondiente al cliente
        var filaCliente = encontrarFilaCliente(cliente);

        // Recorrer los meses y escribir los totales en las columnas correspondientes
        for (var mes in totalsByClientAndMonth[cliente]) {
          if (totalsByClientAndMonth[cliente].hasOwnProperty(mes)) {
            // Encontrar la columna correspondiente al mes
            var columnaMes =
              columnas.indexOf(obtenerNombreMes(Number(mes))) + 1;

            // Escribir el total en la celda correspondiente
            cuentaClientesSheet
              .getRange(filaCliente, columnaMes)
              .setValue(totalsByClientAndMonth[cliente][mes])
              .setNumberFormat('$#,##0.00');
          }
        }
      }
    }

    // Agregar una fila adicional con los totales
    agregarFilaTotales();
  }

  // Función para encontrar la fila correspondiente a un cliente en la hoja CUENTA_CLIENTES
  function encontrarFilaCliente(cliente) {
    var lastRow = cuentaClientesSheet.getLastRow();
    if (lastRow < 2) {
      // Si la hoja está vacía, agregar una nueva fila con el nombre del cliente
      cuentaClientesSheet.getRange(2, 1).setValue(cliente);
      return 2;
    } else {
      // Si hay filas, buscar el cliente existente
      var clientesColumna = cuentaClientesSheet
        .getRange(2, 1, lastRow - 1, 1)
        .getValues();
      for (var i = 0; i < clientesColumna.length; i++) {
        if (clientesColumna[i][0] === cliente) {
          return i + 2; // Sumar 2 para ajustarse a la indexación de Google Sheets
        }
      }
      // Si no se encuentra el cliente, agregar una nueva fila con el nombre del cliente
      var nuevaFila = cuentaClientesSheet.getLastRow() + 1;
      cuentaClientesSheet.getRange(nuevaFila, 1).setValue(cliente);
      return nuevaFila;
    }
  }

  // Función para agregar una fila de totales al final de la hoja CUENTA_CLIENTES
  function agregarFilaTotales() {
    var lastRow = cuentaClientesSheet.getLastRow();
    var lastColumn = cuentaClientesSheet.getLastColumn();

    // Agregar "TOTAL" en negrita en la columna "CLIENTES"
    cuentaClientesSheet
      .getRange(lastRow + 1, 1)
      .setValue('TOTAL')
      .setFontWeight('bold');

    // Calcular y sumar los totales para cada mes
    for (var col = 2; col <= lastColumn; col++) {
      var sumaTotal = 0;

      // Sumar los valores de cada fila para la columna actual
      for (var row = 2; row <= lastRow; row++) {
        sumaTotal += cuentaClientesSheet.getRange(row, col).getValue() || 0;
      }

      // Escribir la suma total en la fila "TOTAL"
      cuentaClientesSheet
        .getRange(lastRow + 1, col)
        .setValue(sumaTotal)
        .setNumberFormat('$#,##0.00');
    }
  }

  /**
   * UTILS
   */
  // Función para obtener el nombre del mes a partir de su número (1-12)
  function obtenerNombreMes(numeroMes) {
    var meses = [
      'ENERO',
      'FEBRERO',
      'MARZO',
      'ABRIL',
      'MAYO',
      'JUNIO',
      'JULIO',
      'AGOSTO',
      'SEPTIEMBRE',
      'OCTUBRE',
      'NOVIEMBRE',
      'DICIEMBRE',
    ];
    return meses[numeroMes - 1];
  }

  // Función para limpiar la tabla, manteniendo la fila de títulos
  function limpiarTabla(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // Borrar todas las filas excepto la primera
      sheet.deleteRows(2, lastRow - 1);
    }
  }
}
