function cuentaClientes() {
  /**
   * CONSTANTS
   */
  const SHEET = {
    CUENTA_CLIENTES: 'CUENTA_CLIENTES',
    FACTURACION: 'FACTURACION',
  };

  const MONTHS = [
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

  const STATE = ['Por cobrar', 'Vencido', 'Cancelado'];

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

    for (var i = 0; i < STATE.length; i++) {
      var estado = STATE[i];
      var totales = obtenerTotalesPorEstado(estado);
      escribirEstadoCuentaClientes(i, estado, totales);
      agregarTotalesPorMes(totales);
    }
  })();

  /**
   * FEATURES
   */
  function escribirEstadoCuentaClientes(i, estado, totales) {
    var filaActual = i === 0 ? 1 : cuentaClientesSheet.getLastRow() + 3;
    // Escribir el estado en la primera fila
    cuentaClientesSheet
      .getRange(filaActual, 1)
      .setValue(estado)
      .setBackground('#FFF2CC');

    // Escribir el encabezado de meses y establecer colores
    escribirEncabezadoCuentaClientes(filaActual + 1);

    // Escribir los datos de clientes y totales
    var clientes = Object.keys(totales);
    for (var j = 0; j < clientes.length; j++) {
      var cliente = clientes[j];
      var datosCliente = [cliente];
      for (var k = 0; k < MONTHS.length; k++) {
        var mes = MONTHS[k].toLowerCase();
        datosCliente.push(
          totales[cliente][mes] ? totales[cliente][mes].total : 0
        );
      }
      cuentaClientesSheet
        .getRange(filaActual + j + 2, 1, 1, datosCliente.length)
        .setValues([datosCliente])
        .setNumberFormat('$#,##0.00')
        .setFontFamily('Inconsolata');
      cuentaClientesSheet.getRange('A:A').setFontFamily('Arial');
    }
  }

  function escribirEncabezadoCuentaClientes(filaInicial) {
    // Escribir encabezado de meses
    var rangeHeader = cuentaClientesSheet.getRange(
      filaInicial,
      1,
      1,
      MONTHS.length + 1
    );
    rangeHeader.setValues([['CLIENTES'].concat(MONTHS)]);

    // Establecer color de fondo para el encabezado de cliente (azul claro)
    rangeHeader.getCell(1, 1).setBackground('#c9daf8');

    // Establecer color de fondo para los meses (verde claro)
    rangeHeader.offset(0, 1, 1, MONTHS.length).setBackground('#d9ead3');
  }

  function obtenerTotalesPorEstado(estado) {
    // Objeto para almacenar los totales por cliente y mes
    var totales = {};

    // Iterar sobre las filas de la hoja de facturación
    var data = facturacionSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      // Empezamos desde 1 para saltar el encabezado
      var cliente = data[i][2];
      var importe = parseFloat(
        String(data[i][3]).replace('$', '').replace(',', '')
      );
      var fechaVencimiento = new Date(data[i][5]);
      var fechaPago = new Date(data[i][6]);
      var mes;
      if (!isNaN(fechaPago.getTime())) {
        mes = fechaPago;
      } else {
        mes = fechaVencimiento;
      }
      var mes = MONTHS[mes.getMonth()].toLowerCase();
      var estadoFacturacion = data[i][7];

      // Verificar si el estado coincide
      if (estadoFacturacion === estado) {
        // Inicializar el objeto de totales si es la primera vez que se encuentra el cliente
        if (!totales[cliente]) {
          totales[cliente] = {};
          MONTHS.forEach(function (month) {
            totales[cliente][month.toLowerCase()] = { total: 0 };
          });
        }
        // Sumar el importe al total correspondiente al cliente y al mes
        totales[cliente][mes].total += importe;
      }
    }
    return totales;
  }

  function agregarTotalesPorMes(totales) {
    var filaActual = cuentaClientesSheet.getLastRow();
    var clientes = Object.keys(totales);
    var totalPorMes = {};

    // Calcular el total por mes
    for (var j = 0; j < clientes.length; j++) {
      var cliente = clientes[j];
      for (var k = 0; k < MONTHS.length; k++) {
        var mes = MONTHS[k].toLowerCase();
        totalPorMes[mes] =
          (totalPorMes[mes] || 0) +
          (totales[cliente][mes] ? totales[cliente][mes].total : 0);
      }
    }

    // Escribir fila 'TOTAL'
    var filaTotal = filaActual + 1;
    cuentaClientesSheet
      .getRange(filaTotal, 1)
      .setValue('TOTAL')
      .setFontWeight('bold');

    // Escribir totales por mes
    var datosTotal = [''];
    for (var l = 0; l < MONTHS.length; l++) {
      var mes = MONTHS[l].toLowerCase();
      datosTotal.push(totalPorMes[mes] || 0);
    }
    // Eliminar el primer elemento vacío del array
    datosTotal.shift();

    cuentaClientesSheet
      .getRange(filaTotal, 2, 1, datosTotal.length)
      .setValues([datosTotal])
      .setFontWeight('bold')
      .setFontFamily('Inconsolata');
  }

  /**
   * UTILS
   */
  // Función para limpiar la tabla, manteniendo la fila de títulos
  function limpiarTabla(sheet) {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow > 0 && lastColumn > 0) {
      // Borrar todas las filas y columnas
      sheet.clear();
    }
  }
}
