function facturacion() {
  /**
   *
   *
   *
   * CONSTANTS
   *
   *
   *
   */
  const SHEET = {
    FACTURACION: 'FACTURACION',
  };

  const COL = {
    fechaDeVencimiento: 5,
    fechaDePago: 6,
  };

  const STATE = {
    POR_COBRAR: 'Por cobrar',
    VENCIDO: 'Vencido',
    CANCELADO: 'Cancelado',
  };

  /**
   *
   *
   *
   * INITIALIZATIONS
   *
   *
   *
   */
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var facturacionSheet = sheet.getSheetByName(SHEET.FACTURACION);

  /**
   *
   *
   *
   * MAIN APP FUNCTION
   *
   *
   *
   */
  (function () {
    var data = facturacionSheet.getDataRange().getValues();

    var today = new Date();

    for (var i = 1; i < data.length; i++) {
      var fechaPago = data[i][COL.fechaDePago];
      var fechaVencimiento = new Date(data[i][COL.fechaDeVencimiento]);

      if (fechaPago) {
        // Si hay una fecha de pago, establecer "Cancelado"
        facturacionSheet.getRange(i + 1, 8).setValue(STATE.CANCELADO);
      } else if (fechaVencimiento < today) {
        // Si la fecha de vencimiento es pasada y no hay fecha de pago, establecer "Vencido"
        facturacionSheet.getRange(i + 1, 8).setValue(STATE.VENCIDO);
      } else {
        // Si no es cancelado ni vencido, establecer "Por cobrar"
        facturacionSheet.getRange(i + 1, 8).setValue(STATE.POR_COBRAR);
      }
    }
  })();
}
