function generarFlujoCaja() {
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
    DATOS_FORMULA: 'DATOS_FORMULA',
    RESUMEN_CONCEPTO: 'RESUMEN_CONCEPTO',
    FLUJO_DE_CAJA: 'FLUJO_DE_CAJA',
  };

  const COLUMN = {
    CONCEPTO_EGRESO: 'B2:B',
  };

  const BANKS_KEY_CHAR = '_';

  const START_COL = {
    egreso: 2,
    ingreso: 4,
    flujoDeCaja: 2,
  };

  const START_ROW = {
    saldoInicial: 2,
    ingreso: 3,
    egreso: 4,
    gananciaPerdida: 5,
    acumulado: 6,
    rentabilidad: 7,
  };

  const JUMP = {
    COL: 4,
  };

  const STYLE = {
    green: '#55B03E',
    red: '#F0462E',
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

  var conceptSummarySheet = sheet.getSheetByName(SHEET.RESUMEN_CONCEPTO);

  var flujoDeCajaSheet = sheet.getSheetByName(SHEET.FLUJO_DE_CAJA);

  var allSheets = sheet.getSheets();

  var sheetCount = 0;
  for (var sheetIndex = 0; sheetIndex < allSheets.length; sheetIndex++) {
    var currentSheet = allSheets[sheetIndex];
    if (currentSheet.getName().startsWith(BANKS_KEY_CHAR)) {
      sheetCount++;
    }
  }

  var tableRowsLength = countConceptoEgresoRows() + 4;

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
    // saldo inicial
    const saldoInicial = Array.from({ length: 13 }, (_, index) => 0);
    writeValuesToSheet(saldoInicial, START_ROW.saldoInicial);

    // ingreso
    const ingresoTotals = sumTotals(START_COL.ingreso);
    writeValuesToSheet(ingresoTotals, START_ROW.ingreso);

    // egreso
    const egresoTotals = sumTotals(START_COL.egreso);
    writeValuesToSheet(egresoTotals, START_ROW.egreso);

    // ganancia/perdida
    const gananciaPerdida = calcularGananciaPerdida(
      ingresoTotals,
      egresoTotals
    );
    writeValuesToSheet(gananciaPerdida, START_ROW.gananciaPerdida, true);

    // acumulado
    const acumulado = calcularAcumulado(gananciaPerdida, saldoInicial);
    writeValuesToSheet(acumulado, START_ROW.acumulado, true);

    // rentabilidad
    const rentabilidad = calcularRentabilidad(egresoTotals, ingresoTotals);
    writeValuesToSheet(rentabilidad, START_ROW.rentabilidad, true, '0.00%');
  })();

  /**
   *
   *
   *
   * FEATURE FUNCTIONS
   *
   *
   *
   */
  function sumTotals(startCol) {
    var resultArray = [];
    let rowJump = tableRowsLength + 3;
    var sumAllValues = 0;

    for (
      var startRow = tableRowsLength + 1;
      startRow <= conceptSummarySheet.getLastRow();
      startRow += rowJump
    ) {
      var sumColumnValues = 0;
      for (var col = startCol; col <= sheetCount * JUMP.COL; col += JUMP.COL) {
        var cellValue = conceptSummarySheet.getRange(startRow, col).getValue();
        // console.log(startRow, col);
        // console.log({ cellValue });
        if (!isNaN(cellValue) && typeof cellValue === 'number') {
          sumColumnValues += Number(cellValue.toFixed(2));
        } else {
          // Logger.log('Non-numeric value encountered at row %s, col %s', startRow, col);
        }
      }

      sumAllValues += sumColumnValues;

      resultArray.push(sumColumnValues);
    }

    resultArray.push(sumAllValues);

    return resultArray;
  }

  function writeValuesToSheet(
    values,
    row,
    paintCells = false,
    numberFormat = ''
  ) {
    for (var i = 0; i < values.length; i++) {
      var cell = flujoDeCajaSheet.getRange(row, START_COL.flujoDeCaja + i);
      cell.setValue(values[i]);

      if (paintCells) {
        cell.setBackground(values[i] >= 0 ? STYLE.green : STYLE.red);
      }

      if (numberFormat) {
        cell.setNumberFormat(numberFormat);
      }
    }
  }

  function calcularGananciaPerdida(ingresos, gastos) {
    var gananciaPerdidaArray = [];
    for (var i = 0; i < ingresos.length; i++) {
      var resultado = ingresos[i] - gastos[i];
      gananciaPerdidaArray.push(resultado);
    }
    return gananciaPerdidaArray;
  }

  function calcularAcumulado(gananciaPerdida, saldoInicial) {
    var acumuladoArray = [];
    for (var i = 0; i < gananciaPerdida.length; i++) {
      var acumulado = gananciaPerdida[i] + saldoInicial[i];
      acumuladoArray.push(acumulado);
    }
    return acumuladoArray;
  }

  function calcularRentabilidad(egreso, ingresos) {
    var rentabilidadArray = [];
    for (var i = 0; i < egreso.length; i++) {
      var resultado =
        egreso[i] !== 0 ? (ingresos[i] - egreso[i]) / egreso[i] : 0;
      rentabilidadArray.push(resultado);
    }
    return rentabilidadArray;
  }

  /**
   *
   *
   *
   * UTIITIES
   *
   *
   *
   */
  function countConceptoEgresoRows() {
    var datosFormSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET.DATOS_FORMULA
    );
    var columnaConceptoEgreso = datosFormSheet.getRange(
      COLUMN.CONCEPTO_EGRESO + datosFormSheet.getLastRow()
    );
    var totalRows = columnaConceptoEgreso.getValues().filter(String).length;
    return totalRows;
  }
}
