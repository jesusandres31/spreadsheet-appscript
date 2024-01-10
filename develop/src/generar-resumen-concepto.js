function generarResumenConcepto() {
  /**
   *
   *
   *
   * CONSTANTS
   *
   *
   *
   */
  const JUMP = {
    ROW: 6,
    COL: 4,
  };

  const BANKS_KEY_CHAR = '_';

  const SHEET = {
    DATOS_FORMULA: 'DATOS_FORMULA',
    RESUMEN_CONCEPTO: 'RESUMEN_CONCEPTO',
  };

  const COLUMN = {
    ConceptoEgreso: 'Concepto Egreso',
    Egreso: 'Egreso',
    ConceptoIngreso: 'Concepto Ingreso',
    Ingreso: 'Ingreso',
  };

  const TABLE_TITLES = Object.values(COLUMN);

  const TOTAL_CELL = 'TOTAL USD:';

  const MONTHS = [
    ['Enero'],
    ['Febrero'],
    ['Marzo'],
    ['Abril'],
    ['Mayo'],
    ['Junio'],
    ['Julio'],
    ['Agosto'],
    ['Septiembre'],
    ['Octubre'],
    ['Noviembre'],
    ['Diciembre'],
  ];

  const COLUMN_INDEX = {
    Egreso: 3,
    Ingreso: 4,
  };

  const STYLE = {
    yellow: '#ffff91',
    purple: '#c1b3e6',
    bold: 'bold',
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
  /**
   * Get the active spreadsheet and the source and destination sheets
   */
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.DATOS_FORMULA
  );
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET.RESUMEN_CONCEPTO
  );

  /**
   * Find the header column
   */
  var headers = dataSheet
    .getRange(1, 1, 1, dataSheet.getLastColumn())
    .getValues()[0];

  /**
   * Get the data from the "Concepto Egreso" column
   */
  var conceptoEgresoData = dataSheet
    .getRange(
      2,
      headers.indexOf(COLUMN.ConceptoEgreso) + 1,
      dataSheet.getLastRow() - 1,
      1
    )
    .getValues();
  conceptoEgresoData = conceptoEgresoData.filter(function (row) {
    return row[0] !== '';
  });

  /**
   * Get data from the "Concepto Ingreso" column
   */
  var conceptoIngresoData = dataSheet
    .getRange(
      2,
      headers.indexOf(COLUMN.ConceptoIngreso) + 1,
      dataSheet.getLastRow() - 1,
      1
    )
    .getValues();
  conceptoIngresoData = conceptoIngresoData.filter(function (row) {
    return row[0] !== '';
  });

  /**
   * get banks
   */
  var banks = getBanksNames();

  /**
   * set JUMP values
   */
  JUMP.ROW += conceptoEgresoData.length;

  /**
   * pesos banks
   */
  const pesosBanks = ['Efectivo', 'BcoCtes'];

  /**
   * cotización periodo
   */
  const contizacionPeriodo = {
    1: 850,
    2: 850,
    3: 850,
    4: 850,
    5: 850,
    6: 850,
    7: 850,
    8: 850,
    9: 850,
    10: 850,
    11: 850,
    12: 850,
  };

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
    // Clear the summary sheet before writing new data
    summarySheet.clear();

    // Write Months row
    for (var i = 0; i < MONTHS.length; i++) {
      var cell = summarySheet.getRange(i * JUMP.ROW + 1, 1);
      cell.setValue(MONTHS[i][0]);
      cell.setBackground(STYLE.yellow);
      cell.setFontWeight(STYLE.bold);

      var monthNumber = getMonthNumber(MONTHS[i][0]);

      // Write Banks row
      writeBanksName(banks, i);
      // Write Table title
      createResumeTable(banks, i, monthNumber);
    }
    // Format sheet
    formatSheet();
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
  function getBanksNames() {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var sheetNamesArray = [];

    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName.startsWith(BANKS_KEY_CHAR)) {
        sheetNamesArray.push(sheetName.substring(1));
      }
    }

    return sheetNamesArray;
  }

  function writeBanksName(banks, i) {
    for (var j = 0; j < banks.length; j++) {
      var bankCell = summarySheet.getRange(i * JUMP.ROW + 2, j * JUMP.COL + 1);
      bankCell.setValue(banks[j]);
      bankCell.setFontWeight(STYLE.bold);
    }
  }

  function createResumeTable(banks, i, targetMonth) {
    // iterate through banks
    for (var k = 0; k < banks.length; k++) {
      var startIndex = i * JUMP.ROW + 3;
      var titleCell = summarySheet.getRange(startIndex, k * JUMP.COL + 1);

      // iterate through table titles
      for (var j = 0; j < TABLE_TITLES.length; j++) {
        titleCell.setValue(TABLE_TITLES[j]);
        titleCell.setBackground(STYLE.purple);
        titleCell.setFontWeight(STYLE.bold);

        fillTable(
          i,
          k,
          targetMonth,
          conceptoEgresoData,
          TABLE_TITLES[j],
          COLUMN.ConceptoEgreso,
          COLUMN.Egreso
        );

        fillTable(
          i,
          k,
          targetMonth,
          conceptoIngresoData,
          TABLE_TITLES[j],
          COLUMN.ConceptoIngreso,
          COLUMN.Ingreso
        );

        // Move to the next cell in the same row
        titleCell = titleCell.offset(0, 1);

        // Set column total
        setColTotal(j, titleCell);
      }
    }
  }

  function fillTable(
    i,
    k,
    targetMonth,
    sheetData,
    tableTitle,
    colLabel,
    colData
  ) {
    fillColumn(i, k, tableTitle, colLabel, sheetData);
    const egresoData = sumAmountByCategory(
      banks[k],
      targetMonth,
      tableTitle,
      colData,
      sheetData
    );
    fillTotal(i, k, tableTitle, colData, egresoData.total);
    fillColumn(i, k, tableTitle, colData, egresoData.resultArray);
  }

  function fillColumn(i, k, tableTitle, columnName, data) {
    if (tableTitle === columnName) {
      var cell = summarySheet.getRange(
        i * JUMP.ROW + TABLE_TITLES.length,
        k * JUMP.COL + TABLE_TITLES.indexOf(columnName) + 1
      );
      for (var j = 0; j < data.length; j++) {
        var currentCell = cell.offset(j, 0);
        currentCell.setValue(data[j][0]);
      }
    }
  }

  function fillTotal(i, k, tableTitle, columnName, total) {
    if (tableTitle === columnName) {
      var cell = summarySheet.getRange(
        i * JUMP.ROW + TABLE_TITLES.length,
        k * JUMP.COL + TABLE_TITLES.indexOf(columnName) + 1
      );
      var totalCell = cell.offset(conceptoEgresoData.length, 0);
      totalCell.setValue(total);

      totalCell.setFontWeight(STYLE.bold);
    }
  }

  function setColTotal(j, titleCell) {
    if (TABLE_TITLES[j] === COLUMN.ConceptoEgreso) {
      var totalCell = titleCell.offset(conceptoEgresoData.length + 1, -1);
      totalCell.setValue(TOTAL_CELL);
      totalCell.setFontWeight(STYLE.bold);
    }
  }

  function sumAmountByCategory(
    bankName,
    targetMonth,
    tableTitle,
    columnName,
    labelArray
  ) {
    if (tableTitle === columnName) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        `${BANKS_KEY_CHAR}${bankName}`
      );

      if (!sheet) {
        throw new Error('Sheet not found: ' + bankName);
      }

      var data = sheet.getDataRange().getValues();
      var categoryTotals = {};
      var total = 0;

      for (var j = 1; j < data.length; j++) {
        var date = data[j][0];
        var monthNumber = parseInt(
          Utilities.formatDate(new Date(date), 'GMT-3', 'MM'),
          10
        );
        var targetMonthInt = parseInt(targetMonth, 10);

        if (
          !isNaN(monthNumber) &&
          !isNaN(targetMonthInt) &&
          monthNumber === targetMonthInt
        ) {
          var category = data[j][2];
          var amount = data[j][COLUMN_INDEX[tableTitle]];

          if (category && amount !== null && !isNaN(amount)) {
            if (categoryTotals[category] === undefined) {
              categoryTotals[category] = 0;
            }

            if (!isNaN(parseFloat(amount))) {
              total += parseFloat(amount);
              categoryTotals[category] += parseFloat(amount);
            }
          }
        }
      }

      var resultArray = [];
      for (var i = 0; i < labelArray.length; i++) {
        var category = labelArray[i][0];
        if (categoryTotals.hasOwnProperty(category)) {
          resultArray.push([categoryTotals[category] || 0]);
        } else {
          resultArray.push([0]);
        }
      }

      if (pesosBanks.includes(bankName)) {
        var cotizacionMes = contizacionPeriodo[targetMonth];
        if (!isNaN(cotizacionMes) && cotizacionMes !== 0) {
          total /= cotizacionMes;
        }
      }

      return { resultArray: resultArray, total: total };
    }
    return { resultArray: [], total: 0 };
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
  function getMonthNumber(monthName) {
    var spanishMonths = {
      Enero: 1,
      Febrero: 2,
      Marzo: 3,
      Abril: 4,
      Mayo: 5,
      Junio: 6,
      Julio: 7,
      Agosto: 8,
      Septiembre: 9,
      Octubre: 10,
      Noviembre: 11,
      Diciembre: 12,
    };

    return spanishMonths[monthName] || null;
  }

  function formatSheet() {
    var startColumnIndex = 2;
    var columnJump = 2;
    var lastColumn = summarySheet.getLastColumn();

    for (
      var columnIndex = startColumnIndex;
      columnIndex <= lastColumn;
      columnIndex += columnJump
    ) {
      summarySheet
        .getRange(1, columnIndex, summarySheet.getLastRow(), 1)
        .setFontFamily('Inconsolata');
      summarySheet
        .getRange(1, columnIndex, summarySheet.getLastRow(), 1)
        .setNumberFormat('$#,##0.00');
    }
  }
}
