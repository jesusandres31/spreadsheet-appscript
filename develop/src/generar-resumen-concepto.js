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
    ROW: 7,
    COL: 4,
  };

  const BANKS_KEY_CHAR = '_'; 
  
  const BANKS_PESOS_STR = 'pesos';

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

  const TOTAL = {
    USD: "Total (USD):",
    AJUSTADO: "Total Ajustado (USD):"    
  }

  const COTIZACION = {
    sheetName: 'COTIZACION',
    range: 'A2:B13',
  };

  const EXCLUDED_CATEGORIES = [
    // Concepto Egreso:
    "Tranf letsbit cripto",
    "Tranf letsbit efectivo",
    "Transf cuentas propias",
    // Concepto Ingreso:
    "Transf cuentas propias"
  ]

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
    excluded: "#faeeed"
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
   * pesos banks
   */
  const pesosBanks = getPesosBanks();

  /**
   * set JUMP values
   */
  JUMP.ROW += conceptoEgresoData.length;

  /**
   * cotización periodo
   */
  const contizacionPeriodo = generarCotizacionPeriodo();

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
        setColTotal(j, titleCell, 1, TOTAL.USD);

        // Set column total ajustado
        setColTotal(j, titleCell, 2 ,TOTAL.AJUSTADO);
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
    fillTotal(i, k, tableTitle, colData, egresoData.total, 0);
    fillTotal(i, k, tableTitle, colData, egresoData.totalAjustado, 1);
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
        console.log(data[j][0])
        console.log(EXCLUDED_CATEGORIES.includes(data[j][0]))
        if (EXCLUDED_CATEGORIES.includes(data[j][0])) {
          currentCell.setBackground(STYLE.excluded);
        }
      }
    }
  }

  function fillTotal(i, k, tableTitle, columnName, total, jump) {
    if (tableTitle === columnName) {
      var cell = summarySheet.getRange(
        i * JUMP.ROW + TABLE_TITLES.length + jump,
        k * JUMP.COL + TABLE_TITLES.indexOf(columnName) + 1
      );
      var totalCell = cell.offset(conceptoEgresoData.length, 0);
      totalCell.setValue(total);

      totalCell.setFontWeight(STYLE.bold);
    }
  }

  function setColTotal(j, titleCell, jump, label) {
    if (TABLE_TITLES[j] === COLUMN.ConceptoEgreso) {
      var totalCell = titleCell.offset(conceptoEgresoData.length + jump, -1);
      totalCell.setValue(label);
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
      var totalAjustado = 0;

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
              
              // calc total ajustado
              if (!EXCLUDED_CATEGORIES.includes(category)) {
                totalAjustado += parseFloat(amount);
              }
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
          total = total / cotizacionMes;
          totalAjustado = totalAjustado / cotizacionMes;
        }
      }

      return { resultArray: resultArray, total: total, totalAjustado: totalAjustado };
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

  function getPesosBanks() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet(); 
    var sheets = sheet.getSheets(); 
    var pesosBankNames = []; 
    for (var i = 0; i < sheets.length; i++) {
      var nombreHoja = sheets[i].getName(); 
      if (nombreHoja.charAt(0) === BANKS_KEY_CHAR && sheets[i].getRange('A1').getValue().toLowerCase() === BANKS_PESOS_STR) {
        pesosBankNames.push(nombreHoja.substring(1));
      }
    }
    return pesosBankNames;
  } 

  function generarCotizacionPeriodo() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var hojaCotizacion = sheet.getSheetByName(COTIZACION.sheetName); 
    var datosTabla = hojaCotizacion.getRange(COTIZACION.range).getValues();
    var cotizacionPeriodo = {};
    
    for (var i = 0; i < datosTabla.length; i++) {
      var cotizacion = datosTabla[i][1];
      cotizacionPeriodo[i + 1] = cotizacion;
    }
    
    return cotizacionPeriodo;
  }


}
