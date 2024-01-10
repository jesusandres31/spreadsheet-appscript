function updateResumen() {
  var ss = SpreadsheetApp.getActive();

  var totalsSheet = ss.getSheetByName('RESUMEN');

  // omitted sheets
  var omitSheets = ['SETTINGS', 'RESUMEN', 'TOTALES'];

  var sheets = ss.getSheets();

  var names = [];
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getSheetName();

    if (!omitSheets.includes(sheetName)) {
      names.push([sheetName]);
    }
  }

  // write names in A column
  if (names.length > 0) {
    totalsSheet.getRange(2, 1, names.length, 1).setValues(names);
  }

  for (var i = 0; i < names.length; i++) {
    var empName = names[i][0];

    // write formulas
    totalsSheet.getRange(i + 2, 2).setFormula(`='${empName}'!B16`);
    totalsSheet.getRange(i + 2, 3).setFormula(`='${empName}'!B17`);
    totalsSheet.getRange(i + 2, 4).setFormula(`='${empName}'!B18`);
    totalsSheet.getRange(i + 2, 5).setFormula(`='${empName}'!B19`);
    totalsSheet.getRange(i + 2, 6).setFormula(`='${empName}'!B20`);
    totalsSheet.getRange(i + 2, 7).setFormula(`='${empName}'!B21`);
    totalsSheet.getRange(i + 2, 8).setFormula(`='${empName}'!B22`);
  }
}
