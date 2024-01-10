function sumDebitsClientsMonth11() {
  for (var i = 0; i < banks.length; i++) {
    console.log(banks[i]);
    // Get the active spreadsheet
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      banks[i]
    );

    // Get data from the sheet
    var data = activeSheet.getDataRange().getValues();

    // Initialize the sum
    var sumDebits = 0;

    // Iterate over the data and sum those that meet the criteria
    for (var i = 1; i < data.length; i++) {
      // Start from the second row assuming headers in the first row
      var date = data[i][0];
      var category = data[i][2];
      var debit = data[i][3];

      var monthNumber = Utilities.formatDate(new Date(date), 'GMT-3', 'MM');

      // Check if the date contains "/11/" and the category is "sueldo"
      if (
        date &&
        monthNumber === '11' &&
        category &&
        category.toLowerCase() === 'sueldos'
      ) {
        if (!isNaN(debit)) {
          sumDebits += parseFloat(debit);
        }
      }
    }

    // Log the result or perform the desired action
    Logger.log(
      'The sum of debits for the "Salary" category in month 11 is: $' +
        sumDebits.toFixed(2)
    );
  }
}
