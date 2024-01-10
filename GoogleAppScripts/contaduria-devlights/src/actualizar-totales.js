function updateTotales() {
  var ss = SpreadsheetApp.getActive();

  var totalsSheet = ss.getSheetByName('TOTALES');

  // get all sheets but omitted
  var omitSheets = ['SETTINGS', 'RESUMEN', 'TOTALES'];
  var allSheets = ss.getSheets();
  var result = [];
  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    if (!omitSheets.includes(sheet.getName())) {
      result.push(sheet);
    }
  }
  var sheets = result;

  var transferSum = 0;
  var payonnerSum = 0;
  var utopiaSum = 0;
  var wiseSum = 0;
  var criptoSum = 0;
  var pesosSum = 0;
  var usdSum = 0; 

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];

    var transfer = sheet.getRange('B16').getValue();
    var payonner = sheet.getRange('B17').getValue(); 
    var utopia = sheet.getRange('B18').getValue();
    var wise = sheet.getRange('B19').getValue(); 
    var cripto = sheet.getRange('B20').getValue();
    var pesos = sheet.getRange('B21').getValue(); 
    var usd = sheet.getRange('B22').getValue();

    if(typeof transfer === "number") {
      transferSum += transfer;
    }
    if(typeof payonner === "number") {
      payonnerSum += payonner;
    }
    if(typeof utopia === "number") {
      utopiaSum += utopia;
    }
    if(typeof wise === "number") {
      wiseSum += wise;
    }
    if(typeof cripto === "number") {
      criptoSum += cripto;
    }
    if(typeof pesos === "number") {
      pesosSum += pesos;
    }
    if(typeof usd === "number") {
      usdSum += usd;
    }
  }

  totalsSheet.getRange('B2').setValue(transferSum);
  totalsSheet.getRange('B3').setValue(payonnerSum);
  totalsSheet.getRange('B4').setValue(utopiaSum);
  totalsSheet.getRange('B5').setValue(wiseSum);
  totalsSheet.getRange('B6').setValue(criptoSum);
  totalsSheet.getRange('B7').setValue(pesosSum);
  totalsSheet.getRange('B8').setValue(usdSum);

  // write total
  var total = transferSum + payonnerSum + utopiaSum + wiseSum + criptoSum + pesosSum + usdSum;
  totalsSheet.getRange('B9').setValue(total);
}