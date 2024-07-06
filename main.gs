function onEdit(e) {
  getColumnGreenSum(4, 5);
  getColumnCreditSum(4,5);
  getColumnBlueSum(4,5);
  getColumnTotalSum(4,5);
}

function onOpen(e) {
  getColumnGreenSum(4, 5);
  getColumnCreditSum(4,5);
  getColumnBlueSum(4,5);
  getColumnTotalSum(4,5);
}


const green = '#93c47d'; //Light green
const red = '#e06666';  //Light red
const blue = '#6d9eeb'; //Light cornflower blue 1


//Credit usage for personal use
function getColumnGreenSum(column, output) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange(1, column, sheet.getLastRow(), 1);
  var backgroundRange = sheet.getRange(1, column - 1, sheet.getLastRow(), 1);
  var data = dataRange.getValues();
  var backgrounds = backgroundRange.getBackgrounds();

  var greenSum = 0.0;

  for (var i = 0; i < data.length; i++) {
    // Additional NaN check would be handy if the end user messes up
    if (backgrounds[i][0] == green) {
      greenSum += parseFloat(data[i][0]) || 0;
    }

  }

  sheet.getRange(2, output).setValue(greenSum);
}

//Total Credit Usage
function getColumnCreditSum(column, output) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange(1, column, sheet.getLastRow(), 1);
  var backgroundRange = sheet.getRange(1, column - 1, sheet.getLastRow(), 1);
  var data = dataRange.getValues();
  var backgrounds = backgroundRange.getBackgrounds();


  var creditSum = 0.0;

  for (var i = 0; i < data.length; i++) {
    // Additional NaN check would be handy if the end user messes up
    if (backgrounds[i][0] == green || backgrounds[i][0] == red) {
      creditSum += parseFloat(data[i][0]) || 0;
    }

  }

  sheet.getRange(3, output).setValue(creditSum);
}

//Blue for cash
function getColumnBlueSum(column, output) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange(1, column, sheet.getLastRow(), 1);
  var backgroundRange = sheet.getRange(1, column - 1, sheet.getLastRow(), 1);
  var data = dataRange.getValues();
  var backgrounds = backgroundRange.getBackgrounds();

  var blueSum = 0.0;

  for (var i = 0; i < data.length; i++) {
    // Additional NaN check would be handy if the end user messes up
    if (backgrounds[i][0] == blue) {
      blueSum += parseFloat(data[i][0]) || 0;
    }

  }

  sheet.getRange(4, output).setValue(blueSum);
}

//All expenses
function getColumnTotalSum(column, output) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getRange(1, column, sheet.getLastRow(), 1);
  var backgroundRange = sheet.getRange(1, column - 1, sheet.getLastRow(), 1);
  var data = dataRange.getValues();
  var backgrounds = backgroundRange.getBackgrounds();

  var totalSum = 0.0;

  for (var i = 0; i < data.length; i++) {
    // Additional NaN check would be handy if the end user messes up
    
    totalSum += parseFloat(data[i][0]) || 0;
    

  }
  sheet.getRange(5, output).setValue(totalSum);
}
