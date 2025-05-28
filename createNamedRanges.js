/** @OnlyCurrentDoc */

function createNamedRanges() {
  // Specify the columns for variable names and values
  var nameColumn = "H"; // Column with variable names
  var valueColumn = "AA"; // Column with values

  // Get the active spreadsheet and the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the data from the name and value columns
  var nameRange = sheet.getRange(nameColumn + "1:" + nameColumn + sheet.getLastRow());
  var valueRange = sheet.getRange(valueColumn + "1:" + valueColumn + sheet.getLastRow());
  var locationRange = sheet.getRange("A1:A" + sheet.getLastRow()); // Columna A para municipio

  var names = nameRange.getValues();
  var values = valueRange.getValues();
  var locations = locationRange.getValues();

  // Delete ALL preexisting named ranges
  //var namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  //namedRanges.forEach(function(namedRange) {
  //  namedRange.remove();
  //});

  // Loop through the names and create named ranges
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] !== "" && values[i][0] !== "") {
      if (locations[i][0] === "Apodaca") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_APO", range);
      }
      else if (locations[i][0] === "Escobedo") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_ESC", range);
      }
      else if (locations[i][0] === "García") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_GAR", range);
      }
      else if (locations[i][0] === "Guadalupe") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_GPE", range);
      }
      else if (locations[i][0] === "Juárez") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_JUA", range);
      }
      else if (locations[i][0] === "Monterrey") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_MTY", range);
      }
      else if (locations[i][0] === "San Nicolás") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_SNG", range);
      }
      else if (locations[i][0] === "San Pedro") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_SPG", range);
      }
      else if (locations[i][0] === "Santa Catarina") {
        var range = sheet.getRange(valueColumn + (i + 1));
        SpreadsheetApp.getActiveSpreadsheet().setNamedRange(names[i][0] + "_MAR25_STC", range);
      }
    }
  }
}