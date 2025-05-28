function replaceURLs() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NombreDeTuHoja"); // Cambia el nombre de la hoja
  const rango = hoja.getRange("A2:Z100"); // Cambia el rango según lo necesites
  const datos = rango.getValues();

  const urlRegex = /(https?:\/\/[^\s]+)/i;

  for (let i = 0; i < datos.length; i++) {
    for (let j = 0; j < datos[i].length; j++) {
      if (typeof datos[i][j] === 'string' && urlRegex.test(datos[i][j])) {
        datos[i][j] = 1;
      }
    }
  }

  rango.setValues(datos);
}
