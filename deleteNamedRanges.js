function deleteNamedRangesInChunks() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var suffixes = ["_ESC", "_GAR", "_GPE", "JUA", "_SNG", "_SPG", "_STC"];
  var namedRanges = spreadsheet.getNamedRanges();
  
  var validRanges = namedRanges.filter(function(namedRange) {
    try {
      namedRange.getRange(); // Validar que el rango es válido
      return true; // Solo mantener los válidos
    } catch (e) {
      Logger.log("Named range roto: " + namedRange.getName());
      return false; // Filtrar los que están rotos
    }
  });

  // Proceder a eliminar los rangos válidos
  var count = 0;
  var maxToDelete = 500;

  for (var i = 0; i < validRanges.length && count < maxToDelete; i++) {
    var namedRange = validRanges[i];
    var name = namedRange.getName();

    if (suffixes.some(suffix => name.endsWith(suffix))) {
      try {
        spreadsheet.removeNamedRange(namedRange);
        Logger.log("Eliminado: " + name);
        count++;
      } catch (e) {
        Logger.log("Error eliminando '" + name + "': " + e.message);
      }
    }
  }

  Logger.log("Eliminados en este bloque: " + count);
  if (count >= maxToDelete) {
    Logger.log("Esperando 5 segundos antes de continuar...");
    Utilities.sleep(5000); // Espera 5 segundos antes de continuar con otro bloque
    deleteNamedRangesInChunks(); // Llama de nuevo la función para procesar el siguiente bloque
  }
}