function updateSecondArgument() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const columnaFormulas = 10; // Columna I
  const filaInicio = 2;
  const filaFin = 1404;
  const numFilas = filaFin - filaInicio + 1;

  // Leer todas las fórmulas del rango de una vez
  const rango = hoja.getRange(filaInicio, columnaFormulas, numFilas, 1);
  const formulas = rango.getFormulas();

  // Actualizar las fórmulas en memoria
  for (let i = 0; i < formulas.length; i++) {
    const formula = formulas[i][0];
    const filaActual = filaInicio + i;
    if (formula && formula.includes('CALCULATE_INDICATOR')) {
      const nuevaReferencia = `$A${filaActual}`;
      formulas[i][0] = formula.replace(/\$A\d+/, nuevaReferencia);
    }
  }

  // Escribir todas las fórmulas de vuelta en una sola operación
  rango.setFormulas(formulas);
}
