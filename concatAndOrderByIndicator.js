function concatAndOrderByIndicator_old() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getActiveSheet();

  const hoja1 = ss.getSheetByName("BD_componentes");
  const hoja2 = ss.getSheetByName("BD_indicadores");

  // Mapeo: nombre final → nombre en hoja1 y hoja2
  const columnasMapeo = {
    "municipio": { hoja1: "municipio", hoja2: "municipio" },
    "id_indicador": { hoja1: "id_indicador", hoja2: "id_indicador" },
    "eje": { hoja1: "eje", hoja2: "eje" },
    "tema": { hoja1: "tema", hoja2: "tema" },
    "nombre": { hoja1: "nombre_componente", hoja2: "nombre_indicador" },
    "codigo": { hoja1: "codigo_componente", hoja2: "code"},
    "tipo de dato": { hoja1: null, hoja2: null } // Esta se agregará manualmente
  };

  const columnasDestino = Object.keys(columnasMapeo); // Mantiene el orden deseado

  // Función auxiliar para extraer datos alineados por nombre
  function extraerDatosMapeados(hoja, fuente, etiquetaTipo) {
  const datos = hoja.getDataRange().getValues();
  const encabezado = datos[0];
  const cuerpo = datos.slice(1);

  const indices = columnasDestino
    .filter(col => col !== "tipo de dato")
    .map(col => {
      const nombreCol = columnasMapeo[col][fuente];
      const idx = encabezado.indexOf(nombreCol);
      if (idx === -1) throw new Error(`Columna "${nombreCol}" no encontrada en hoja "${hoja.getName()}"`);
      return idx;
    });

  return cuerpo.map(fila => {
    const datosFila = indices.map(i => fila[i]);
    datosFila.push(etiquetaTipo); // Añade "tipo de dato"
    return datosFila;
  });
}

  // Obtener datos de ambas hojas
  const datos1 = extraerDatosMapeados(hoja1, "hoja1", "Componente");
  const datos2 = extraerDatosMapeados(hoja2, "hoja2", "Indicador");

  const datosCombinados = datos1.concat(datos2);

  // Índices de las columnas por nombre, basado en el orden del mapeo
  const indiceId = columnasDestino.indexOf("id_indicador");
  const indiceMunicipio = columnasDestino.indexOf("municipio");

  // Ordenar por municipio (alfabético), luego por id_indicator (numérico ascendente)
datosCombinados.sort((a, b) => {
  const compMun = a[indiceMunicipio].toString().localeCompare(b[indiceMunicipio].toString(), 'es');
  if (compMun !== 0) return compMun;
  return Number(a[indiceId]) - Number(b[indiceId]);
});

  // Escribir resultados
  //hojaDestino.clear();
  hojaDestino.getRange(1, 1, 1, columnasDestino.length).setValues([columnasDestino]);
  hojaDestino.getRange(2, 1, datosCombinados.length, columnasDestino.length).setValues(datosCombinados);
}