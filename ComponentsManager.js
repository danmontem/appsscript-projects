/**
 * Components Dataset Management Functions
 */

function showGenerateComponentsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('GenerateComponentsDialog')
    .setWidth(400)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Municipality Components Dataset');
}

function generateDataset(params) {
  // Get current spreadsheet for target sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(params.targetSheet);
  
  if (!targetSheet) {
    return {success: false, message: `Target sheet "${params.targetSheet}" not found`};
  }
  
  // Access the source document
  let sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(params.sourceDocumentId);
  } catch (e) {
    return {success: false, message: `Cannot access source document. Please check the ID and sharing permissions. Error: ${e.message}`};
  }
  
  // Get the source sheets
  const indicadoresComponentesSheet = sourceSpreadsheet.getSheetByName("Indicadores_Componentes");
  const bancoIndicadoresSheet = sourceSpreadsheet.getSheetByName("Banco_indicadores_ACV");
  
  if (!indicadoresComponentesSheet) {
    return {success: false, message: 'Sheet "Indicadores_Componentes" not found in source document'};
  }
  
  if (!bancoIndicadoresSheet) {
    return {success: false, message: 'Sheet "Banco_indicadores_ACV" not found in source document'};
  }
  
  // Get the data from both source sheets
  const componentesData = indicadoresComponentesSheet.getDataRange().getValues();
  const componentesHeaders = componentesData[0];
  const componentesRows = componentesData.slice(1);
  
  const indicadoresData = bancoIndicadoresSheet.getDataRange().getValues();
  const indicadoresHeaders = indicadoresData[0];
  const indicadoresRows = indicadoresData.slice(1);
  
  // Find column indices in the componentes data
  const compIndicadorIdCol = componentesHeaders.indexOf("ID Indicador");
  const compComponenteIdCol = componentesHeaders.indexOf("ID Componente");
  const compNombreComponenteCol = componentesHeaders.indexOf("Componente");
  const compSolicitudCol = componentesHeaders.indexOf("Solicitud de información");
  const compCodigoCol = componentesHeaders.indexOf("Código de componente");
  
  // Find column indices in the indicadores data
  const indIndicadorIdCol = indicadoresHeaders.indexOf("ID Indicador");
  const indEjeCol = indicadoresHeaders.indexOf("Eje de evaluación");
  const indTemaCol = indicadoresHeaders.indexOf("Tema");
  const indIndicadorCol = indicadoresHeaders.indexOf("Indicador");
  
  // Create the output headers - now with 9 fixed columns to include both ID and code
  const fixedHeaders = [
    "municipio", 
    "id_indicador", 
    "id_componente",
    "codigo_componente", // Added this column for the component code
    "eje", 
    "tema", 
    "nombre_indicador", 
    "nombre_componente", 
    "solicitud_de_informacion"
  ];
  
  // Month headers - these can be extended in the future
  const monthHeaders = [
    "octubre 2023", 
    "noviembre 2023", 
    "diciembre 2023", 
    "enero 2024", 
    "febrero 2024", 
    "marzo 2024", 
    "abril 2024", 
    "mayo 2024", 
    "junio 2024", 
    "julio 2024", 
    "agosto 2024", 
    "septiembre 2024", 
    "octubre 2024", 
    "noviembre 2024", 
    "diciembre 2024",
    "enero 2025",
    "febrero 2025",
    "marzo 2025"
  ];
  
  // Combine fixed headers and month headers
  const outputHeaders = fixedHeaders.concat(monthHeaders);
  
  // Create a lookup map for indicator data
  const indicadoresMap = {};
  indicadoresRows.forEach(row => {
    const indicadorId = row[indIndicadorIdCol];
    indicadoresMap[indicadorId] = {
      eje: row[indEjeCol],
      tema: row[indTemaCol],
      indicador: row[indIndicadorCol]
    };
  });
  
  // Create result array
  const result = [outputHeaders];
  
  // Store configuration for future refreshes
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SOURCE_DOCUMENT_ID', params.sourceDocumentId);
  scriptProperties.setProperty('TARGET_SHEET', params.targetSheet);
  scriptProperties.setProperty('MUNICIPALITIES', JSON.stringify(params.municipalities));
  scriptProperties.setProperty('FIXED_HEADERS_COUNT', fixedHeaders.length.toString());
  
  // Generate rows for all municipalities and all components
  params.municipalities.forEach(municipio => {
    componentesRows.forEach(componenteRow => {
      const indicadorId = componenteRow[compIndicadorIdCol];
      const componenteId = componenteRow[compComponenteIdCol];
      const nombreComponente = componenteRow[compNombreComponenteCol];
      const solicitudInfo = componenteRow[compSolicitudCol];
      const codigoComponente = componenteRow[compCodigoCol]; // Extract the component code
      
      // Handle components with ID Indicador 0 (general components)
      if (indicadorId === 0 || indicadorId === "0") {
        const newRow = [];
        
        // Add fixed columns for general components
        newRow.push(municipio); // municipio
        newRow.push(indicadorId); // id_indicador
        newRow.push(componenteId); // id_componente - keep original ID
        newRow.push(codigoComponente || ""); // codigo_componente - add code as a separate column
        newRow.push("General"); // eje - use "General" for these components
        newRow.push("General"); // tema - use "General" for these components
        newRow.push("Sin indicador"); // nombre_indicador
        newRow.push(nombreComponente); // nombre_componente
        newRow.push(solicitudInfo); // solicitud_de_informacion
        
        // Add empty values for month columns
        monthHeaders.forEach(() => {
          newRow.push("");
        });
        
        result.push(newRow);
        return; // Continue to next component
      }
      
      // For components with valid indicators, keep the existing logic
      // Look up indicator data
      const indicadorData = indicadoresMap[indicadorId] || {};
      
      if (!indicadorData.eje) {
        // Skip this row if indicator not found
        return;
      }
      
      const newRow = [];
      
      // Add fixed columns
      newRow.push(municipio); // municipio
      newRow.push(indicadorId); // id_indicador
      newRow.push(componenteId); // id_componente - keep original ID
      newRow.push(codigoComponente || ""); // codigo_componente - add code as a separate column
      newRow.push(indicadorData.eje); // eje
      newRow.push(indicadorData.tema); // tema
      newRow.push(indicadorData.indicador); // nombre_indicador
      newRow.push(nombreComponente); // nombre_componente
      newRow.push(solicitudInfo); // solicitud_de_informacion
      
      // Add empty values for month columns
      monthHeaders.forEach(() => {
        newRow.push("");
      });
      
      result.push(newRow);
    });
  });
  
  // Clear target sheet and write data
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  
  return {
    success: true, 
    message: `Dataset created in sheet "${params.targetSheet}" with ${result.length - 1} rows for ${params.municipalities.length} municipalities. Month values can be edited manually.`
  };
}

function refreshDatasetData() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Get stored properties
  const sourceDocumentId = scriptProperties.getProperty('SOURCE_DOCUMENT_ID');
  const targetSheetName = scriptProperties.getProperty('TARGET_SHEET');
  const municipalities = JSON.parse(scriptProperties.getProperty('MUNICIPALITIES') || '[]');
  const fixedHeadersCount = parseInt(scriptProperties.getProperty('FIXED_HEADERS_COUNT') || '9'); // Update to 9 for the new fixed headers count
  
  if (!sourceDocumentId || !targetSheetName || !municipalities.length) {
    ui.alert('Error', 'No dataset configuration found. Please generate a dataset first.', ui.ButtonSet.OK);
    return;
  }
  
  // Get current spreadsheet for target sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    ui.alert('Error', 'Target sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  // Access the source document
  let sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocumentId);
  } catch (e) {
    ui.alert('Error', 'Cannot access source document. Please check if it still exists and is shared with you.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the source sheets
  const indicadoresComponentesSheet = sourceSpreadsheet.getSheetByName("Indicadores_Componentes");
  const bancoIndicadoresSheet = sourceSpreadsheet.getSheetByName("Banco_indicadores_ACV");
  
  if (!indicadoresComponentesSheet || !bancoIndicadoresSheet) {
    ui.alert('Error', 'Required sheets not found in source document.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the current target data including all columns
  const targetData = targetSheet.getDataRange().getValues();
  const targetHeaders = targetData[0];
  
  // Split the headers into fixed and dynamic parts
  const fixedTargetHeaders = targetHeaders.slice(0, fixedHeadersCount);
  const dynamicTargetHeaders = targetHeaders.slice(fixedHeadersCount);
  
  // Get a map of all current data to preserve
  // CRITICAL UPDATE: Now using a more detailed key that includes codigo_componente
  // KEY FORMAT: municipio_indicadorId_componentId_codigoComponente
  const preservedDataMap = {};
  
  // Skip header row
  for (let i = 1; i < targetData.length; i++) {
    const row = targetData[i];
    const municipio = row[0];
    const indicadorId = row[1];
    const componentId = row[2];
    const codigoComponente = row[3]; // Include the codigo_componente in the key
    
    // Create a composite key that uniquely identifies each component
    // We'll use multiple key formats to maximize chances of matching
    const keyWithCode = `${municipio}_${codigoComponente}`; // Primary key using codigo_componente
    const keyWithId = `${municipio}_${indicadorId}_${componentId}`; // Backup key using IDs
    
    // Store the full row data (except headers) for preservation
    // We're now storing more information to help with data alignment
    preservedDataMap[keyWithCode] = {
      municipio: municipio,
      indicadorId: indicadorId,
      componentId: componentId,
      codigoComponente: codigoComponente,
      dynamicValues: row.slice(fixedHeadersCount) // Month values
    };
    
    // Also store it with the ID-based key as a fallback
    if (!preservedDataMap[keyWithId]) {
      preservedDataMap[keyWithId] = {
        municipio: municipio,
        indicadorId: indicadorId,
        componentId: componentId,
        codigoComponente: codigoComponente,
        dynamicValues: row.slice(fixedHeadersCount) // Month values
      };
    }
  }
  
  // Get the component data
  const componentesData = indicadoresComponentesSheet.getDataRange().getValues();
  const componentesHeaders = componentesData[0];
  const componentesRows = componentesData.slice(1);
  
  // Get the indicator data
  const indicadoresData = bancoIndicadoresSheet.getDataRange().getValues();
  const indicadoresHeaders = indicadoresData[0];
  const indicadoresRows = indicadoresData.slice(1);
  
  // Find column indices in the componentes data
  const compIndicadorIdCol = componentesHeaders.indexOf("ID Indicador");
  const compComponenteIdCol = componentesHeaders.indexOf("ID Componente");
  const compNombreComponenteCol = componentesHeaders.indexOf("Componente");
  const compSolicitudCol = componentesHeaders.indexOf("Solicitud de información");
  const compCodigoCol = componentesHeaders.indexOf("Código de componente");
  
  // Find column indices in the indicadores data
  const indIndicadorIdCol = indicadoresHeaders.indexOf("ID Indicador");
  const indEjeCol = indicadoresHeaders.indexOf("Eje de evaluación");
  const indTemaCol = indicadoresHeaders.indexOf("Tema");
  const indIndicadorCol = indicadoresHeaders.indexOf("Indicador");
  
  // Create a lookup map for indicator data
  const indicadoresMap = {};
  indicadoresRows.forEach(row => {
    const indicadorId = row[indIndicadorIdCol];
    indicadoresMap[indicadorId] = {
      eje: row[indEjeCol],
      tema: row[indTemaCol],
      indicador: row[indIndicadorCol]
    };
  });
  
  // Keep the headers from target sheet
  const result = [targetHeaders];
  
  // Regenerate the dataset while preserving month values
  municipalities.forEach(municipio => {
    componentesRows.forEach(componenteRow => {
      const indicadorId = componenteRow[compIndicadorIdCol];
      const componentId = componenteRow[compComponenteIdCol];
      const nombreComponente = componenteRow[compNombreComponenteCol];
      const solicitudInfo = componenteRow[compSolicitudCol];
      const codigoComponente = componenteRow[compCodigoCol]; // Extract the component code
      
      let newRow = [];
      let preservedData = null;
      
      // Try to find matching preserved data using codigo_componente first (most reliable)
      const keyWithCode = `${municipio}_${codigoComponente}`;
      const keyWithId = `${municipio}_${indicadorId}_${componentId}`;
      
      // This is the key improvement - try multiple matching strategies to find the right data
      if (preservedDataMap[keyWithCode]) {
        preservedData = preservedDataMap[keyWithCode];
      } else if (preservedDataMap[keyWithId]) {
        preservedData = preservedDataMap[keyWithId];
      }
      
      // Handle components with ID Indicador 0 (general components)
      if (indicadorId === 0 || indicadorId === "0") {
        // Add fixed columns for general components
        newRow.push(municipio); // municipio
        newRow.push(indicadorId); // id_indicador
        newRow.push(componentId); // id_componente - keep original ID
        newRow.push(codigoComponente || ""); // codigo_componente - add code as a separate column
        newRow.push("General"); // eje - use "General" for these components
        newRow.push("General"); // tema - use "General" for these components
        newRow.push("Sin indicador"); // nombre_indicador
        newRow.push(nombreComponente); // nombre_componente
        newRow.push(solicitudInfo); // solicitud_de_informacion
      } 
      // Handle regular components with indicators
      else {
        // Look up indicator data
        const indicadorData = indicadoresMap[indicadorId] || {};
        
        if (!indicadorData.eje) {
          // Skip this row if indicator not found
          return;
        }
        
        // Add fixed columns
        newRow.push(municipio); // municipio
        newRow.push(indicadorId); // id_indicador
        newRow.push(componentId); // id_componente - keep original ID
        newRow.push(codigoComponente || ""); // codigo_componente - add code as a separate column
        newRow.push(indicadorData.eje); // eje
        newRow.push(indicadorData.tema); // tema
        newRow.push(indicadorData.indicador); // nombre_indicador
        newRow.push(nombreComponente); // nombre_componente
        newRow.push(solicitudInfo); // solicitud_de_informacion
      }
      
      // Add dynamic values - preserve existing month values if available
      if (preservedData && preservedData.dynamicValues && preservedData.dynamicValues.length > 0) {
        // Add all the preserved month values
        preservedData.dynamicValues.forEach(value => {
          newRow.push(value);
        });
        
        // Mark this data as used to avoid duplicates
        if (codigoComponente) {
          delete preservedDataMap[keyWithCode];
        }
        delete preservedDataMap[keyWithId];
      } else {
        // If no preserved values, add empty values
        dynamicTargetHeaders.forEach(() => {
          newRow.push("");
        });
      }
      
      result.push(newRow);
    });
  });
  
  // Update target sheet
  targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  targetSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  
  ui.alert('Success', 'Dataset refreshed successfully while preserving month values.', ui.ButtonSet.OK);
}

/**
 * Creates a backup copy of the current dataset with timestamp
 */
function backupCurrentDataset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const targetSheetName = scriptProperties.getProperty('TARGET_SHEET');
  
  if (!targetSheetName) {
    SpreadsheetApp.getUi().alert('Error', 'No target sheet configured.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Target sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Create backup with timestamp
  const timestamp = new Date().toISOString().slice(0,16).replace(/:/g, '-');
  const backupName = `${targetSheetName}_backup_${timestamp}`;
  
  try {
    const backupSheet = targetSheet.copyTo(ss);
    backupSheet.setName(backupName);
    
    SpreadsheetApp.getUi().alert('Success', `Backup created: "${backupName}"`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', `Failed to create backup: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}