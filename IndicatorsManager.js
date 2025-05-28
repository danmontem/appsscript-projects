/**
 * Indicators Dataset Management
 */

function showGenerateIndicatorsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('GenerateIndicatorsDialog')
    .setWidth(400)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Municipality Indicators Dataset');
}

/**
 * Generate indicators dataset from the Banco_indicadores_ACV sheet
 */
function generateIndicatorsDataset(params) {
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
  
  // Get the source sheet
  const bancoIndicadoresSheet = sourceSpreadsheet.getSheetByName("Banco_indicadores_ACV");
  
  if (!bancoIndicadoresSheet) {
    return {success: false, message: 'Sheet "Banco_indicadores_ACV" not found in source document'};
  }
  
  // Get the data from the source sheet
  const indicadoresData = bancoIndicadoresSheet.getDataRange().getValues();
  const indicadoresHeaders = indicadoresData[0];
  const indicadoresRows = indicadoresData.slice(1);
  
  // Find column indices in the indicadores data
  const indIndicadorIdCol = indicadoresHeaders.indexOf("ID Indicador");
  const indEjeCol = indicadoresHeaders.indexOf("Eje de evaluación");
  const indTemaCol = indicadoresHeaders.indexOf("Tema");
  const indIndicadorCol = indicadoresHeaders.indexOf("Indicador");
  
  // Create the output headers - only the fixed headers
  const outputHeaders = [
    "municipio", 
    "id_indicador", 
    "eje", 
    "tema", 
    "nombre_indicador"
  ];
  
  // Create result array
  const result = [outputHeaders];
  
  // Store configuration for future refreshes
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('INDICATORS_SOURCE_DOCUMENT_ID', params.sourceDocumentId);
  scriptProperties.setProperty('INDICATORS_TARGET_SHEET', params.targetSheet);
  scriptProperties.setProperty('INDICATORS_MUNICIPALITIES', JSON.stringify(params.municipalities));
  scriptProperties.setProperty('INDICATORS_FIXED_HEADERS_COUNT', outputHeaders.length.toString());
  
  // Generate rows for all municipalities and all indicators
  params.municipalities.forEach(municipio => {
    indicadoresRows.forEach(indicadorRow => {
      const indicadorId = indicadorRow[indIndicadorIdCol];
      const eje = indicadorRow[indEjeCol];
      const tema = indicadorRow[indTemaCol];
      const nombreIndicador = indicadorRow[indIndicadorCol];
      
      // Skip rows with invalid or missing indicator ID
      if (!indicadorId) {
        return;
      }
      
      const newRow = [];
      
    // Add fixed columns
      newRow.push(municipio); // municipio
      newRow.push(indicadorId); // id_indicador
      newRow.push(eje); // eje
      newRow.push(tema); // tema
      newRow.push(nombreIndicador); // nombre_indicador
      
      result.push(newRow);
    });
  });
  
  // Clear target sheet and write data
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  
  return {
    success: true, 
    message: `Indicators dataset created in sheet "${params.targetSheet}" with ${result.length - 1} rows for ${params.municipalities.length} municipalities. You can now add calculation columns manually.`
  };
}

/**
 * Refresh the indicators dataset while preserving all additional columns
 */
function refreshIndicatorsDataset() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Get stored properties
  const sourceDocumentId = scriptProperties.getProperty('INDICATORS_SOURCE_DOCUMENT_ID');
  const targetSheetName = scriptProperties.getProperty('INDICATORS_TARGET_SHEET');
  const municipalities = JSON.parse(scriptProperties.getProperty('INDICATORS_MUNICIPALITIES') || '[]');
  const fixedHeadersCount = parseInt(scriptProperties.getProperty('INDICATORS_FIXED_HEADERS_COUNT') || '5');
  
  if (!sourceDocumentId || !targetSheetName || !municipalities.length) {
    ui.alert('Error', 'No indicators dataset configuration found. Please generate a dataset first.', ui.ButtonSet.OK);
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
  
  // Get the source sheet
  const bancoIndicadoresSheet = sourceSpreadsheet.getSheetByName("Banco_indicadores_ACV");
  
  if (!bancoIndicadoresSheet) {
    ui.alert('Error', 'Sheet "Banco_indicadores_ACV" not found in source document.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the current target data including all columns
  const targetData = targetSheet.getDataRange().getValues();
  const targetHeaders = targetData[0];
  
  // Get the number of columns in the target sheet
  const totalColumns = targetHeaders.length;
  
  // Make sure we have at least the fixed columns
  if (totalColumns < fixedHeadersCount) {
    ui.alert('Error', 'Target sheet has been modified and is missing required columns.', ui.ButtonSet.OK);
    return;
  }
  
  // Get a map of all current data to preserve
  // KEY FORMAT: municipio_indicadorId
  const preservedDataMap = {};
  
  // Skip header row
  for (let i = 1; i < targetData.length; i++) {
    const row = targetData[i];
    const municipio = row[0];
    const indicadorId = row[1];
    
    // Create a composite key that uniquely identifies each indicator
    const key = `${municipio}_${indicadorId}`;
    
    // Store all values beyond the fixed headers for preservation
    preservedDataMap[key] = {
      dynamicValues: row.slice(fixedHeadersCount) // All additional columns
    };
  }
  
  // Get the indicator data
  const indicadoresData = bancoIndicadoresSheet.getDataRange().getValues();
  const indicadoresHeaders = indicadoresData[0];
  const indicadoresRows = indicadoresData.slice(1);
  
  // Find column indices in the indicadores data
  const indIndicadorIdCol = indicadoresHeaders.indexOf("ID Indicador");
  const indEjeCol = indicadoresHeaders.indexOf("Eje de evaluación");
  const indTemaCol = indicadoresHeaders.indexOf("Tema");
  const indIndicadorCol = indicadoresHeaders.indexOf("Indicador");
  
  // Keep the headers from target sheet
  const result = [targetHeaders];
  
  // Regenerate the dataset while preserving additional column values
  municipalities.forEach(municipio => {
    indicadoresRows.forEach(indicadorRow => {
      const indicadorId = indicadorRow[indIndicadorIdCol];
      const eje = indicadorRow[indEjeCol];
      const tema = indicadorRow[indTemaCol];
      const nombreIndicador = indicadorRow[indIndicadorCol];
      
      // Skip rows with invalid or missing indicator ID
      if (!indicadorId) {
        return;
      }
      
      const newRow = [];
      
      // Add fixed columns
      newRow.push(municipio); // municipio
      newRow.push(indicadorId); // id_indicador
      newRow.push(eje); // eje
      newRow.push(tema); // tema
      newRow.push(nombreIndicador); // nombre_indicador
      
      // Try to find matching preserved data
      const key = `${municipio}_${indicadorId}`;
      const preservedData = preservedDataMap[key];
      
      // Add dynamic values - preserve existing values if available
      if (preservedData && preservedData.dynamicValues) {
        // Add all the preserved additional column values
        preservedData.dynamicValues.forEach(value => {
          newRow.push(value);
        });
        
        // Mark this data as used to avoid duplicates
        delete preservedDataMap[key];
      } else {
        // If no preserved values, add empty values for all additional columns
        for (let i = 0; i < totalColumns - fixedHeadersCount; i++) {
          newRow.push("");
        }
      }
      
      result.push(newRow);
    });
  });
  
  // Update target sheet
  targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
  targetSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  
  ui.alert('Success', 'Indicators dataset refreshed successfully while preserving all additional columns.', ui.ButtonSet.OK);
}