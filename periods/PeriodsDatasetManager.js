/**
 * PeriodsDatasetManager.js
 * Main periods dataset creation and management
 * Dependencies: CommonHelpers.js
 */

/**
 * Step 1: Create evaluation periods dataset (components + indicators with period columns)
 */
function createEvaluationPeriodsDataset() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask user to select the combined dataset sheet (source)
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  const sourceResult = ui.prompt(
    'Select Source Combined Dataset',
    `Available sheets: ${sheetsText}\n\nEnter the name of your COMBINED dataset sheet (source):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sourceResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceName = sourceResult.getResponseText().trim();
  const sourceSheet = ss.getSheetByName(sourceName);
  
  if (!sourceSheet) {
    ui.alert('Error', `Sheet "${sourceName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Ask for destination sheet name
  const destResult = ui.prompt(
    'Evaluation Periods Dataset Name',
    'Enter name for the new evaluation periods dataset sheet (e.g., "Periodos_Evaluacion"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (destResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const destName = destResult.getResponseText().trim();
  
  // Check if sheet already exists
  let destSheet = ss.getSheetByName(destName);
  if (destSheet) {
    const overwriteResult = ui.alert(
      'Sheet Exists',
      `Sheet "${destName}" already exists. Overwrite it?`,
      ui.ButtonSet.YES_NO
    );
    
    if (overwriteResult !== ui.Button.YES) {
      return;
    }
    
    destSheet.clear();
  } else {
    destSheet = ss.insertSheet(destName);
  }
  
  // Get source data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  // Create headers for the new dataset
  const newHeaders = [
    'municipio',
    'id_indicador', 
    'eje',
    'tema',
    'tipo de dato',
    'nombre',
    'codigo'
  ].concat(EVALUATION_PERIOD_COLUMNS);
  
  // Find source column indices using CommonHelpers
  const columnIndices = getColumnIndices(sourceHeaders, [
    'municipio', 'id_indicador', 'eje', 'tema', 'tipo de dato', 'nombre', 'codigo'
  ]);
  
  if (!columnIndices.valid) {
    ui.alert('Error', 'Required columns not found in source dataset.', ui.ButtonSet.OK);
    return;
  }
  
  // Set up the destination sheet
  destSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Apply header formatting
  const headerRange = destSheet.getRange(1, 1, 1, newHeaders.length);
  headerRange.setFontWeight('bold');
  
  // Color code different sections
  destSheet.getRange(1, 1, 1, 7).setBackground('#e8f4fd'); // Core columns
  destSheet.getRange(1, 8, 1, EVALUATION_PERIOD_COLUMNS.length).setBackground('#fff2cc'); // Period columns
  
  // Copy core data with PROPER handling of id_indicador = 0
  const resultData = [];
  
  sourceRows.forEach((row, index) => {
    const newRow = [
      getValueOrEmpty(row[columnIndices.municipio]),
      getValueOrZero(row[columnIndices.id_indicador]),  // Fixed: Use special handler for id_indicador
      getValueOrEmpty(row[columnIndices.eje]),
      getValueOrEmpty(row[columnIndices.tema]),
      getValueOrEmpty(row[columnIndices.tipo_de_dato]),  // FIXED: Use correct property name with underscore
      getValueOrEmpty(row[columnIndices.nombre]),
      getValueOrEmpty(row[columnIndices.codigo])
    ];
    
    // Add empty evaluation period columns
    EVALUATION_PERIOD_COLUMNS.forEach(() => {
      newRow.push('');
    });
    
    resultData.push(newRow);
    
    // ENHANCED DEBUGGING: Log first few rows
    if (index < 3) {
      console.log(`Row ${index + 2} processed: municipio=${newRow[0]}, id_indicador=${newRow[1]}, tipo="${newRow[4]}", codigo=${newRow[6]}`);
    }
  });
  
  // Write data to destination
  if (resultData.length > 0) {
    destSheet.getRange(2, 1, resultData.length, newHeaders.length).setValues(resultData);
  }
  
  // Auto-resize columns
  destSheet.autoResizeColumns(1, newHeaders.length);
  
  // Add instructions note
  destSheet.getRange('A1').setNote(
    'Evaluation Periods Dataset\n\n' +
    'Next steps:\n' +
    '1. Run "Map Periods from External Table" to populate period columns\n' +
    '2. Manually set periods for components with id_indicador = 0\n' +
    '3. Use this dataset for generating calculations'
  );
  
  // Count components with id_indicador = 0 for reporting
  const componentsWithZero = resultData.filter(row => row[1] === 0).length;
  
  // Success message
  ui.alert(
    'Evaluation Periods Dataset Created!',
    `Dataset created successfully!\n\n` +
    `• Sheet: "${destName}"\n` +
    `• Total rows: ${resultData.length}\n` +
    `• Components with id_indicador = 0: ${componentsWithZero}\n` +
    `• Evaluation period columns: ${EVALUATION_PERIOD_COLUMNS.length}\n\n` +
    `✓ Fixed: Components with id_indicador = 0 are properly preserved\n\n` +
    `Next step: Run "Map Periods from External Table" to populate the period columns.`,
    ui.ButtonSet.OK
  );
}

/**
 * Automated periods dataset recreation (no UI prompts)
 */
function recreatePeriodsDatasetAutomated(sourceSheet, destName, ss) {
  console.log(`Recreating periods dataset: ${destName}`);
  
  // Get source data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  console.log(`Source data: ${sourceRows.length} rows, headers: ${sourceHeaders.join(', ')}`);
  
  // Create headers for the new dataset
  const newHeaders = [
    'municipio',
    'id_indicador', 
    'eje',
    'tema',
    'tipo de dato',
    'nombre',
    'codigo'
  ].concat(EVALUATION_PERIOD_COLUMNS);
  
  // Find source column indices using CommonHelpers
  const columnIndices = getColumnIndices(sourceHeaders, [
    'municipio', 'id_indicador', 'eje', 'tema', 'tipo de dato', 'nombre', 'codigo'
  ]);
  
  console.log('Column indices:', columnIndices);
  
  if (!columnIndices.valid) {
    throw new Error(`Required columns not found in source dataset "${sourceSheet.getName()}". Available columns: ${sourceHeaders.join(', ')}`);
  }
  
  // Create or clear destination sheet
  let destSheet = ss.getSheetByName(destName);
  if (destSheet) {
    console.log(`Clearing existing sheet: ${destName}`);
    destSheet.clear();
  } else {
    console.log(`Creating new sheet: ${destName}`);
    destSheet = ss.insertSheet(destName);
  }
  
  // Set up the destination sheet
  destSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Apply header formatting
  const headerRange = destSheet.getRange(1, 1, 1, newHeaders.length);
  headerRange.setFontWeight('bold');
  destSheet.getRange(1, 1, 1, 7).setBackground('#e8f4fd'); // Core columns
  destSheet.getRange(1, 8, 1, EVALUATION_PERIOD_COLUMNS.length).setBackground('#fff2cc'); // Period columns
  
  // Copy core data
  const resultData = [];
  sourceRows.forEach((row, index) => {
    const newRow = [
      getValueOrEmpty(row[columnIndices.municipio]),
      getValueOrZero(row[columnIndices.id_indicador]),
      getValueOrEmpty(row[columnIndices.eje]),
      getValueOrEmpty(row[columnIndices.tema]),
      getValueOrEmpty(row[columnIndices.tipo_de_dato]),
      getValueOrEmpty(row[columnIndices.nombre]),
      getValueOrEmpty(row[columnIndices.codigo])
    ];
    
    // Add empty evaluation period columns
    EVALUATION_PERIOD_COLUMNS.forEach(() => {
      newRow.push('');
    });
    
    resultData.push(newRow);
    
    // Log first few rows for debugging
    if (index < 3) {
      console.log(`Row ${index + 2} processed: municipio=${newRow[0]}, id_indicador=${newRow[1]}, tipo=${newRow[4]}`);
    }
  });
  
  console.log(`Processed ${resultData.length} data rows`);
  
  // Write data to destination
  if (resultData.length > 0) {
    destSheet.getRange(2, 1, resultData.length, newHeaders.length).setValues(resultData);
    console.log(`Data written to ${destName}: ${resultData.length} rows x ${newHeaders.length} columns`);
  }
  
  // Auto-resize columns
  destSheet.autoResizeColumns(1, newHeaders.length);
  
  console.log(`Periods dataset ${destName} recreated successfully`);
}
