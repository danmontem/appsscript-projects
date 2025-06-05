/**
 * PeriodsManager.js
 * Consolidated evaluation periods management system
 * Dependencies: CommonHelpers.js
 * 
 * This file combines functionality from:
 * - EvaluationPeriodsManager.js
 * - EvaluationPeriodsMapping.js
 * - EvaluationPeriodsHelpers.js
 * - EvaluationPeriodsSmartRefresh.js
 */

// =============================================
// MAIN PERIODS DATASET FUNCTIONS
// =============================================

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

// =============================================
// PERIODS MAPPING FUNCTIONS
// =============================================

/**
 * FLEXIBLE VERSION: Map periods from external table with user-selected sheets
 */
function mapPeriodsFromExternalTableFlexible() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('=== FLEXIBLE PERIODS MAPPING ===');
  
  // Step 1: Ask for the evaluation periods dataset sheet
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  const periodsDatasetResult = ui.prompt(
    'Select Evaluation Periods Dataset',
    `Available sheets: ${sheetsText}\n\nEnter the name of your evaluation periods dataset:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodsDatasetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const periodsDatasetName = periodsDatasetResult.getResponseText().trim();
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  
  if (!periodsDataset) {
    ui.alert('Error', `Sheet "${periodsDatasetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Step 2: Ask for external document ID
  const sourceDocResult = ui.prompt(
    'External Document ID',
    'Enter the external Google Sheets document ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sourceDocResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceDocId = sourceDocResult.getResponseText().trim();
  
  // Step 3: Try to access external document and get available sheets
  let sourceSpreadsheet, externalSheets;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
    externalSheets = sourceSpreadsheet.getSheets().map(sheet => sheet.getName());
  } catch (e) {
    ui.alert('Error', `Cannot access external document: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  // Step 4: Ask user to choose from available sheets in external document
  const externalSheetsText = externalSheets.join(', ');
  
  const externalSheetResult = ui.prompt(
    'Select External Sheet',
    `Available sheets in external document: ${externalSheetsText}\n\nEnter the sheet name containing the periods data:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (externalSheetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const externalSheetName = externalSheetResult.getResponseText().trim();
  const externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  
  if (!externalSheet) {
    ui.alert('Error', `Sheet "${externalSheetName}" not found in external document.`, ui.ButtonSet.OK);
    return;
  }
  
  // Step 5: Show confirmation
  const confirmResult = ui.alert(
    'Confirm Mapping Setup',
    `Periods Dataset: "${periodsDatasetName}"\n` +
    `External Document: ${sourceDocId}\n` +
    `External Sheet: "${externalSheetName}"\n\n` +
    `Proceed with mapping?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    return;
  }
  
  // Step 6: Perform the mapping
  performMappingWithBatches(externalSheet, periodsDataset, externalSheetName, periodsDatasetName);
}

/**
 * ULTRA FAST VERSION: Uses preset names for quick mapping
 * NOTE: Update these names to match your current sheet setup
 */
function mapPeriodsFromExternalTableUltraFast() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // CONFIGURABLE: Update these names to match your current setup
  const periodsDatasetName = 'Periodos_Evaluacion'; // Update if your periods sheet has a different name
  const sourceDocId = '10AkW5RTc-MGXh26V2sXkGxBQ0cINRVFq0-AXaSWjiHc'; // Update if using different external document
  const externalSheetName = 'Fechas_evaluación'; // Update if external sheet has different name
  
  console.log('=== ULTRA FAST PERIODS MAPPING ===');
  
  // Get sheets with better error handling
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  if (!periodsDataset) {
    const availableSheets = ss.getSheets().map(sheet => sheet.getName()).join(', ');
    ui.alert(
      'Periods Dataset Not Found', 
      `Sheet "${periodsDatasetName}" not found.\n\n` +
      `Available sheets: ${availableSheets}\n\n` +
      `Either:\n` +
      `1. Use "Map Periods (Flexible)" to choose sheets manually\n` +
      `2. Update the hardcoded name in this function\n` +
      `3. Create/rename your periods dataset to "${periodsDatasetName}"`,
      ui.ButtonSet.OK
    );
    return;
  }
  
  let sourceSpreadsheet, externalSheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
    externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  } catch (e) {
    ui.alert('Error', `Cannot access external document: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  if (!externalSheet) {
    ui.alert('Error', `Sheet "${externalSheetName}" not found in external document.`, ui.ButtonSet.OK);
    return;
  }
  
  // Perform the mapping
  performMappingWithBatches(externalSheet, periodsDataset, externalSheetName, periodsDatasetName);
}

/**
 * Enhanced batch processing with detailed progress and manual entry preservation
 */
function performMappingWithBatches(externalSheet, periodsDataset, externalSheetName, periodsDatasetName) {
  const ui = SpreadsheetApp.getUi();
  
  console.log('Building external lookup...');
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  const sourceIdIndex = findIdColumn(sourceHeaders);
  if (sourceIdIndex === -1) {
    ui.alert('Error', `Could not find ID Indicador column in "${externalSheetName}".`, ui.ButtonSet.OK);
    return;
  }
  
  // Find period columns using CommonHelpers
  const periodColumns = findPeriodColumns(sourceHeaders);
  if (periodColumns.length === 0) {
    ui.alert('Error', `No period columns found in "${externalSheetName}".`, ui.ButtonSet.OK);
    return;
  }
  
  console.log(`Found ${periodColumns.length} period columns`);
  
  // Build periods lookup
  const periodsLookup = new Map();
  sourceRows.forEach(row => {
    const id = normalizeIndicatorId(row[sourceIdIndex]);
    if (id !== null) {
      const periods = {};
      periodColumns.forEach(col => {
        const value = row[col.sourceIndex];
        if (value && value !== '' && value !== '...') {
          periods[col.target] = normalizePeriodText(value.toString());
        }
      });
      periodsLookup.set(id, periods);
    }
  });
  
  console.log(`External lookup built with ${periodsLookup.size} indicators`);
  
  // Process periods dataset in batches with manual entry preservation
  const batchSize = 50;
  const totalRows = periodsDataset.getLastRow() - 1;
  const periodsHeaders = periodsDataset.getRange(1, 1, 1, periodsDataset.getLastColumn()).getValues()[0];
  
  const idIndex = periodsHeaders.indexOf('id_indicador');
  const tipoIndex = periodsHeaders.indexOf('tipo de dato');
  
  let totalMapped = 0;
  let totalComponents = 0;
  let manualEntriesPreserved = 0;
  let newManualEntrySlots = 0;
  let processedRows = 0;
  
  console.log(`Processing ${totalRows} rows in batches of ${batchSize}...`);
  
  // Process in batches
  for (let startRow = 2; startRow <= totalRows + 1; startRow += batchSize) {
    const endRow = Math.min(startRow + batchSize - 1, totalRows + 1);
    const batchRows = endRow - startRow + 1;
    
    console.log(`Batch ${Math.ceil((startRow-1)/batchSize)}: rows ${startRow}-${endRow} (${batchRows} rows)`);
    
    const batchData = periodsDataset.getRange(startRow, 1, batchRows, periodsDataset.getLastColumn()).getValues();
    
    batchData.forEach((row, batchIndex) => {
      const actualRow = startRow + batchIndex;
      const id = row[idIndex];
      const tipo = row[tipoIndex];
      
      const normalizedId = normalizeIndicatorId(id);
      
      if (id === 0 || id === '0' || normalizedId === '0') {
        // Component without indicator - handle manual entries carefully
        totalComponents++;
        
        const firstPeriodColIndex = periodsHeaders.indexOf('Periodo_Base_EVA_Mayo');
        if (firstPeriodColIndex !== -1) {
          const currentCell = periodsDataset.getRange(actualRow, firstPeriodColIndex + 1);
          const currentValue = currentCell.getValue();
          
          // Check if there's already a manual entry
          if (currentValue && currentValue !== '' && currentValue !== 'Manual entry needed for component without indicator') {
            // PRESERVE EXISTING MANUAL ENTRY
            manualEntriesPreserved++;
            console.log(`Row ${actualRow}: Preserving manual entry "${currentValue}"`);
            // Don't change anything - keep the manually entered value
          } else {
            // NEW EMPTY SLOT - Mark for manual entry
            newManualEntrySlots++;
            currentCell.setBackground('#ffcccb');
            currentCell.setNote('Manual entry needed for component without indicator');
            console.log(`Row ${actualRow}: Marked new slot for manual entry`);
          }
        }
        
      } else if (normalizedId && periodsLookup.has(normalizedId)) {
        // Found matching periods - update with external data
        const periods = periodsLookup.get(normalizedId);
        
        periodColumns.forEach(col => {
          const targetColIndex = periodsHeaders.indexOf(col.target);
          if (targetColIndex !== -1 && periods[col.target]) {
            periodsDataset.getRange(actualRow, targetColIndex + 1).setValue(periods[col.target]);
            totalMapped++;
          }
        });
      }
      
      processedRows++;
    });
    
    // Show progress every batch
    const progress = Math.round((processedRows / totalRows) * 100);
    console.log(`Progress: ${progress}% (${processedRows}/${totalRows} rows) - Mapped: ${totalMapped}, Components: ${totalComponents}`);
    
    // Small pause every 500 rows
    if ((startRow - 2) % 500 === 0 && startRow > 2) {
      console.log('Taking short break to prevent timeout...');
      Utilities.sleep(50);
    }
  }
  
  console.log('=== MAPPING COMPLETE ===');
  console.log(`Final results: Mapped=${totalMapped}, Components=${totalComponents}, Processed=${processedRows}`);
  
  // Auto-resize columns
  periodsDataset.autoResizeColumns(1, periodsHeaders.length);
  
  // Success message with preservation info
  ui.alert(
    'Mapping Complete!',
    `✅ SUCCESSFULLY COMPLETED!\n\n` +
    `• Source: "${externalSheetName}"\n` +
    `• Destination: "${periodsDatasetName}"\n` +
    `• Periods mapped: ${totalMapped}\n` +
    `• Components for manual entry: ${totalComponents}\n` +
    `• Manual entries preserved: ${manualEntriesPreserved}\n` +
    `• New slots marked: ${newManualEntrySlots}\n` +
    `• Total rows processed: ${processedRows}\n` +
    `• External indicators used: ${periodsLookup.size}\n\n` +
    `🎉 Your periods dataset is ready!\n` +
    `${manualEntriesPreserved > 0 ? `🛡️ ${manualEntriesPreserved} manual entries were preserved.` : ''}`,
    ui.ButtonSet.OK
  );
}

// =============================================
// SMART REFRESH SYSTEM
// =============================================

/**
 * Update sheet configuration without rebuilding datasets
 */
function updateSheetConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  console.log('=== UPDATE SHEET CONFIGURATION ===');
  
  // Get available sheets
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  // Show current configuration
  const currentConfig = {
    combined: scriptProperties.getProperty('PERIODS_COMBINED_DATASET_NAME'),
    periods: scriptProperties.getProperty('PERIODS_DATASET_NAME'),
    docId: scriptProperties.getProperty('PERIODS_EXTERNAL_DOC_ID'),
    sheetName: scriptProperties.getProperty('PERIODS_EXTERNAL_SHEET_NAME'),
    components: scriptProperties.getProperty('CONCAT_COMPONENTS_SHEET'),
    indicators: scriptProperties.getProperty('CONCAT_INDICATORS_SHEET'),
    destination: scriptProperties.getProperty('CONCAT_DESTINATION_SHEET'),
    componentsData: scriptProperties.getProperty('COMPONENTS_DATA_SHEET_NAME')
  };
  
  let configDisplay = 'CURRENT CONFIGURATION:\n\n';
  configDisplay += `📊 Dataset Configuration:\n`;
  configDisplay += `• Components sheet: "${currentConfig.components || 'Not set'}"\n`;
  configDisplay += `• Indicators sheet: "${currentConfig.indicators || 'Not set'}"\n`;
  configDisplay += `• Combined dataset: "${currentConfig.combined || 'Not set'}"\n`;
  configDisplay += `• Calculation destination: "${currentConfig.destination || 'Not set'}"\n\n`;
  configDisplay += `📅 Periods Configuration:\n`;
  configDisplay += `• Periods dataset: "${currentConfig.periods || 'Not set'}"\n`;
  configDisplay += `• External document: ${currentConfig.docId || 'Not set'}\n`;
  configDisplay += `• External sheet: "${currentConfig.sheetName || 'Not set'}"\n\n`;
  configDisplay += `🧮 Components Data Sheet:\n`;
  configDisplay += `• Components data: "${currentConfig.componentsData || 'Not set'}"\n\n`;
  configDisplay += `Available sheets: ${sheetsText}\n\n`;
  configDisplay += `Which configuration would you like to update?`;
  
  const choice = ui.alert(
    'Update Sheet Configuration',
    configDisplay,
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (choice === ui.Button.CANCEL) return;
  
  if (choice === ui.Button.YES) {
    // Update dataset configuration
    updateDatasetConfiguration(ui, availableSheets, currentConfig, scriptProperties);
  } else {
    // Update periods configuration  
    updatePeriodsConfiguration(ui, availableSheets, currentConfig, scriptProperties);
  }
}

/**
 * Update dataset configuration (components, indicators, destination)
 */
function updateDatasetConfiguration(ui, availableSheets, currentConfig, scriptProperties) {
  const sheetsText = availableSheets.join(', ');
  
  // Update components sheet
  const componentsResult = ui.prompt(
    'Update Components Sheet',
    `Current: "${currentConfig.components || 'Not set'}"\n\n` +
    `Available: ${sheetsText}\n\n` +
    `Enter new components sheet name (or press Cancel to skip):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (componentsResult.getSelectedButton() === ui.Button.OK) {
    const newComponentsName = componentsResult.getResponseText().trim();
    if (newComponentsName && availableSheets.includes(newComponentsName)) {
      scriptProperties.setProperty('CONCAT_COMPONENTS_SHEET', newComponentsName);
      scriptProperties.setProperty('COMPONENTS_DATA_SHEET_NAME', newComponentsName);
      console.log(`Updated components sheet: ${newComponentsName}`);
    } else {
      ui.alert('Error', `Sheet "${newComponentsName}" not found.`, ui.ButtonSet.OK);
      return;
    }
  }
  
  // Update indicators sheet
  const indicatorsResult = ui.prompt(
    'Update Indicators Sheet',
    `Current: "${currentConfig.indicators || 'Not set'}"\n\n` +
    `Available: ${sheetsText}\n\n` +
    `Enter new indicators sheet name (or press Cancel to skip):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (indicatorsResult.getSelectedButton() === ui.Button.OK) {
    const newIndicatorsName = indicatorsResult.getResponseText().trim();
    if (newIndicatorsName && availableSheets.includes(newIndicatorsName)) {
      scriptProperties.setProperty('CONCAT_INDICATORS_SHEET', newIndicatorsName);
      console.log(`Updated indicators sheet: ${newIndicatorsName}`);
    } else {
      ui.alert('Error', `Sheet "${newIndicatorsName}" not found.`, ui.ButtonSet.OK);
      return;
    }
  }
  
  // Update destination sheet
  const destResult = ui.prompt(
    'Update Destination Sheet',
    `Current: "${currentConfig.destination || 'Not set'}"\n\n` +
    `Available: ${sheetsText}\n\n` +
    `Enter new destination sheet name (or press Cancel to skip):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (destResult.getSelectedButton() === ui.Button.OK) {
    const newDestName = destResult.getResponseText().trim();
    if (newDestName && availableSheets.includes(newDestName)) {
      scriptProperties.setProperty('CONCAT_DESTINATION_SHEET', newDestName);
      scriptProperties.setProperty('PERIODS_COMBINED_DATASET_NAME', newDestName);
      console.log(`Updated destination sheet: ${newDestName}`);
    } else {
      ui.alert('Error', `Sheet "${newDestName}" not found.`, ui.ButtonSet.OK);
      return;
    }
  }
  
  ui.alert(
    'Dataset Configuration Updated!',
    'Your dataset configuration has been updated.\n\n' +
    'You can now use "Refresh Combined Dataset" and other functions\n' +
    'without losing your existing data!',
    ui.ButtonSet.OK
  );
}

/**
 * Update periods configuration (periods dataset, external document)
 */
function updatePeriodsConfiguration(ui, availableSheets, currentConfig, scriptProperties) {
  const sheetsText = availableSheets.join(', ');
  
  // Update periods dataset
  const periodsResult = ui.prompt(
    'Update Periods Dataset Sheet',
    `Current: "${currentConfig.periods || 'Not set'}"\n\n` +
    `Available: ${sheetsText}\n\n` +
    `Enter new periods dataset name (or press Cancel to skip):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodsResult.getSelectedButton() === ui.Button.OK) {
    const newPeriodsName = periodsResult.getResponseText().trim();
    if (newPeriodsName && availableSheets.includes(newPeriodsName)) {
      scriptProperties.setProperty('PERIODS_DATASET_NAME', newPeriodsName);
      console.log(`Updated periods dataset: ${newPeriodsName}`);
    } else {
      ui.alert('Error', `Sheet "${newPeriodsName}" not found.`, ui.ButtonSet.OK);
      return;
    }
  }
  
  // Update external document (optional)
  const extDocResult = ui.prompt(
    'Update External Document (Optional)',
    `Current: ${currentConfig.docId || 'Not set'}\n\n` +
    `Enter new external document ID (or press Cancel to skip):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (extDocResult.getSelectedButton() === ui.Button.OK) {
    const newDocId = extDocResult.getResponseText().trim();
    if (newDocId) {
      scriptProperties.setProperty('PERIODS_EXTERNAL_DOC_ID', newDocId);
      console.log(`Updated external document: ${newDocId}`);
    }
  }
  
  // Update external sheet (optional)
  const extSheetResult = ui.prompt(
    'Update External Sheet (Optional)',
    `Current: "${currentConfig.sheetName || 'Not set'}"\n\n` +
    `Enter new external sheet name (or press Cancel to skip):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (extSheetResult.getSelectedButton() === ui.Button.OK) {
    const newSheetName = extSheetResult.getResponseText().trim();
    if (newSheetName) {
      scriptProperties.setProperty('PERIODS_EXTERNAL_SHEET_NAME', newSheetName);
      console.log(`Updated external sheet: ${newSheetName}`);
    }
  }
  
  ui.alert(
    'Periods Configuration Updated!',
    'Your periods configuration has been updated.\n\n' +
    'You can now use "Smart Refresh All Periods Data" and other functions\n' +
    'without issues!',
    ui.ButtonSet.OK
  );
}
