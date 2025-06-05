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
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Get available sheets
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  console.log('=== CONFIGURE SMART REFRESH ===');
  
  // Step 1: Combined dataset name
  const combinedResult = ui.prompt(
    'Configure Smart Refresh - Step 1/4',
    `COMBINED DATASET SHEET:\nThis is your merged components + indicators sheet.\n\n` +
    `Available sheets: ${sheetsText}\n\n` +
    `Enter your combined dataset sheet name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (combinedResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const combinedDatasetName = combinedResult.getResponseText().trim();
  
  // Validate combined dataset exists
  if (!ss.getSheetByName(combinedDatasetName)) {
    ui.alert('Error', `Sheet "${combinedDatasetName}" not found. Please check the name and try again.`, ui.ButtonSet.OK);
    return;
  }
  
  // Step 2: Periods dataset name
  const periodsResult = ui.prompt(
    'Configure Smart Refresh - Step 2/4',
    `PERIODS DATASET SHEET:\nThis will contain your evaluation periods data.\n\n` +
    `Enter your periods dataset sheet name (will be created/recreated as needed):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodsResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const periodsDatasetName = periodsResult.getResponseText().trim();
  
  // Step 3: External document ID
  const docResult = ui.prompt(
    'Configure Smart Refresh - Step 3/4',
    `EXTERNAL DOCUMENT ID:\nThis is the Google Sheets document containing your evaluation dates.\n\n` +
    `Enter the external Google Sheets document ID:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (docResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const externalDocId = docResult.getResponseText().trim();
  
  // Step 4: Test external document access and get sheet name
  let sourceSpreadsheet, availableExternalSheets;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(externalDocId);
    availableExternalSheets = sourceSpreadsheet.getSheets().map(sheet => sheet.getName());
  } catch (e) {
    ui.alert('Error', `Cannot access external document with ID: ${externalDocId}\n\nError: ${e.message}\n\nPlease check the document ID and make sure it's shared with you.`, ui.ButtonSet.OK);
    return;
  }
  
  const externalSheetsText = availableExternalSheets.join(', ');
  
  const sheetResult = ui.prompt(
    'Configure Smart Refresh - Step 4/4',
    `EXTERNAL SHEET NAME:\nChoose the sheet containing your evaluation periods data.\n\n` +
    `Available sheets in external document: ${externalSheetsText}\n\n` +
    `Enter the sheet name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const externalSheetName = sheetResult.getResponseText().trim();
  
  // Validate external sheet exists
  const externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  if (!externalSheet) {
    ui.alert('Error', `Sheet "${externalSheetName}" not found in external document.`, ui.ButtonSet.OK);
    return;
  }
  
  // Test external sheet has required columns
  try {
    const testData = externalSheet.getDataRange().getValues();
    const testHeaders = testData[0];
    const idColumn = findIdColumn(testHeaders);
    const periodColumns = findPeriodColumns(testHeaders);
    
    if (idColumn === -1) {
      ui.alert('Warning', `Could not find ID Indicador column in "${externalSheetName}". Make sure it has a column with "ID" and "Indicador" in the name.`, ui.ButtonSet.OK);
    }
    
    if (periodColumns.length === 0) {
      ui.alert('Warning', `Could not find evaluation period columns in "${externalSheetName}". Make sure it has columns like "Periodo Base EVA Mayo", etc.`, ui.ButtonSet.OK);
    }
    
    console.log(`External validation: ID column ${idColumn >= 0 ? 'found' : 'NOT found'}, ${periodColumns.length} period columns found`);
    
  } catch (e) {
    ui.alert('Warning', `Could not validate external sheet structure: ${e.message}`, ui.ButtonSet.OK);
  }
  
  // Final confirmation
  const confirmResult = ui.alert(
    'Confirm Smart Refresh Configuration',
    `Please confirm your configuration:\n\n` +
    `✅ Combined dataset: "${combinedDatasetName}"\n` +
    `✅ Periods dataset: "${periodsDatasetName}"\n` +
    `✅ External document: ${externalDocId}\n` +
    `✅ External sheet: "${externalSheetName}"\n\n` +
    `Save this configuration?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    return;
  }
  
  // Save configuration
  scriptProperties.setProperties({
    'PERIODS_COMBINED_DATASET_NAME': combinedDatasetName,
    'PERIODS_DATASET_NAME': periodsDatasetName,
    'PERIODS_EXTERNAL_DOC_ID': externalDocId,
    'PERIODS_EXTERNAL_SHEET_NAME': externalSheetName
  });
  
  console.log('Configuration saved:', {
    combined: combinedDatasetName,
    periods: periodsDatasetName,
    docId: externalDocId,
    sheet: externalSheetName
  });
  
  ui.alert(
    '🎉 Smart Refresh Configured!',
    `Configuration saved successfully!\n\n` +
    `Your settings:\n` +
    `• Combined dataset: "${combinedDatasetName}"\n` +
    `• Periods dataset: "${periodsDatasetName}"\n` +
    `• External document ID: ${externalDocId}\n` +
    `• External sheet: "${externalSheetName}"\n\n` +
    `✅ You can now use "Refresh All Periods Data" for one-click updates!\n` +
    `✅ Use "Reconfigure Smart Refresh" anytime to change these settings.`,
    ui.ButtonSet.OK
  );
}

/**
 * Smart refresh all periods data with manual entry preservation
 */
function refreshAllPeriodsData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  console.log('=== ENHANCED SMART REFRESH ALL PERIODS DATA ===');
  
  // Check if configuration exists
  const config = {
    combined: scriptProperties.getProperty('PERIODS_COMBINED_DATASET_NAME'),
    periods: scriptProperties.getProperty('PERIODS_DATASET_NAME'),
    docId: scriptProperties.getProperty('PERIODS_EXTERNAL_DOC_ID'),
    sheetName: scriptProperties.getProperty('PERIODS_EXTERNAL_SHEET_NAME')
  };
  
  // Validate configuration exists
  if (!config.combined || !config.periods || !config.docId || !config.sheetName) {
    ui.alert(
      'Configuration Required',
      `Smart Refresh is not configured yet.\n\n` +
      `Please run "Configure Smart Refresh" first to set up your sheet names and external document.`,
      ui.ButtonSet.OK
    );
    return;
  }
  
  console.log('Using configuration:', config);
  
  // Confirm with user
  const confirmResult = ui.alert(
    'Refresh All Periods Data',
    `This will refresh your complete periods workflow using:\n\n` +
    `📊 Combined dataset: "${config.combined}"\n` +
    `📅 Periods dataset: "${config.periods}"\n` +
    `🔗 External source: "${config.sheetName}"\n\n` +
    `Steps:\n` +
    `1. ✅ Refresh Combined Dataset\n` +
    `2. ✅ Recreate Periods Dataset\n` +
    `3. ✅ Remap Periods (preserving manual entries)\n\n` +
    `⚠️ Manual entries will be preserved!\n\n` +
    `This will take 3-5 minutes. Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    return;
  }
  
  const startTime = new Date();
  let currentStep = 1;
  
  try {
    // Step 1: Refresh Combined Dataset
    console.log(`Step ${currentStep}: Refreshing Combined Dataset "${config.combined}"...`);
    
    if (typeof refreshConcatenatedDataset === 'function') {
      refreshConcatenatedDataset();
      console.log('✅ Combined dataset refreshed');
    } else {
      throw new Error('refreshConcatenatedDataset function not found. Make sure DataCombiner.js is loaded.');
    }
    
    currentStep++;
    
    // Step 2: Recreate Evaluation Periods Dataset with Preservation
    console.log(`Step ${currentStep}: Recreating Periods Dataset "${config.periods}" (preserving manual entries)...`);
    
    const combinedSheet = ss.getSheetByName(config.combined);
    if (!combinedSheet) {
      throw new Error(`Combined dataset sheet "${config.combined}" not found after refresh.`);
    }
    
    // Extract and preserve manual entries
    const existingPeriodsSheet = ss.getSheetByName(config.periods);
    let manualEntries = {};
    
    if (existingPeriodsSheet) {
      console.log('Extracting manual entries before recreation...');
      manualEntries = extractManualEntries(existingPeriodsSheet);
      console.log(`Found ${Object.keys(manualEntries).length} manual entries to preserve`);
    }
    
    recreatePeriodsDatasetAutomated(combinedSheet, config.periods, ss);
    
    // Restore manual entries
    if (Object.keys(manualEntries).length > 0) {
      console.log('Restoring manual entries...');
      const restoredCount = restoreManualEntries(ss.getSheetByName(config.periods), manualEntries);
      console.log(`Restored ${restoredCount} manual entries`);
    }
    
    console.log('✅ Periods dataset recreated with manual entries preserved');
    
    currentStep++;
    
    // Step 3: Remap Periods from External Document
    console.log(`Step ${currentStep}: Remapping periods from external "${config.sheetName}"...`);
    
    const periodsDataset = ss.getSheetByName(config.periods);
    if (!periodsDataset) {
      throw new Error(`Periods dataset "${config.periods}" not found after recreation.`);
    }
    
    let sourceSpreadsheet, externalSheet;
    try {
      sourceSpreadsheet = SpreadsheetApp.openById(config.docId);
      externalSheet = sourceSpreadsheet.getSheetByName(config.sheetName);
    } catch (e) {
      throw new Error(`Cannot access external document: ${e.message}`);
    }
    
    if (!externalSheet) {
      throw new Error(`External sheet "${config.sheetName}" not found in external document.`);
    }
    
    // Perform mapping with manual entry preservation
    const mappingResult = performMappingWithBatchesSilent(externalSheet, periodsDataset, config.sheetName, config.periods);
    console.log('✅ Periods remapped successfully with preservation');
    
    const endTime = new Date();
    const totalTime = Math.round((endTime - startTime) / 1000);
    
    console.log('=== ENHANCED SMART REFRESH COMPLETE ===');
    console.log(`Results: ${mappingResult.totalMapped} mapped, ${mappingResult.manualEntriesPreserved} preserved, ${totalTime}s total`);
    
    // Enhanced success message with preservation info
    ui.alert(
      '🎉 Smart Refresh Complete!',
      `All periods data refreshed successfully!\n\n` +
      `✅ Combined dataset refreshed\n` +
      `✅ Periods dataset recreated\n` +
      `✅ ${mappingResult.totalMapped} periods remapped\n` +
      `✅ ${mappingResult.manualEntriesPreserved} manual entries preserved\n` +
      `✅ ${mappingResult.newManualEntrySlots} new slots marked for manual entry\n\n` +
      `📊 Configuration used:\n` +
      `• Combined dataset: "${config.combined}"\n` +
      `• Periods dataset: "${config.periods}"\n` +
      `• External source: "${config.sheetName}"\n\n` +
      `⏱️ Total time: ${totalTime} seconds\n\n` +
      `${mappingResult.manualEntriesPreserved > 0 ? 
        `🛡️ Your ${mappingResult.manualEntriesPreserved} manual entries were preserved and NOT updated.\n` +
        `To update them, change those cells manually.\n\n` : ''}` +
      `🎉 Your periods data is fully updated!`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Enhanced smart refresh error:', error);
    ui.alert(
      'Smart Refresh Error',
      `Error occurred during step ${currentStep}:\n\n${error.message}\n\n` +
      `Configuration being used:\n` +
      `• Combined: "${config.combined}"\n` +
      `• Periods: "${config.periods}"\n` +
      `• External: "${config.sheetName}"\n\n` +
      `You can:\n` +
      `1. Try running individual steps manually\n` +
      `2. Run "Reconfigure Smart Refresh" to update settings\n` +
      `3. Check that all sheets exist and are accessible`,
      ui.ButtonSet.OK
    );
  }
}

// =============================================
// SMART CALCULATION REFRESH SYSTEM
// =============================================

/**
 * NOTE: Smart formula refresh functionality has been moved to SmartFormulaRefresh.js
 * for better code organization and maintainability.
 * 
 * Use: smartRefreshCalculationFormulas() from SmartFormulaRefresh.js
 * 
 * This function intelligently updates calculation formulas when periods change
 * (SINGLE ↔ SUM) while preserving manual formula entries.
 */

// =============================================
// CALCULATION GENERATION
// =============================================

/**
 * Generate calculation formulas using the evaluation periods dataset
 */
function generateCalculationsFromPeriodsDataset() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask for the evaluation periods dataset (SOURCE of period data)
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  const periodsDatasetResult = ui.prompt(
    'Select Evaluation Periods Dataset (SOURCE)',
    `Available sheets: ${sheetsText}\n\nEnter the name of your evaluation periods dataset (where the dates are stored):`,
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
  
  // Ask for the calculations destination sheet (WHERE to put the formulas)
  const calculationsResult = ui.prompt(
    'Select Calculations Destination Sheet',
    `Available sheets: ${sheetsText}\n\nEnter the name of the sheet where you want to ADD the calculation formulas (e.g., "calculos"):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (calculationsResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const calculationsSheetName = calculationsResult.getResponseText().trim();
  const calculationsSheet = ss.getSheetByName(calculationsSheetName);
  
  if (!calculationsSheet) {
    ui.alert('Error', `Calculations sheet "${calculationsSheetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Ask which evaluation period to generate formulas for
  const periodResult = ui.prompt(
    'Select Evaluation Period',
    `Available periods:\n${EVALUATION_PERIOD_COLUMNS.join('\n')}\n\nEnter the period name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const selectedPeriod = periodResult.getResponseText().trim();
  
  // Confirm the setup
  const confirmResult = ui.alert(
    'Confirm Setup',
    `Period data source: "${periodsDatasetName}"\n` +
    `Calculations destination: "${calculationsSheetName}"\n` +
    `Selected period: "${selectedPeriod}"\n\n` +
    `This will add a new calculation column to your "${calculationsSheetName}" sheet.\n\n` +
    `Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    return;
  }
  
  // Generate formulas
  generateFormulasFromPeriodsDataToDestination(periodsDataset, calculationsSheet, selectedPeriod, ss);
}

// =============================================
// HELPER FUNCTIONS FOR MANUAL ENTRY PRESERVATION
// =============================================

/**
 * Extract manual entries from existing periods dataset before recreation
 */
function extractManualEntries(periodsSheet) {
  const data = periodsSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const manualEntries = {};
  
  // Find relevant column indices
  const municipioIndex = headers.indexOf('municipio');
  const idIndicadorIndex = headers.indexOf('id_indicador');
  const tipoIndex = headers.indexOf('tipo de dato');
  const codigoIndex = headers.indexOf('codigo');
  
  rows.forEach((row, rowIndex) => {
    const municipio = row[municipioIndex];
    const idIndicador = row[idIndicadorIndex];
    const tipo = row[tipoIndex];
    const codigo = row[codigoIndex];
    
    // Only check components with id_indicador = 0 (manual entry rows)
    if (idIndicador === 0 || idIndicador === '0') {
      const rowKey = `${municipio}_${idIndicador}_${tipo}_${codigo}`;
      
      // Check each period column for manual entries
      EVALUATION_PERIOD_COLUMNS.forEach(periodCol => {
        const colIndex = headers.indexOf(periodCol);
        if (colIndex !== -1) {
          const cellValue = row[colIndex];
          
          // If there's a value that's not empty and not the default note
          if (cellValue && 
              cellValue !== '' && 
              cellValue !== 'Manual entry needed for component without indicator') {
            
            if (!manualEntries[rowKey]) {
              manualEntries[rowKey] = {};
            }
            manualEntries[rowKey][periodCol] = cellValue;
            console.log(`Extracted manual entry: ${rowKey} -> ${periodCol} = "${cellValue}"`);
          }
        }
      });
    }
  });
  
  return manualEntries;
}

/**
 * Restore manual entries to recreated periods dataset
 */
function restoreManualEntries(periodsSheet, manualEntries) {
  if (Object.keys(manualEntries).length === 0) {
    return 0;
  }
  
  const data = periodsSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  // Find relevant column indices
  const municipioIndex = headers.indexOf('municipio');
  const idIndicadorIndex = headers.indexOf('id_indicador');
  const tipoIndex = headers.indexOf('tipo de dato');
  const codigoIndex = headers.indexOf('codigo');
  
  let restoredCount = 0;
  
  rows.forEach((row, rowIndex) => {
    const actualRowNum = rowIndex + 2; // +2 for 1-based indexing and header
    const municipio = row[municipioIndex];
    const idIndicador = row[idIndicadorIndex];
    const tipo = row[tipoIndex];
    const codigo = row[codigoIndex];
    
    // Only restore for components with id_indicador = 0
    if (idIndicador === 0 || idIndicador === '0') {
      const rowKey = `${municipio}_${idIndicador}_${tipo}_${codigo}`;
      
      if (manualEntries[rowKey]) {
        // Restore each manual entry for this row
        Object.entries(manualEntries[rowKey]).forEach(([periodCol, value]) => {
          const colIndex = headers.indexOf(periodCol);
          if (colIndex !== -1) {
            periodsSheet.getRange(actualRowNum, colIndex + 1).setValue(value);
            restoredCount++;
            console.log(`Restored: ${rowKey} -> ${periodCol} = "${value}"`);
          }
        });
      }
    }
  });
  
  return restoredCount;
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

/**
 * Silent mapping function for smart refresh (returns results object)
 */
function performMappingWithBatchesSilent(externalSheet, periodsDataset, externalSheetName, periodsDatasetName) {
  console.log('Building external lookup for silent mapping...');
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  const sourceIdIndex = findIdColumn(sourceHeaders);
  if (sourceIdIndex === -1) {
    throw new Error(`Could not find ID Indicador column in external sheet "${externalSheetName}".`);
  }
  
  // Find period columns
  const periodColumns = findPeriodColumns(sourceHeaders);
  if (periodColumns.length === 0) {
    throw new Error(`No period columns found in external sheet "${externalSheetName}".`);
  }
  
  console.log(`Found ${periodColumns.length} period columns for silent mapping`);
  
  // Build periods lookup
  const periodsLookup = new Map();
  sourceRows.forEach((row, index) => {
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
  
  console.log(`External lookup built with ${periodsLookup.size} indicators for silent mapping`);
  
  // Process periods dataset (same logic as batch processing but silent)
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
  
  console.log(`Processing ${totalRows} rows silently in batches of ${batchSize}...`);
  
  // Process in batches (same logic as before but without UI updates)
  for (let startRow = 2; startRow <= totalRows + 1; startRow += batchSize) {
    const endRow = Math.min(startRow + batchSize - 1, totalRows + 1);
    const batchRows = endRow - startRow + 1;
    
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
          
          if (currentValue && currentValue !== '' && currentValue !== 'Manual entry needed for component without indicator') {
            // PRESERVE EXISTING MANUAL ENTRY
            manualEntriesPreserved++;
          } else {
            // NEW EMPTY SLOT - Mark for manual entry
            newManualEntrySlots++;
            currentCell.setBackground('#ffcccb');
            currentCell.setNote('Manual entry needed for component without indicator');
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
    
    // Small pause every 500 rows
    if ((startRow - 2) % 500 === 0 && startRow > 2) {
      Utilities.sleep(50);
    }
  }
  
  console.log(`Silent mapping complete: ${totalMapped} mapped, ${totalComponents} total components`);
  console.log(`Manual entries: ${manualEntriesPreserved} preserved, ${newManualEntrySlots} new slots marked`);
  
  // Auto-resize columns
  periodsDataset.autoResizeColumns(1, periodsHeaders.length);
  
  return { 
    totalMapped, 
    totalComponents: totalComponents,
    manualEntriesPreserved,
    newManualEntrySlots,
    processedRows 
  };
}

/**
 * Generate formulas from periods data to destination sheet (optimized version)
 */
function generateFormulasFromPeriodsDataToDestination(periodsDataset, destinationSheet, selectedPeriod, ss) {
  const ui = SpreadsheetApp.getUi();
  
  console.log('Starting optimized formula generation...');
  const startTime = new Date();
  
  // Ask user for components dataset sheet name
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  const componentsSheetResult = ui.prompt(
    'Select Components Dataset Sheet',
    `Available sheets: ${sheetsText}\n\nEnter the name of your COMPONENTS dataset sheet (for formula references):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (componentsSheetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const componentsSheetName = componentsSheetResult.getResponseText().trim();
  const componentsSheet = ss.getSheetByName(componentsSheetName);
  
  if (!componentsSheet) {
    ui.alert('Error', `Components sheet "${componentsSheetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Get data
  const periodsData = periodsDataset.getDataRange().getValues();
  const periodsHeaders = periodsData[0];
  const periodsRows = periodsData.slice(1);
  
  const destData = destinationSheet.getDataRange().getValues();
  const destHeaders = destData[0];
  const destRows = destData.slice(1);
  
  // Validate columns using FLEXIBLE column detection
  const periodsIndices = getFlexibleColumnIndices(periodsHeaders, ['municipio', 'id_indicador', 'tipo de dato', 'codigo']);
  const destIndices = getFlexibleColumnIndices(destHeaders, ['municipio', 'id_indicador', 'tipo de dato', 'codigo']);
  
  // Validate selected period exists
  const selectedPeriodIndex = periodsHeaders.indexOf(selectedPeriod);
  
  if (!periodsIndices.valid || !destIndices.valid || selectedPeriodIndex === -1) {
    let errorMsg = 'Required columns not found.';
    if (selectedPeriodIndex === -1) {
      errorMsg += `\n\nPeriod "${selectedPeriod}" not found in periods dataset.\nAvailable periods: ${periodsHeaders.filter(h => EVALUATION_PERIOD_COLUMNS.includes(h)).join(', ')}`;
    }
    ui.alert('Error', errorMsg, ui.ButtonSet.OK);
    return;
  }
  
  // Create periods lookup
  const periodsLookup = new Map();
  periodsRows.forEach(row => {
    const key = `${row[periodsIndices.municipio]}_${row[periodsIndices.id_indicador]}_${row[periodsIndices.tipo_de_dato]}_${row[periodsIndices.codigo]}`;
    // FIXED: Use the actual column index for the selected period
    const selectedPeriodIndex = periodsHeaders.indexOf(selectedPeriod);
    if (selectedPeriodIndex !== -1) {
      periodsLookup.set(key, row[selectedPeriodIndex]);
    }
  });
  
  // Add calculation column
  const calculationColumnIndex = destHeaders.length + 1;
  const calculationColumnName = `Calculo_${selectedPeriod}`;
  
  destinationSheet.getRange(1, calculationColumnIndex).setValue(calculationColumnName);
  destinationSheet.getRange(1, calculationColumnIndex).setFontWeight('bold').setBackground('#e8f4fd');
  
  // Generate formulas in batch - NOW WITH PROPER SHEET REFERENCE
  const formulas = generateFormulasBatch(destRows, destIndices, periodsLookup, calculationColumnIndex, componentsSheetName, destHeaders);
  
  // Write all formulas at once
  if (formulas.length > 0) {
    const range = destinationSheet.getRange(2, calculationColumnIndex, formulas.length, 1);
    range.setFormulas(formulas);
  }
  
  destinationSheet.autoResizeColumn(calculationColumnIndex);
  
  const endTime = new Date();
  const executionTime = (endTime - startTime) / 1000;
  
  ui.alert('Success!', `Formulas generated in ${executionTime.toFixed(1)} seconds!`, ui.ButtonSet.OK);
}

/**
 * Generate formulas in batch for optimization - FIXED VERSION WITH PROPER SHEET REFERENCES
 * Now generates proper sheet row references instead of timestamps
 */
function generateFormulasBatch(destRows, destIndices, periodsLookup, calculationColumnIndex, componentsSheetName, destHeaders) {
  const formulas = [];
  
  // Get the components sheet to build row lookup
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const componentsSheet = ss.getSheetByName(componentsSheetName);
  const componentsData = componentsSheet.getDataRange().getValues();
  const componentsHeaders = componentsData[0];
  const componentsRows = componentsData.slice(1);
  
  // Build lookup map: "municipio_codigo" -> actual row number in components sheet
  const componentRowLookup = new Map();
  const municipioIndex = componentsHeaders.indexOf('municipio');
  const codigoIndex = componentsHeaders.findIndex(h => h.includes('codigo') || h.includes('código'));
  
  if (municipioIndex !== -1 && codigoIndex !== -1) {
    componentsRows.forEach((row, index) => {
      const municipio = row[municipioIndex];
      const codigo = row[codigoIndex];
      if (municipio && codigo) {
        const key = `${municipio}_${codigo}`;
        const actualRowNum = index + 2; // +2 for 1-based indexing and header
        componentRowLookup.set(key, actualRowNum);
      }
    });
  }
  
  console.log(`Built component lookup with ${componentRowLookup.size} entries`);
  
  destRows.forEach((row, rowIndex) => {
    const actualRowNum = rowIndex + 2;
    const municipio = row[destIndices.municipio];
    const id_indicador = row[destIndices.id_indicador];
    const tipo = row[destIndices.tipo_de_dato];
    const codigo = row[destIndices.codigo];
    
    if (!codigo || codigo === '') {
      formulas.push(['']);
      return;
    }
    
    const lookupKey = `${municipio}_${id_indicador}_${tipo}_${codigo}`;
    const period = periodsLookup.get(lookupKey);
    
    let formula = '';
    
    if (tipo === 'Componente') {
      // FIXED: Better handling of manual periods and validation
      if (!period || period === '' || period === 'NA') {
        // Period is truly missing
        formula = `=IFERROR("PERIOD_MISSING_FOR_${codigo}", "")`;
      } else if (period.includes('ERROR:')) {
        // Period has normalization error
        formula = `=IFERROR("PERIOD_ERROR_FOR_${codigo}", "")`;
      } else if (period.includes('Manual entry needed')) {
        // Period needs manual entry (should not reach formula generation)
        formula = `=IFERROR("MANUAL_ENTRY_NEEDED_${codigo}", "")`;
      } else {
        // FIXED: Valid period found - now finds correct row in components sheet
        try {
          const formulaDateRange = convertNormalizedPeriodToFormula(period);
          const municipioCell = `${getColumnLetter(destIndices.municipio + 1)}${actualRowNum}`;
          
          // FIXED: Find actual row in components sheet for this municipality + codigo
          const componentLookupKey = `${municipio}_${codigo}`;
          const componentRowNum = componentRowLookup.get(componentLookupKey);
          
          if (componentRowNum) {
            const componentRowRef = `${componentsSheetName}!${componentRowNum}:${componentRowNum}`;
            
            if (formulaDateRange.includes('-')) {
              formula = `=CALCULATE_INDICATOR("SUM(${codigo}:${formulaDateRange})", ${municipioCell}, ${componentRowRef})`;
            } else {
              formula = `=CALCULATE_INDICATOR("SINGLE(${codigo}:${formulaDateRange})", ${municipioCell}, ${componentRowRef})`;
            }
          } else {
            formula = `=IFERROR("COMPONENT_NOT_FOUND_${codigo}_${municipio}", "")`;
            console.log(`Component not found in ${componentsSheetName}: ${municipio}_${codigo}`);
          }
        } catch (error) {
          console.log(`Error processing period "${period}" for ${codigo}: ${error.message}`);
          formula = `=IFERROR("PERIOD_FORMAT_ERROR_${codigo}", "")`;
        }
      }
    } else if (tipo === 'Indicador') {
      formula = generateIndicatorCellFormulaOptimized(codigo, actualRowNum, getColumnLetter(calculationColumnIndex), destRows, destIndices);
    }
    
    formulas.push([formula]);
  });
  
  return formulas;
}



/**
 * Generate indicator formula for destination sheet (optimized)
 */
function generateIndicatorCellFormulaOptimized(formulaText, currentRow, calculationColumn, rows, columnIndices) {
  console.log(`Processing indicator formula: "${formulaText}" at row ${currentRow}`);
  
  const componentCodes = extractComponentCodes(formulaText);
  console.log(`Extracted component codes: ${JSON.stringify(componentCodes)}`);
  
  if (componentCodes.length === 0) {
    console.log(`No component codes found in formula: "${formulaText}"`);
    return '';
  }
  
  const currentRowData = rows[currentRow - 2];
  const currentMunicipio = currentRowData[columnIndices.municipio];
  
  const componentRefs = {};
  const componentLookup = new Map();
  
  rows.forEach((row, index) => {
    if (row[columnIndices.tipo_de_dato] === 'Componente' && 
        row[columnIndices.municipio] === currentMunicipio) {
      const codigo = row[columnIndices.codigo];
      if (codigo) {
        const key = codigo.toLowerCase();
        if (!componentLookup.has(key)) {
          componentLookup.set(key, {
            rowNum: index + 2,
            originalCode: codigo
          });
        }
      }
    }
  });
  
  console.log(`Component lookup for ${currentMunicipio}: ${componentLookup.size} components`);
  
  componentCodes.forEach(code => {
    const componentInfo = componentLookup.get(code.toLowerCase());
    if (componentInfo) {
      componentRefs[code] = `${calculationColumn}${componentInfo.rowNum}`;
      console.log(`Mapped ${code} -> ${componentRefs[code]}`);
    } else {
      console.log(`Component not found: ${code} for municipality ${currentMunicipio}`);
    }
  });
  
  // SPECIAL CASE: If formula is just a single component code, make it a simple reference
  if (componentCodes.length === 1 && formulaText.trim() === componentCodes[0]) {
    const singleCode = componentCodes[0];
    const cellRef = componentRefs[singleCode];
    if (cellRef) {
      console.log(`Simple reference case: ${formulaText} -> =${cellRef}`);
      return `=${cellRef}`;
    } else {
      console.log(`Simple reference failed - component not found: ${singleCode}`);
      return '';
    }
  }
  
  let finalFormula = formulaText;
  Object.entries(componentRefs).forEach(([code, cellRef]) => {
    const regex = new RegExp(`\\b${code}\\b`, 'gi');
    finalFormula = finalFormula.replace(regex, cellRef);
  });
  
  // ENHANCED FIX: Safe string check before using startsWith
  if (finalFormula && typeof finalFormula === 'string' && 
      !finalFormula.startsWith('=') && finalFormula.trim() !== '') {
    finalFormula = '=' + finalFormula;
  }
  
  console.log(`Final formula: ${finalFormula}`);
  return finalFormula;
}