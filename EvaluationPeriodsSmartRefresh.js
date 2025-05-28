/**
 * EvaluationPeriodsSmartRefresh.gs
 * Smart refresh functions for automated periods data updates
 * Separate file to keep EvaluationPeriodsMapping.gs manageable
 */

/**
 * ENHANCED SMART REFRESH - PRESERVES MANUAL ENTRIES
 * Replace the performMappingWithBatchesSilentRobust function in your EvaluationPeriodsSmartRefresh.gs file
 * This version preserves manually entered dates and only marks new empty cells
 */

/**
 * RECONFIGURE SMART REFRESH: Update any configuration at any time
 */
function reconfigureSmartRefresh() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Show current configuration if it exists
  const currentConfig = {
    combined: scriptProperties.getProperty('PERIODS_COMBINED_DATASET_NAME') || 'Not set',
    periods: scriptProperties.getProperty('PERIODS_DATASET_NAME') || 'Not set',
    docId: scriptProperties.getProperty('PERIODS_EXTERNAL_DOC_ID') || 'Not set',
    sheetName: scriptProperties.getProperty('PERIODS_EXTERNAL_SHEET_NAME') || 'Not set'
  };
  
  const showCurrentResult = ui.alert(
    'Current Smart Refresh Configuration',
    `Current settings:\n\n` +
    `• Combined dataset: "${currentConfig.combined}"\n` +
    `• Periods dataset: "${currentConfig.periods}"\n` +
    `• External document ID: ${currentConfig.docId}\n` +
    `• External sheet: "${currentConfig.sheetName}"\n\n` +
    `Would you like to update these settings?`,
    ui.ButtonSet.YES_NO
  );
  
  if (showCurrentResult !== ui.Button.YES) {
    return;
  }
  
  // Run the configuration process
  configureSmartRefresh();
}

/**
 * ENHANCED CONFIGURE SMART REFRESH: More user-friendly with validation
 */
function configureSmartRefresh() {
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
    const idColumn = findIdColumnRobust(testHeaders);
    const periodColumns = findPeriodColumnsRobust(testHeaders);
    
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
 * ENHANCED REFRESH FUNCTION: Better success message with preservation info
 * Replace the refreshAllPeriodsData function in your EvaluationPeriodsSmartRefresh.gs file
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
    // ==============================================
    // STEP 1: REFRESH COMBINED DATASET
    // ==============================================
    
    console.log(`Step ${currentStep}: Refreshing Combined Dataset "${config.combined}"...`);
    
    if (typeof refreshConcatenatedDataset === 'function') {
      refreshConcatenatedDataset();
      console.log('✅ Combined dataset refreshed');
    } else {
      throw new Error('refreshConcatenatedDataset function not found. Make sure DataCombiner.gs is loaded.');
    }
    
    currentStep++;
    
    // ==============================================
    // STEP 2: RECREATE EVALUATION PERIODS DATASET WITH PRESERVATION
    // ==============================================
    
    console.log(`Step ${currentStep}: Recreating Periods Dataset "${config.periods}" (preserving manual entries)...`);
    
    const combinedSheet = ss.getSheetByName(config.combined);
    if (!combinedSheet) {
      throw new Error(`Combined dataset sheet "${config.combined}" not found after refresh.`);
    }
    
    // ENHANCED: Preserve manual entries during recreation
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
    
    // ==============================================
    // STEP 3: REMAP PERIODS FROM EXTERNAL DOCUMENT
    // ==============================================
    
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
    const mappingResult = performMappingWithBatchesSilentRobust(externalSheet, periodsDataset, config.sheetName, config.periods);
    console.log('✅ Periods remapped successfully with preservation');
    
    // ==============================================
    // COMPLETION
    // ==============================================
    
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
// ROBUST HELPER FUNCTIONS (NO HARDCODING)
// =============================================

/**
 * Robust ID column finder (standalone, no dependencies)
 */
function findIdColumnRobust(headers) {
  const possibleNames = [
    'ID_Indicador', 'id_indicador', 'Id_Indicador', 'ID Indicador', 'id indicador',
    'IdIndicador', 'IDINDICADOR', 'Indicador_ID', 'IndicadorID'
  ];
  
  for (const name of possibleNames) {
    const index = headers.findIndex(h => h.toLowerCase().trim() === name.toLowerCase().trim());
    if (index !== -1) return index;
  }
  
  // Fallback: search for any column containing 'id' and 'indicador'
  return headers.findIndex(h => 
    h.toLowerCase().includes('id') && h.toLowerCase().includes('indicador')
  );
}

/**
 * Robust period columns finder (standalone, no dependencies)
 */
function findPeriodColumnsRobust(headers) {
  const expectedColumns = [
    { target: 'Periodo_Base_EVA_Mayo', patterns: ['Periodo_Base_EVA_Mayo', 'Periodo Base EVA Mayo'] },
    { target: 'Periodo_Base_EVA_Noviembre', patterns: ['Periodo_Base_EVA_Noviembre', 'Periodo Base EVA Noviembre'] },
    { target: 'Diagnostico_Mayo_25', patterns: ['Diagnostico_Mayo_25', 'Diagnóstico (Mayo 25)', 'Diagnostico Mayo 25'] },
    { target: 'Evaluacion_1_Noviembre_25', patterns: ['Evaluacion_1_Noviembre_25', 'Evaluación 1 (Noviembre 25)', 'Evaluacion 1 Noviembre 25'] },
    { target: 'Evaluacion_2_Mayo_26', patterns: ['Evaluacion_2_Mayo_26', 'Evaluación 2 (Mayo 26)', 'Evaluacion 2 Mayo 26'] },
    { target: 'Evaluacion_3_Noviembre_26', patterns: ['Evaluacion_3_Noviembre_26', 'Evaluación 3 (Noviembre 26)', 'Evaluacion 3 Noviembre 26'] },
    { target: 'Evaluacion_4_Mayo_27', patterns: ['Evaluacion_4_Mayo_27', 'Evaluación 4 (Mayo 27)', 'Evaluacion 4 Mayo 27'] },
    { target: 'Evaluacion_5_Octubre_27', patterns: ['Evaluacion_5_Octubre_27', 'Evaluación 5 (Octubre 27)', 'Evaluacion 5 Octubre 27'] }
  ];
  
  const mappings = [];
  expectedColumns.forEach(col => {
    for (const pattern of col.patterns) {
      const index = headers.findIndex(h => h.includes(pattern));
      if (index !== -1) {
        mappings.push({
          target: col.target,
          sourceIndex: index,
          sourceHeader: headers[index]
        });
        break;
      }
    }
  });
  
  return mappings;
}

/**
 * Enhanced silent mapping with manual entry preservation
 */
function performMappingWithBatchesSilentRobust(externalSheet, periodsDataset, externalSheetName, periodsDatasetName) {
  console.log('Building external lookup...');
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  const sourceIdIndex = findIdColumnRobust(sourceHeaders);
  if (sourceIdIndex === -1) {
    throw new Error(`Could not find ID Indicador column in external sheet "${externalSheetName}". Available columns: ${sourceHeaders.join(', ')}`);
  }
  
  console.log(`Found external ID column: "${sourceHeaders[sourceIdIndex]}" at index ${sourceIdIndex}`);
  
  // Find period columns
  const periodColumns = findPeriodColumnsRobust(sourceHeaders);
  if (periodColumns.length === 0) {
    throw new Error(`No period columns found in external sheet "${externalSheetName}". Available columns: ${sourceHeaders.join(', ')}`);
  }
  
  console.log(`Found ${periodColumns.length} period columns:`, periodColumns.map(c => c.sourceHeader));
  
  // Build periods lookup
  const periodsLookup = new Map();
  sourceRows.forEach((row, index) => {
    const id = String(row[sourceIdIndex]).trim();
    if (id && id !== '' && id !== '...') {
      const periods = {};
      periodColumns.forEach(col => {
        const value = row[col.sourceIndex];
        if (value && value !== '' && value !== '...') {
          periods[col.target] = normalizePeriodText(value.toString());
        }
      });
      periodsLookup.set(id, periods);
      
      // Log first few for debugging
      if (index < 3) {
        console.log(`External ID ${id} -> ${Object.keys(periods).length} periods`);
      }
    }
  });
  
  console.log(`External lookup built with ${periodsLookup.size} indicators`);
  
  // Process periods dataset with manual entry preservation
  const batchSize = 50;
  const totalRows = periodsDataset.getLastRow() - 1;
  const periodsHeaders = periodsDataset.getRange(1, 1, 1, periodsDataset.getLastColumn()).getValues()[0];
  
  const idIndex = periodsHeaders.indexOf('id_indicador');
  const tipoIndex = periodsHeaders.indexOf('tipo de dato');
  
  if (idIndex === -1 || tipoIndex === -1) {
    throw new Error(`Required columns not found in periods dataset "${periodsDatasetName}". Available columns: ${periodsHeaders.join(', ')}`);
  }
  
  console.log(`Periods dataset structure: id_indicador at ${idIndex}, tipo at ${tipoIndex}`);
  
  // Find period column indices in the periods dataset
  const periodColumnIndices = {};
  periodColumns.forEach(col => {
    const index = periodsHeaders.indexOf(col.target);
    if (index !== -1) {
      periodColumnIndices[col.target] = index;
    }
  });
  
  console.log('Period column indices:', periodColumnIndices);
  
  let totalMapped = 0;
  let totalComponents = 0;
  let manualEntriesPreserved = 0;
  let newManualEntrySlots = 0;
  let processedRows = 0;
  let debugSample = [];
  
  console.log(`Processing ${totalRows} rows in batches of ${batchSize}...`);
  
  // Process in batches
  for (let startRow = 2; startRow <= totalRows + 1; startRow += batchSize) {
    const endRow = Math.min(startRow + batchSize - 1, totalRows + 1);
    const batchRows = endRow - startRow + 1;
    
    const batchData = periodsDataset.getRange(startRow, 1, batchRows, periodsDataset.getLastColumn()).getValues();
    
    batchData.forEach((row, batchIndex) => {
      const actualRow = startRow + batchIndex;
      const id = row[idIndex];
      const tipo = row[tipoIndex];
      
      const normalizedId = String(id).trim();
      
      // Collect debug sample from first batch
      if (startRow === 2 && batchIndex < 5) {
        debugSample.push({ row: actualRow, id: normalizedId, tipo: tipo, hasLookup: periodsLookup.has(normalizedId) });
      }
      
      if (id === 0 || id === '0' || normalizedId === '0') {
        // Component without indicator - handle manual entries carefully
        totalComponents++;
        
        const firstPeriodColIndex = periodColumnIndices['Periodo_Base_EVA_Mayo'];
        if (firstPeriodColIndex !== undefined) {
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
        
      } else if (periodsLookup.has(normalizedId)) {
        // Found matching periods - update with external data
        const periods = periodsLookup.get(normalizedId);
        
        periodColumns.forEach(col => {
          const targetColIndex = periodColumnIndices[col.target];
          if (targetColIndex !== undefined && periods[col.target]) {
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
  
  // Enhanced logging with preservation info
  console.log(`Mapping complete: ${totalMapped} mapped, ${totalComponents} total components`);
  console.log(`Manual entries: ${manualEntriesPreserved} preserved, ${newManualEntrySlots} new slots marked`);
  console.log('Debug sample from first 5 rows:', debugSample);
  console.log('External lookup keys (first 10):', Array.from(periodsLookup.keys()).slice(0, 10));
  
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
 * MISSING HELPER FUNCTIONS - Add these to make smart refresh fully self-contained
 */

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
  
  // Define evaluation period columns
  const evaluationColumns = [
    'Periodo_Base_EVA_Mayo',
    'Periodo_Base_EVA_Noviembre', 
    'Diagnostico_Mayo_25',
    'Evaluacion_1_Noviembre_25',
    'Evaluacion_2_Mayo_26',
    'Evaluacion_3_Noviembre_26',
    'Evaluacion_4_Mayo_27',
    'Evaluacion_5_Octubre_27'
  ];
  
  // Create headers for the new dataset
  const newHeaders = [
    'municipio',
    'id_indicador', 
    'eje',
    'tema',
    'tipo de dato',
    'nombre',
    'codigo'
  ].concat(evaluationColumns);
  
  // Find source column indices
  const municipioIndex = sourceHeaders.indexOf('municipio');
  const indicadorIdIndex = sourceHeaders.indexOf('id_indicador');
  const ejeIndex = sourceHeaders.indexOf('eje');
  const temaIndex = sourceHeaders.indexOf('tema');
  const tipoIndex = sourceHeaders.indexOf('tipo de dato');
  const nombreIndex = sourceHeaders.indexOf('nombre');
  const codigoIndex = sourceHeaders.indexOf('codigo');
  
  console.log('Column indices:', { municipioIndex, indicadorIdIndex, ejeIndex, temaIndex, tipoIndex, nombreIndex, codigoIndex });
  
  if ([municipioIndex, indicadorIdIndex, tipoIndex, codigoIndex].includes(-1)) {
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
  destSheet.getRange(1, 8, 1, evaluationColumns.length).setBackground('#fff2cc'); // Period columns
  
  // Copy core data
  const resultData = [];
  sourceRows.forEach((row, index) => {
    const newRow = [
      getValueOrEmptyRobust(row[municipioIndex]),
      getValueOrZeroRobust(row[indicadorIdIndex]),
      getValueOrEmptyRobust(row[ejeIndex]),
      getValueOrEmptyRobust(row[temaIndex]),
      getValueOrEmptyRobust(row[tipoIndex]),
      getValueOrEmptyRobust(row[nombreIndex]),
      getValueOrEmptyRobust(row[codigoIndex])
    ];
    
    // Add empty evaluation period columns
    evaluationColumns.forEach(() => {
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
 * Helper function to get value or empty string (robust version)
 */
function getValueOrEmptyRobust(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value;
}

/**
 * Helper function to get value or 0 (robust version for id_indicador)
 */
function getValueOrZeroRobust(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  return value;
}

/**
 * Robust period text normalizer (standalone version)
 */
function normalizePeriodText(text) {
  if (!text || text === '') return '';
  
  const originalText = text.toLowerCase().trim();
  
  // Handle NA, N/A, and similar cases
  if (originalText === 'na' || originalText === 'n/a' || originalText === '-') {
    return 'NA';
  }
  
  // Handle empty/blank
  if (originalText === '') {
    return '';
  }
  
  // Handle specific patterns (add more as needed)
  if (originalText.includes('2do semestre') || originalText.includes('segundo semestre')) {
    const yearMatch = originalText.match(/202\d/);
    if (yearMatch) {
      return `junio ${yearMatch[0]}`;
    }
  }
  
  if (originalText.includes('1er semestre') || originalText.includes('primer semestre')) {
    const yearMatch = originalText.match(/202\d/);
    if (yearMatch) {
      const prevYear = parseInt(yearMatch[0]) - 1;
      return `diciembre ${prevYear}`;
    }
  }
  
  // Pattern: "enero-diciembre 2024" -> "enero 2024 - diciembre 2024"
  if (originalText.includes('-') && !originalText.includes(' - ')) {
    const parts = originalText.split('-').map(p => p.trim());
    if (parts.length === 2) {
      const yearMatch = originalText.match(/202\d/);
      if (yearMatch) {
        const year = yearMatch[0];
        const firstMonth = parts[0].trim();
        const secondPart = parts[1].trim();
        const secondMonth = secondPart.replace(/202\d/, '').trim();
        
        const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
                       'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
        
        if (months.includes(firstMonth) && months.includes(secondMonth)) {
          return `${firstMonth} ${year} - ${secondMonth} ${year}`;
        }
      }
    }
  }
  
  // Pattern: "octubre 2024 a marzo 2025" -> "octubre 2024 - marzo 2025"
  if (originalText.includes(' a ')) {
    return originalText.replace(' a ', ' - ');
  }
  
  // Pattern: Just year "2024" -> needs manual specification
  if (/^202\d$/.test(originalText)) {
    return `ERROR: "${text}" needs specific month`;
  }
  
  // Return as-is if already in correct format
  return originalText;
}

// =============================================
// MANUAL ENTRY PRESERVATION FUNCTIONS
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
  
  // Period columns to check
  const periodColumns = [
    'Periodo_Base_EVA_Mayo',
    'Periodo_Base_EVA_Noviembre', 
    'Diagnostico_Mayo_25',
    'Evaluacion_1_Noviembre_25',
    'Evaluacion_2_Mayo_26',
    'Evaluacion_3_Noviembre_26',
    'Evaluacion_4_Mayo_27',
    'Evaluacion_5_Octubre_27'
  ];
  
  rows.forEach((row, rowIndex) => {
    const municipio = row[municipioIndex];
    const idIndicador = row[idIndicadorIndex];
    const tipo = row[tipoIndex];
    const codigo = row[codigoIndex];
    
    // Only check components with id_indicador = 0 (manual entry rows)
    if (idIndicador === 0 || idIndicador === '0') {
      const rowKey = `${municipio}_${idIndicador}_${tipo}_${codigo}`;
      
      // Check each period column for manual entries
      periodColumns.forEach(periodCol => {
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
