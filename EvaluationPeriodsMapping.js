/**
 * EvaluationPeriodsMapping.gs
 * Complete, clean version for mapping periods data and generating calculations
 * Organized with flexible functions and optimized performance
 */

// =============================================
// MAIN MAPPING FUNCTIONS
// =============================================

/**
 * FLEXIBLE VERSION: Allows you to choose any sheet names
 * This is the recommended function for most use cases
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
 * Use this when you always use the same sheet names
 */
function mapPeriodsFromExternalTableUltraFast() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Your specific configuration - you can modify these
  const periodsDatasetName = 'Periodos_Evaluacion';
  const sourceDocId = '10AkW5RTc-MGXh26V2sXkGxBQ0cINRVFq0-AXaSWjiHc';
  const externalSheetName = 'Fechas_evaluación';
  
  console.log('=== ULTRA FAST PERIODS MAPPING ===');
  
  // Get sheets
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  if (!periodsDataset) {
    ui.alert('Error', `Sheet "${periodsDatasetName}" not found.`, ui.ButtonSet.OK);
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
 * CORE FUNCTION: Perform periods mapping with enhanced performance
 * This is the main function used by the menu system
 */
function performPeriodsMapping(sourceSheet, periodsDataset, ui) {
  try {
    console.log('=== PERIODS MAPPING STARTED ===');
    
    // Get source data
    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceHeaders = sourceData[0];
    const sourceRows = sourceData.slice(1);
    
    // Find ID column with flexible matching
    const sourceIdIndex = findIdColumn(sourceHeaders);
    if (sourceIdIndex === -1) {
      ui.alert('Error', 'Could not find ID Indicador column in source data.', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`Found ID column: "${sourceHeaders[sourceIdIndex]}" at index ${sourceIdIndex}`);
    
    // Find period columns
    const evaluationColumnMappings = findPeriodColumns(sourceHeaders);
    if (evaluationColumnMappings.length === 0) {
      ui.alert('Error', 'No evaluation period columns found in source data.', ui.ButtonSet.OK);
      return;
    }
    
    // Show confirmation
    const mappingInfo = `Found ${evaluationColumnMappings.length} evaluation period columns.\n\nProceed with mapping?`;
    const confirmResult = ui.alert('Confirm Column Mappings', mappingInfo, ui.ButtonSet.YES_NO);
    
    if (confirmResult !== ui.Button.YES) {
      return;
    }
    
    // Build lookup table from external data
    const periodsLookup = buildPeriodsLookup(sourceRows, sourceIdIndex, evaluationColumnMappings);
    console.log(`Built lookup for ${Object.keys(periodsLookup).length} indicators`);
    
    // Apply mappings to periods dataset
    const result = applyMappingsToDataset(periodsDataset, periodsLookup, evaluationColumnMappings);
    
    // Show results
    const successMessage = `✅ MAPPING COMPLETE!\n\n` +
      `• Periods mapped: ${result.mappedCount}\n` +
      `• Components for manual entry: ${result.componentsWithoutIndicator}\n` +
      `• External indicators: ${Object.keys(periodsLookup).length}\n\n` +
      `Next step: Generate calculation formulas.`;
    
    ui.alert('Periods Mapping Complete!', successMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Error in performPeriodsMapping:', error);
    ui.alert('Error', `Error during mapping: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * OPTIMIZED: Generate formulas from periods dataset to destination sheet
 */
function generateFormulasFromPeriodsDataToDestinationOptimized(periodsDataset, destinationSheet, selectedPeriod, ss) {
  const ui = SpreadsheetApp.getUi();
  
  console.log('Starting optimized formula generation...');
  const startTime = new Date();
  
  // Get data
  const periodsData = periodsDataset.getDataRange().getValues();
  const periodsHeaders = periodsData[0];
  const periodsRows = periodsData.slice(1);
  
  const destData = destinationSheet.getDataRange().getValues();
  const destHeaders = destData[0];
  const destRows = destData.slice(1);
  
  // Validate columns
  const periodsIndices = getColumnIndices(periodsHeaders, ['municipio', 'id_indicador', 'tipo de dato', 'codigo', selectedPeriod]);
  const destIndices = getColumnIndices(destHeaders, ['municipio', 'id_indicador', 'tipo de dato', 'codigo']);
  
  if (!periodsIndices.valid || !destIndices.valid) {
    ui.alert('Error', 'Required columns not found.', ui.ButtonSet.OK);
    return;
  }
  
  // Create periods lookup
  const periodsLookup = new Map();
  periodsRows.forEach(row => {
    const key = `${row[periodsIndices.municipio]}_${row[periodsIndices.id_indicador]}_${row[periodsIndices.tipo]}_${row[periodsIndices.codigo]}`;
    periodsLookup.set(key, row[periodsIndices[selectedPeriod]]);
  });
  
  // Get BD_componentes mapping
  const lookupMap = getBDComponentesMapping(ss);
  
  // Add calculation column
  const calculationColumnIndex = destHeaders.length + 1;
  const calculationColumnName = `Calculo_${selectedPeriod}`;
  
  destinationSheet.getRange(1, calculationColumnIndex).setValue(calculationColumnName);
  destinationSheet.getRange(1, calculationColumnIndex).setFontWeight('bold').setBackground('#e8f4fd');
  
  // Generate formulas in batch
  const formulas = generateFormulasBatch(destRows, destIndices, periodsLookup, lookupMap, calculationColumnIndex);
  
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

// =============================================
// BATCH PROCESSING FUNCTIONS
// =============================================

/**
 * Enhanced batch processing with detailed progress
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
  
  // Find period columns
  const periodColumns = findPeriodColumns(sourceHeaders);
  if (periodColumns.length === 0) {
    ui.alert('Error', `No period columns found in "${externalSheetName}".`, ui.ButtonSet.OK);
    return;
  }
  
  console.log(`Found ${periodColumns.length} period columns`);
  
  // Build periods lookup
  const periodsLookup = new Map();
  sourceRows.forEach(row => {
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
    }
  });
  
  console.log(`External lookup built with ${periodsLookup.size} indicators`);
  
  // Process periods dataset in batches
  const batchSize = 50;
  const totalRows = periodsDataset.getLastRow() - 1;
  const periodsHeaders = periodsDataset.getRange(1, 1, 1, periodsDataset.getLastColumn()).getValues()[0];
  
  const idIndex = periodsHeaders.indexOf('id_indicador');
  const tipoIndex = periodsHeaders.indexOf('tipo de dato');
  
  let totalMapped = 0;
  let totalComponents = 0;
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
      
      const normalizedId = String(id).trim();
      
      if (id === 0 || id === '0' || normalizedId === '0') {
        // Component without indicator
        totalComponents++;
        const firstPeriodCol = periodsHeaders.indexOf('Periodo_Base_EVA_Mayo');
        if (firstPeriodCol !== -1) {
          const cell = periodsDataset.getRange(actualRow, firstPeriodCol + 1);
          cell.setBackground('#ffcccb');
          cell.setNote('Manual entry needed for component without indicator');
        }
      } else if (periodsLookup.has(normalizedId)) {
        // Found matching periods
        const periods = periodsLookup.get(normalizedId);
        
        periodColumns.forEach(col => {
          const targetCol = periodsHeaders.indexOf(col.target);
          if (targetCol !== -1 && periods[col.target]) {
            periodsDataset.getRange(actualRow, targetCol + 1).setValue(periods[col.target]);
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
  
  // Success message
  ui.alert(
    'Mapping Complete!',
    `✅ SUCCESSFULLY COMPLETED!\n\n` +
    `• Source: "${externalSheetName}"\n` +
    `• Destination: "${periodsDatasetName}"\n` +
    `• Periods mapped: ${totalMapped}\n` +
    `• Components for manual entry: ${totalComponents}\n` +
    `• Total rows processed: ${processedRows}\n` +
    `• External indicators used: ${periodsLookup.size}\n\n` +
    `🎉 Your periods dataset is ready!`,
    ui.ButtonSet.OK
  );
}

// =============================================
// TESTING FUNCTIONS
// =============================================

/**
 * Test mapping with first 200 rows only
 */
function testMappingFirst200Rows() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask for configuration
  const periodsDatasetResult = ui.prompt(
    'Test Mapping Setup',
    'Enter your periods dataset sheet name:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodsDatasetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const periodsDatasetName = periodsDatasetResult.getResponseText().trim();
  
  const sourceDocResult = ui.prompt(
    'External Document ID',
    'Enter the external document ID:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sourceDocResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceDocId = sourceDocResult.getResponseText().trim();
  
  const externalSheetResult = ui.prompt(
    'External Sheet Name',
    'Enter the external sheet name:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (externalSheetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const externalSheetName = externalSheetResult.getResponseText().trim();
  
  console.log('=== TEST MAPPING - FIRST 200 ROWS ONLY ===');
  
  // Get sheets
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  let sourceSpreadsheet, externalSheet;
  
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
    externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  } catch (e) {
    ui.alert('Error', `Cannot access external document: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  // Build external lookup
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  const sourceIdIndex = findIdColumn(sourceHeaders);
  
  const periodsLookup = new Map();
  sourceRows.forEach(row => {
    const id = String(row[sourceIdIndex]).trim();
    if (id && id !== '') {
      periodsLookup.set(id, `TEST_PERIOD_FOR_${id}`);
    }
  });
  
  console.log(`Built lookup with ${periodsLookup.size} indicators`);
  
  // Process ONLY first 200 rows for testing
  const testRows = Math.min(200, periodsDataset.getLastRow() - 1);
  const periodsHeaders = periodsDataset.getRange(1, 1, 1, periodsDataset.getLastColumn()).getValues()[0];
  const testData = periodsDataset.getRange(2, 1, testRows, periodsDataset.getLastColumn()).getValues();
  
  const idIndex = periodsHeaders.indexOf('id_indicador');
  let mappedCount = 0;
  let componentsCount = 0;
  
  console.log(`Testing with first ${testRows} rows...`);
  
  testData.forEach((row, index) => {
    const actualRow = index + 2;
    const id = String(row[idIndex]).trim();
    
    if (id === '0' || id === '') {
      componentsCount++;
    } else if (periodsLookup.has(id)) {
      // Just set first period column as test
      const firstPeriodCol = periodsHeaders.indexOf('Periodo_Base_EVA_Mayo');
      if (firstPeriodCol !== -1) {
        periodsDataset.getRange(actualRow, firstPeriodCol + 1).setValue(`MAPPED_${id}`);
        mappedCount++;
      }
    }
  });
  
  console.log(`Test complete: ${mappedCount} mapped, ${componentsCount} components`);
  
  ui.alert(
    'Test Complete!',
    `Test mapping of first ${testRows} rows:\n\n` +
    `• Periods mapped: ${mappedCount}\n` +
    `• Components found: ${componentsCount}\n` +
    `• External indicators: ${periodsLookup.size}\n\n` +
    `If this looks good, run the full mapping.`,
    ui.ButtonSet.OK
  );
}

// =============================================
// HELPER FUNCTIONS
// =============================================

/**
 * Find ID column with flexible matching
 */
function findIdColumn(headers) {
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
 * Find period columns in headers
 */
function findPeriodColumns(headers) {
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
 * Normalize indicator ID for consistent comparison
 */
function normalizeIndicatorId(value) {
  if (value === null || value === undefined || value === '' || value === '...') {
    return null;
  }
  
  const stringValue = String(value).trim();
  if (stringValue === '') return null;
  
  const numValue = Number(stringValue);
  if (!isNaN(numValue) && isFinite(numValue)) {
    return String(Math.round(numValue));
  }
  
  return stringValue;
}

/**
 * Build periods lookup from external data
 */
function buildPeriodsLookup(sourceRows, sourceIdIndex, evaluationColumnMappings) {
  const periodsLookup = {};
  
  sourceRows.forEach((row, rowIndex) => {
    try {
      const rawId = row[sourceIdIndex];
      const normalizedId = normalizeIndicatorId(rawId);
      
      if (normalizedId !== null) {
        periodsLookup[normalizedId] = {};
        
        evaluationColumnMappings.forEach(mapping => {
          const rawValue = row[mapping.sourceIndex];
          if (rawValue && rawValue !== '' && rawValue !== '...') {
            periodsLookup[normalizedId][mapping.target] = normalizePeriodText(rawValue.toString());
          }
        });
      }
    } catch (error) {
      console.error(`Error processing external row ${rowIndex + 2}:`, error);
    }
  });
  
  return periodsLookup;
}

/**
 * Apply mappings to periods dataset
 */
function applyMappingsToDataset(periodsDataset, periodsLookup, evaluationColumnMappings) {
  const periodsData = periodsDataset.getDataRange().getValues();
  const periodsHeaders = periodsData[0];
  const periodsRows = periodsData.slice(1);
  
  const periodsIdIndex = periodsHeaders.indexOf('id_indicador');
  const periodsTipoIndex = periodsHeaders.indexOf('tipo de dato');
  
  let mappedCount = 0;
  let componentsWithoutIndicator = 0;
  
  periodsRows.forEach((row, rowIndex) => {
    try {
      const actualRowNum = rowIndex + 2;
      const rawId = row[periodsIdIndex];
      const tipo = row[periodsTipoIndex];
      const normalizedId = normalizeIndicatorId(rawId);
      
      if (rawId === 0 || rawId === '0' || normalizedId === '0') {
        // Component without indicator - mark for manual entry
        componentsWithoutIndicator++;
        const firstPeriodColIndex = periodsHeaders.indexOf('Periodo_Base_EVA_Mayo');
        if (firstPeriodColIndex !== -1) {
          const cell = periodsDataset.getRange(actualRowNum, firstPeriodColIndex + 1);
          cell.setNote('Manual entry needed for component without indicator');
          cell.setBackground('#ffcccb');
        }
      } else if (normalizedId && periodsLookup[normalizedId]) {
        // Found matching periods
        const indicatorPeriods = periodsLookup[normalizedId];
        
        evaluationColumnMappings.forEach(mapping => {
          const targetColIndex = periodsHeaders.indexOf(mapping.target);
          if (targetColIndex !== -1 && indicatorPeriods[mapping.target]) {
            periodsDataset.getRange(actualRowNum, targetColIndex + 1).setValue(indicatorPeriods[mapping.target]);
            mappedCount++;
          }
        });
      }
    } catch (error) {
      console.error(`Error processing periods row ${rowIndex + 2}:`, error);
    }
  });
  
  periodsDataset.autoResizeColumns(1, periodsHeaders.length);
  
  return { mappedCount, componentsWithoutIndicator };
}

/**
 * Get column indices for required columns
 */
function getColumnIndices(headers, requiredColumns) {
  const indices = {};
  let valid = true;
  
  requiredColumns.forEach(colName => {
    const index = headers.indexOf(colName);
    if (index === -1) {
      valid = false;
    }
    indices[colName.replace(' ', '_')] = index;
  });
  
  indices.valid = valid;
  return indices;
}

/**
 * Get BD_componentes mapping
 */
function getBDComponentesMapping(ss) {
  const lookupMap = new Map();
  const bdSheet = ss.getSheetByName('BD_componentes');
  
  if (bdSheet) {
    const bdData = bdSheet.getDataRange().getValues();
    const bdHeaders = bdData[0];
    const bdMunicipioIndex = bdHeaders.indexOf('municipio');
    const bdCodigoIndex = bdHeaders.indexOf('codigo_componente');
    
    if (bdMunicipioIndex >= 0 && bdCodigoIndex >= 0) {
      for (let i = 1; i < bdData.length; i++) {
        const municipio = bdData[i][bdMunicipioIndex];
        const codigo = bdData[i][bdCodigoIndex];
        if (municipio && codigo) {
          lookupMap.set(`${municipio}_${codigo}`, i + 1);
        }
      }
    }
  }
  
  return lookupMap;
}

/**
 * Generate formulas in batch
 */
function generateFormulasBatch(destRows, destIndices, periodsLookup, lookupMap, calculationColumnIndex) {
  const formulas = [];
  
  destRows.forEach((row, rowIndex) => {
    const actualRowNum = rowIndex + 2;
    const municipio = row[destIndices.municipio];
    const id_indicador = row[destIndices.id_indicador];
    const tipo = row[destIndices.tipo];
    const codigo = row[destIndices.codigo];
    
    if (!codigo || codigo === '') {
      formulas.push(['']);
      return;
    }
    
    const lookupKey = `${municipio}_${id_indicador}_${tipo}_${codigo}`;
    const period = periodsLookup.get(lookupKey);
    
    let formula = '';
    
    if (tipo === 'Componente') {
      if (!period || period === '' || period.includes('ERROR:') || period.includes('Manual entry needed') || period === 'NA') {
        formula = `=IFERROR("PERIOD_MISSING_FOR_${codigo}", "")`;
      } else {
        const formulaDateRange = convertNormalizedPeriodToFormula(period);
        const municipioCell = `${getColumnLetter(destIndices.municipio + 1)}${actualRowNum}`;
        
        const bdRow = lookupMap.get(`${municipio}_${codigo}`);
        const bdReference = bdRow ? `BD_componentes!${bdRow}:${bdRow}` : 'BD_componentes!1:1';
        
        if (formulaDateRange.includes('-')) {
          formula = `=CALCULATE_INDICATOR("SUM(${codigo}:${formulaDateRange})", ${municipioCell}, ${bdReference})`;
        } else {
          formula = `=CALCULATE_INDICATOR("SINGLE(${codigo}:${formulaDateRange})", ${municipioCell}, ${bdReference})`;
        }
      }
    } else if (tipo === 'Indicador') {
      formula = generateIndicatorCellFormulaForDestinationOptimized(codigo, actualRowNum, getColumnLetter(calculationColumnIndex), destRows, destIndices);
    }
    
    formulas.push([formula]);
  });
  
  return formulas;
}

/**
 * Generate indicator formula for destination sheet
 */
function generateIndicatorCellFormulaForDestinationOptimized(formulaText, currentRow, calculationColumn, rows, columnIndices) {
  const componentCodes = extractComponentCodes(formulaText);
  
  if (componentCodes.length === 0) {
    return '';
  }
  
  const currentRowData = rows[currentRow - 2];
  const currentMunicipio = currentRowData[columnIndices.municipio];
  
  const componentRefs = {};
  const componentLookup = new Map();
  
  rows.forEach((row, index) => {
    if (row[columnIndices.tipo] === 'Componente' && 
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
  
  componentCodes.forEach(code => {
    const componentInfo = componentLookup.get(code.toLowerCase());
    if (componentInfo) {
      componentRefs[code] = `${calculationColumn}${componentInfo.rowNum}`;
    }
  });
  
  let finalFormula = formulaText;
  Object.entries(componentRefs).forEach(([code, cellRef]) => {
    const regex = new RegExp(`\\b${code}\\b`, 'gi');
    finalFormula = finalFormula.replace(regex, cellRef);
  });
  
  if (!finalFormula.startsWith('=') && finalFormula.trim() !== '') {
    finalFormula = '=' + finalFormula;
  }
  
  return finalFormula;
}