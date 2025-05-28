/**
 * FAST DIAGNOSTIC: Check only first 50 rows to quickly identify the issue
 * Add this to your EvaluationPeriodsManager.gs file
 */
function fastPeriodsDatasetDiagnostic() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const periodsDatasetName = 'Periodos_Evaluacion';
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  
  if (!periodsDataset) {
    ui.alert('Error', `Sheet "${periodsDatasetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Get only first 51 rows (header + 50 data rows) for speed
  const range = periodsDataset.getRange(1, 1, Math.min(51, periodsDataset.getLastRow()), periodsDataset.getLastColumn());
  const data = range.getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  console.log('=== FAST DIAGNOSTIC (First 50 rows) ===');
  
  // Find columns
  const idIndex = headers.indexOf('id_indicador');
  const municipioIndex = headers.indexOf('municipio');
  const tipoIndex = headers.indexOf('tipo de dato');
  const codigoIndex = headers.indexOf('codigo');
  
  // Analyze first 50 rows only
  const foundIds = new Set();
  const sampleData = [];
  
  rows.forEach((row, index) => {
    const id = row[idIndex];
    const tipo = row[tipoIndex];
    const municipio = row[municipioIndex];
    const codigo = row[codigoIndex];
    
    // Track unique IDs
    if (id !== null && id !== undefined && id !== '' && id !== 0 && id !== '0') {
      foundIds.add(String(id));
    }
    
    // Collect sample data for first 10 rows
    if (index < 10) {
      sampleData.push({
        rowNum: index + 2,
        id: id,
        tipo: tipo,
        municipio: municipio,
        codigo: codigo
      });
    }
  });
  
  // Build quick report
  let report = `FAST DIAGNOSTIC RESULTS (First 50 rows):\n\n`;
  report += `Sheet: "${periodsDatasetName}"\n`;
  report += `Total rows in sheet: ${periodsDataset.getLastRow() - 1}\n`;
  report += `Rows analyzed: ${rows.length}\n\n`;
  
  report += `INDICATOR IDs FOUND IN FIRST 50 ROWS:\n`;
  if (foundIds.size === 0) {
    report += `❌ NO non-zero indicator IDs found in first 50 rows!\n\n`;
    report += `This means the first 50 rows either:\n`;
    report += `• All have id_indicador = 0 (components without indicators)\n`;
    report += `• All have empty/null id_indicador values\n`;
    report += `• There's a data structure issue\n\n`;
  } else {
    const sortedIds = Array.from(foundIds).sort((a, b) => Number(a) - Number(b));
    report += `Found ${foundIds.size} unique IDs: ${sortedIds.join(', ')}\n\n`;
    
    // Check if any match external range (1-37)
    const matchingIds = sortedIds.filter(id => {
      const num = Number(id);
      return num >= 1 && num <= 37;
    });
    
    if (matchingIds.length > 0) {
      report += `✅ MATCHES FOUND! IDs in external range (1-37): ${matchingIds.join(', ')}\n`;
      report += `These ${matchingIds.length} indicators should map successfully.\n\n`;
    } else {
      report += `❌ NO MATCHES! Your IDs (${sortedIds.join(', ')}) don't overlap with external IDs (1-37)\n\n`;
    }
  }
  
  report += `SAMPLE DATA (First 10 rows):\n`;
  sampleData.forEach(row => {
    report += `Row ${row.rowNum}: id=${row.id}, tipo="${row.tipo}", municipio="${row.municipio}", codigo="${row.codigo}"\n`;
  });
  
  report += `\nEXTERNAL SHEET COMPARISON:\n`;
  report += `External sheet has: IDs 1-37 (37 indicators total)\n`;
  
  if (foundIds.size === 0) {
    report += `\n🔍 RECOMMENDATION:\n`;
    report += `Check if your periods dataset was created correctly.\n`;
    report += `You might need to regenerate it from your combined dataset\n`;
    report += `where indicator IDs should be preserved properly.`;
  }
  
  ui.alert('Fast Diagnostic Results', report, ui.ButtonSet.OK);
  console.log('=== FAST DIAGNOSTIC COMPLETE ===');
}

/**
 * DETAILED DIAGNOSTIC: Analyze your periods dataset thoroughly
 * Add this function to your EvaluationPeriodsManager.gs file temporarily
 */
function detailedPeriodsDatasetDiagnostic() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Your specific sheet name
  const periodsDatasetName = 'Periodos_Evaluacion';
  
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  if (!periodsDataset) {
    ui.alert('Error', `Sheet "${periodsDatasetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Get data
  const data = periodsDataset.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  console.log('=== DETAILED PERIODS DATASET ANALYSIS ===');
  console.log('Headers:', headers);
  console.log('Total rows:', rows.length);
  
  // Find relevant columns
  const idIndex = headers.indexOf('id_indicador');
  const municipioIndex = headers.indexOf('municipio');
  const tipoIndex = headers.indexOf('tipo de dato');
  const codigoIndex = headers.indexOf('codigo');
  
  console.log('Column indices:');
  console.log('  id_indicador:', idIndex);
  console.log('  municipio:', municipioIndex);
  console.log('  tipo:', tipoIndex);
  console.log('  codigo:', codigoIndex);
  
  // Analyze ID distribution
  const idCounts = {};
  const sampleByIdAndType = {};
  
  rows.forEach((row, index) => {
    const id = row[idIndex];
    const tipo = row[tipoIndex];
    const municipio = row[municipioIndex];
    const codigo = row[codigoIndex];
    
    // Count IDs
    if (id !== null && id !== undefined && id !== '') {
      const normalizedId = String(id).trim();
      idCounts[normalizedId] = (idCounts[normalizedId] || 0) + 1;
      
      // Sample data for each ID and type
      const key = `${normalizedId}_${tipo}`;
      if (!sampleByIdAndType[key]) {
        sampleByIdAndType[key] = {
          id: normalizedId,
          tipo: tipo,
          municipio: municipio,
          codigo: codigo,
          rowNum: index + 2
        };
      }
    }
  });
  
  // Sort IDs numerically for better display
  const sortedIds = Object.keys(idCounts).sort((a, b) => {
    const numA = Number(a);
    const numB = Number(b);
    if (!isNaN(numA) && !isNaN(numB)) {
      return numA - numB;
    }
    return a.localeCompare(b);
  });
  
  // Build detailed report
  let report = `DETAILED PERIODS DATASET ANALYSIS:\n\n`;
  report += `Sheet: "${periodsDatasetName}"\n`;
  report += `Total rows: ${rows.length}\n`;
  report += `Headers: ${headers.join(', ')}\n\n`;
  
  report += `ID DISTRIBUTION ANALYSIS:\n`;
  if (sortedIds.length === 0) {
    report += `❌ NO INDICATOR IDs FOUND!\n`;
    report += `This means all id_indicador values are empty, null, or 0.\n\n`;
    
    // Show first few rows to see what's actually there
    report += `First 5 rows of actual data:\n`;
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i];
      report += `Row ${i + 2}: id_indicador="${row[idIndex]}", tipo="${row[tipoIndex]}", municipio="${row[municipioIndex]}"\n`;
    }
  } else {
    report += `Found ${sortedIds.length} unique indicator IDs:\n`;
    
    // Show ID counts
    sortedIds.forEach(id => {
      report += `• ID ${id}: ${idCounts[id]} rows\n`;
    });
    
    report += `\nSample data by ID and type:\n`;
    sortedIds.slice(0, 10).forEach(id => {
      const componentSample = sampleByIdAndType[`${id}_Componente`];
      const indicatorSample = sampleByIdAndType[`${id}_Indicador`];
      
      report += `\nID ${id}:\n`;
      if (componentSample) {
        report += `  Componente: ${componentSample.municipio} - ${componentSample.codigo} (row ${componentSample.rowNum})\n`;
      }
      if (indicatorSample) {
        report += `  Indicador: ${indicatorSample.municipio} - ${indicatorSample.codigo} (row ${indicatorSample.rowNum})\n`;
      }
    });
  }
  
  // Check for zero values specifically
  const zeroCount = rows.filter(row => row[idIndex] === 0 || row[idIndex] === '0').length;
  if (zeroCount > 0) {
    report += `\n⚠️ IMPORTANT: Found ${zeroCount} rows with id_indicador = 0\n`;
    report += `These are components without indicators (expected for manual entry).\n`;
  }
  
  // External comparison
  report += `\nCOMPARISON WITH EXTERNAL SHEET:\n`;
  report += `External sheet has IDs: 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 (up to 37)\n`;
  
  if (sortedIds.length > 0) {
    const commonIds = sortedIds.filter(id => {
      const num = Number(id);
      return !isNaN(num) && num >= 1 && num <= 37;
    });
    
    report += `Your dataset has IDs in range 1-37: ${commonIds.join(', ')}\n`;
    report += `Potential matches: ${commonIds.length}\n`;
    
    if (commonIds.length === 0) {
      report += `\n❌ PROBLEM: No overlapping IDs found!\n`;
      report += `Your periods dataset has IDs: ${sortedIds.slice(0, 10).join(', ')}\n`;
      report += `But external sheet expects IDs: 1-37\n`;
      report += `\nThis is a data synchronization issue.`;
    } else {
      report += `\n✅ Found ${commonIds.length} potentially matching IDs.\n`;
      report += `The mapping should work for these indicators.`;
    }
  }
  
  ui.alert('Detailed Periods Dataset Analysis', report, ui.ButtonSet.OK);
  console.log('=== DETAILED ANALYSIS COMPLETE ===');
}

/**
 * EvaluationPeriodsManager.gs
 * Main functions for managing evaluation periods
 * Creates periods datasets and maps evaluation dates from external sources
 * UPDATED VERSION with enhanced debugging capabilities
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
  
  if ([municipioIndex, indicadorIdIndex, tipoIndex, codigoIndex].includes(-1)) {
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
  destSheet.getRange(1, 8, 1, evaluationColumns.length).setBackground('#fff2cc'); // Period columns
  
  // Copy core data with PROPER handling of id_indicador = 0
  const resultData = [];
  
  sourceRows.forEach(row => {
    // CRITICAL FIX: Handle each column properly, especially id_indicador
    const newRow = [
      getValueOrEmpty(row[municipioIndex]),
      getValueOrZero(row[indicadorIdIndex]),  // ← FIXED: Use special handler for id_indicador
      getValueOrEmpty(row[ejeIndex]),
      getValueOrEmpty(row[temaIndex]),
      getValueOrEmpty(row[tipoIndex]),
      getValueOrEmpty(row[nombreIndex]),
      getValueOrEmpty(row[codigoIndex])
    ];
    
    // Add empty evaluation period columns
    evaluationColumns.forEach(() => {
      newRow.push('');
    });
    
    resultData.push(newRow);
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
    `• Evaluation period columns: ${evaluationColumns.length}\n\n` +
    `✓ Fixed: Components with id_indicador = 0 are properly preserved\n\n` +
    `Next step: Run "Map Periods from External Table" to populate the period columns.`,
    ui.ButtonSet.OK
  );
}

/**
 * Step 2: Map evaluation periods from external table to the periods dataset
 * ENHANCED with better error handling and debugging
 */
function mapPeriodsFromExternalTable() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask for the evaluation periods dataset sheet
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
  
  // ENHANCED: Better instructions for external document access
  const instructions = `To access an external Google Sheets document, you need:\n\n` +
    `1. The document ID (from the URL)\n` +
    `2. The external document must be shared with you\n` +
    `3. You must have at least "Viewer" access\n\n` +
    `Example URL:\n` +
    `https://docs.google.com/spreadsheets/d/ABC123DEF456/edit\n` +
    `Document ID: ABC123DEF456\n\n` +
    `Enter the document ID:`;
  
  const sourceDocResult = ui.prompt(
    'External Periods Table Document',
    instructions,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sourceDocResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceDocId = sourceDocResult.getResponseText().trim();
  
  // ENHANCED: Better error handling and alternative options
  let sourceSpreadsheet;
  try {
    // Test access to the external document
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
    
    // Try to read basic info to verify access
    const testSheets = sourceSpreadsheet.getSheets();
    if (testSheets.length === 0) {
      throw new Error('Document appears to have no sheets');
    }
    
  } catch (e) {
    console.error('External document access error:', e);
    
    // Provide detailed troubleshooting
    const errorMessage = `Cannot access external document.\n\n` +
      `Common issues:\n` +
      `• Document ID is incorrect\n` +
      `• Document is not shared with you\n` +
      `• You don't have sufficient permissions\n` +
      `• Document has been deleted or moved\n\n` +
      `Steps to fix:\n` +
      `1. Copy the document ID from the URL\n` +
      `2. Ask the document owner to share it with you\n` +
      `3. Ensure you have at least "Viewer" access\n\n` +
      `Would you like to try a different approach?`;
    
    const retryResult = ui.alert('Access Error', errorMessage, ui.ButtonSet.YES_NO);
    
    if (retryResult === ui.Button.YES) {
      // Offer alternative: manual upload option
      showManualUploadOption(periodsDataset, ui);
    }
    
    return;
  }
  
  // Get available sheets in external document
  const externalSheets = sourceSpreadsheet.getSheets().map(sheet => sheet.getName());
  const externalSheetsText = externalSheets.join(', ');
  
  const externalSheetResult = ui.prompt(
    'Select External Periods Table Sheet',
    `Available sheets in external document: ${externalSheetsText}\n\nEnter the sheet name:`,
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
  
  // Continue with the rest of the original mapping logic...
  try {
    performPeriodsMapping(externalSheet, periodsDataset, ui);
  } catch (e) {
    console.error('Mapping error:', e);
    ui.alert('Error', `Error during mapping: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * NEW: Debug function to help diagnose mapping issues
 * Run this when having problems with period mapping
 */
function debugPeriodsMapping() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask for the evaluation periods dataset sheet
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
  
  // Ask for external document info
  const sourceDocResult = ui.prompt(
    'External Document ID',
    'Enter the external document ID (e.g., 10AkW5RTc-MGXh26V2sXkGxBQ0cINRVFq0-AXaSWjiHc):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sourceDocResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceDocId = sourceDocResult.getResponseText().trim();
  
  const externalSheetResult = ui.prompt(
    'External Sheet Name',
    'Enter the external sheet name (e.g., Fechas_evaluación):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (externalSheetResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const externalSheetName = externalSheetResult.getResponseText().trim();
  
  let sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
  } catch (e) {
    ui.alert('Error', `Cannot access external document: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  const externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  if (!externalSheet) {
    ui.alert('Error', `Sheet "${externalSheetName}" not found in external document.`, ui.ButtonSet.OK);
    return;
  }
  
  // Get and analyze both datasets
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  const periodsData = periodsDataset.getDataRange().getValues();
  const periodsHeaders = periodsData[0];
  const periodsRows = periodsData.slice(1);
  
  console.log('=== DETAILED DEBUG ANALYSIS ===');
  console.log('External headers:', sourceHeaders);
  console.log('Periods headers:', periodsHeaders);
  
  // Find ID columns
  const sourceIdIndex = sourceHeaders.findIndex(h => 
    h.toLowerCase().includes('id') && h.toLowerCase().includes('indicador')
  );
  const periodsIdIndex = periodsHeaders.indexOf('id_indicador');
  
  console.log('Source ID column:', sourceHeaders[sourceIdIndex], 'at index', sourceIdIndex);
  console.log('Periods ID column: id_indicador at index', periodsIdIndex);
  
  // Helper function for normalizing IDs (same as in mapping function)
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
  
  // Analyze ID values
  const externalIds = new Set();
  const periodsIds = new Set();
  
  // Get unique IDs from external sheet
  sourceRows.forEach((row, index) => {
    const rawId = row[sourceIdIndex];
    const normalizedId = normalizeIndicatorId(rawId);
    if (normalizedId !== null) {
      externalIds.add(normalizedId);
    }
    if (index < 5) {
      console.log(`External row ${index + 2}: "${rawId}" -> "${normalizedId}"`);
    }
  });
  
  // Get unique IDs from periods sheet (excluding 0)
  periodsRows.forEach((row, index) => {
    const rawId = row[periodsIdIndex];
    const normalizedId = normalizeIndicatorId(rawId);
    if (normalizedId !== null && normalizedId !== '0') {
      periodsIds.add(normalizedId);
    }
    if (index < 5) {
      console.log(`Periods row ${index + 2}: "${rawId}" -> "${normalizedId}"`);
    }
  });
  
  // Find overlaps
  const matches = [];
  externalIds.forEach(id => {
    if (periodsIds.has(id)) {
      matches.push(id);
    }
  });
  
  // Build detailed report
  let report = `DETAILED DEBUG REPORT:\n\n`;
  report += `External Sheet Analysis:\n`;
  report += `• Document ID: ${sourceDocId}\n`;
  report += `• Sheet: ${externalSheetName}\n`;
  report += `• ID Column: "${sourceHeaders[sourceIdIndex]}" (index ${sourceIdIndex})\n`;
  report += `• Total rows: ${sourceRows.length}\n`;
  report += `• Unique IDs: ${externalIds.size}\n`;
  report += `• Sample IDs: ${Array.from(externalIds).slice(0, 10).join(', ')}\n\n`;
  
  report += `Periods Dataset Analysis:\n`;
  report += `• Sheet: ${periodsDatasetName}\n`;
  report += `• ID Column: id_indicador (index ${periodsIdIndex})\n`;
  report += `• Total rows: ${periodsRows.length}\n`;
  report += `• Unique IDs (non-zero): ${periodsIds.size}\n`;
  report += `• Sample IDs: ${Array.from(periodsIds).slice(0, 10).join(', ')}\n\n`;
  
  report += `MATCHING ANALYSIS:\n`;
  report += `• Matches found: ${matches.length}\n`;
  if (matches.length > 0) {
    report += `• Matching IDs: ${matches.slice(0, 20).join(', ')}\n`;
    report += `\n✅ You should have ${matches.length} successful mappings!`;
  } else {
    report += `• No matches found - this is the problem!\n\n`;
    report += `Troubleshooting:\n`;
    report += `• Check if IDs are in same format in both sheets\n`;
    report += `• Look for extra spaces or special characters\n`;
    report += `• Verify column names are correct\n`;
    report += `• Check data types (number vs text)\n\n`;
    
    // Show first few IDs from each sheet for comparison
    const firstExternalIds = Array.from(externalIds).slice(0, 5);
    const firstPeriodsIds = Array.from(periodsIds).slice(0, 5);
    
    report += `First 5 external IDs: ${firstExternalIds.join(', ')}\n`;
    report += `First 5 periods IDs: ${firstPeriodsIds.join(', ')}\n\n`;
    report += `Compare these carefully - they should match exactly!`;
  }
  
  ui.alert('Debug Report', report, ui.ButtonSet.OK);
  console.log('=== DEBUG COMPLETE ===');
}

/**
 * MODIFIED: Generate calculation formulas using the evaluation periods dataset
 * Now asks where to put the calculations (separate from periods dataset)
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
  const availablePeriods = [
    'Periodo_Base_EVA_Mayo',
    'Periodo_Base_EVA_Noviembre', 
    'Diagnostico_Mayo_25',
    'Evaluacion_1_Noviembre_25',
    'Evaluacion_2_Mayo_26',
    'Evaluacion_3_Noviembre_26',
    'Evaluacion_4_Mayo_27',
    'Evaluacion_5_Octubre_27'
  ];
  
  const periodResult = ui.prompt(
    'Select Evaluation Period',
    `Available periods:\n${availablePeriods.join('\n')}\n\nEnter the period name:`,
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
  
  // Generate formulas using the periods dataset but put them in calculations sheet
  generateFormulasFromPeriodsDataToDestinationOptimized(periodsDataset, calculationsSheet, selectedPeriod, ss);
}

/**
 * Map periods from local template sheet
 */
function mapPeriodsFromLocalTemplate() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if template exists
  const templateSheet = ss.getSheetByName('Periodos_Template');
  if (!templateSheet) {
    ui.alert('Error', 'Template sheet "Periodos_Template" not found. Please create it first.', ui.ButtonSet.OK);
    return;
  }
  
  // Ask for the evaluation periods dataset
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
  
  // Use the local template as the external source
  performPeriodsMapping(templateSheet, periodsDataset, ui);
}

/**
 * QUICK FIX: One-click debug for your specific case
 * Use this to quickly test your exact scenario
 */
function quickDebugYourCase() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Your specific values
  const periodsDatasetName = 'Periodos_Evaluacion'; // Adjust if different
  const sourceDocId = '10AkW5RTc-MGXh26V2sXkGxBQ0cINRVFq0-AXaSWjiHc';
  const externalSheetName = 'Fechas_evaluación';
  
  // Confirm these are correct
  const confirmResult = ui.alert(
    'Quick Debug Your Case',
    `This will debug your specific setup:\n\n` +
    `• Periods Dataset: "${periodsDatasetName}"\n` +
    `• External Doc ID: ${sourceDocId}\n` +
    `• External Sheet: "${externalSheetName}"\n\n` +
    `Are these values correct?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    ui.alert('Info', 'Please update the values in the quickDebugYourCase function or use the regular debugPeriodsMapping function.', ui.ButtonSet.OK);
    return;
  }
  
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  if (!periodsDataset) {
    ui.alert('Error', `Sheet "${periodsDatasetName}" not found. Please check the name.`, ui.ButtonSet.OK);
    return;
  }
  
  let sourceSpreadsheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
  } catch (e) {
    ui.alert('Error', `Cannot access external document: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  const externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  if (!externalSheet) {
    ui.alert('Error', `Sheet "${externalSheetName}" not found in external document.`, ui.ButtonSet.OK);
    return;
  }
  
  // Run the debug analysis
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  const periodsData = periodsDataset.getDataRange().getValues();
  const periodsHeaders = periodsData[0];
  const periodsRows = periodsData.slice(1);
  
  console.log('=== QUICK DEBUG FOR YOUR CASE ===');
  console.log('External headers:', sourceHeaders);
  console.log('Periods headers:', periodsHeaders);
  
  // Find ID columns
  const sourceIdIndex = sourceHeaders.findIndex(h => 
    h.toLowerCase().includes('id') && h.toLowerCase().includes('indicador')
  );
  const periodsIdIndex = periodsHeaders.indexOf('id_indicador');
  
  // Helper function for normalizing IDs
  function normalizeId(value) {
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
  
  // Sample first 10 IDs from each sheet
  console.log('=== FIRST 10 IDS FROM EACH SHEET ===');
  console.log('EXTERNAL SHEET:');
  const externalIds = [];
  for (let i = 0; i < Math.min(10, sourceRows.length); i++) {
    const rawId = sourceRows[i][sourceIdIndex];
    const normalizedId = normalizeId(rawId);
    externalIds.push(normalizedId);
    console.log(`Row ${i + 2}: "${rawId}" -> "${normalizedId}"`);
  }
  
  console.log('PERIODS DATASET:');
  const periodsIds = [];
  for (let i = 0; i < Math.min(10, periodsRows.length); i++) {
    const rawId = periodsRows[i][periodsIdIndex];
    const normalizedId = normalizeId(rawId);
    if (normalizedId !== '0') { // Exclude 0 values
      periodsIds.push(normalizedId);
    }
    console.log(`Row ${i + 2}: "${rawId}" -> "${normalizedId}"`);
  }
  
  // Check for matches in these samples
  const sampleMatches = externalIds.filter(id => id && periodsIds.includes(id));
  
  let report = `QUICK DEBUG RESULTS:\n\n`;
  report += `External Sheet "${externalSheetName}":\n`;
  report += `• ID Column: "${sourceHeaders[sourceIdIndex]}" (index ${sourceIdIndex})\n`;
  report += `• Total rows: ${sourceRows.length}\n`;
  report += `• First 10 IDs: ${externalIds.join(', ')}\n\n`;
  
  report += `Periods Dataset "${periodsDatasetName}":\n`;
  report += `• ID Column: "id_indicador" (index ${periodsIdIndex})\n`;
  report += `• Total rows: ${periodsRows.length}\n`;
  report += `• First 10 IDs (non-zero): ${periodsIds.join(', ')}\n\n`;
  
  report += `SAMPLE MATCHES: ${sampleMatches.length}\n`;
  if (sampleMatches.length > 0) {
    report += `Matching IDs: ${sampleMatches.join(', ')}\n\n`;
    report += `✅ Good news! Found matches in sample data.\n`;
    report += `The mapping should work. If it's not working, there might be an issue with the normalization logic.`;
  } else {
    report += `❌ No matches found in sample data!\n\n`;
    report += `This explains why mapping is failing.\n\n`;
    report += `Check:\n`;
    report += `• Are the ID formats exactly the same?\n`;
    report += `• Any extra spaces or characters?\n`;
    report += `• Are both using numbers or both using text?\n`;
  }
  
  ui.alert('Quick Debug Results', report, ui.ButtonSet.OK);
  console.log('=== QUICK DEBUG COMPLETE ===');
}

/**
 * DEBUG: Focus on the actual mapping logic between external sheet and periods dataset
 * Add this to your EvaluationPeriodsManager.gs file
 */
function debugMappingLogic() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Your specific values
  const periodsDatasetName = 'Periodos_Evaluacion';
  const sourceDocId = '10AkW5RTc-MGXh26V2sXkGxBQ0cINRVFq0-AXaSWjiHc';
  const externalSheetName = 'Fechas_evaluación';
  
  console.log('=== MAPPING LOGIC DEBUG ===');
  
  // Get periods dataset
  const periodsDataset = ss.getSheetByName(periodsDatasetName);
  if (!periodsDataset) {
    ui.alert('Error', `Sheet "${periodsDatasetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  // Get external sheet
  let sourceSpreadsheet, externalSheet;
  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
    externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
  } catch (e) {
    ui.alert('Error', `Cannot access external document: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  // Get data from external sheet
  const sourceData = externalSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  
  // Find ID column in external sheet
  const sourceIdIndex = sourceHeaders.findIndex(h => 
    h.toLowerCase().includes('id') && h.toLowerCase().includes('indicador')
  );
  
  console.log('External sheet ID column:', sourceHeaders[sourceIdIndex], 'at index', sourceIdIndex);
  
  // Build external lookup exactly like the mapping function does
  const periodsLookup = {};
  
  sourceRows.forEach((row, rowIndex) => {
    const rawIndicatorId = row[sourceIdIndex];
    
    // Use the same normalization logic as the actual mapping function
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
    
    const normalizedIndicatorId = normalizeIndicatorId(rawIndicatorId);
    
    if (normalizedIndicatorId !== null) {
      console.log(`External row ${rowIndex + 2}: "${rawIndicatorId}" -> "${normalizedIndicatorId}"`);
      periodsLookup[normalizedIndicatorId] = `PERIODS_FOR_${normalizedIndicatorId}`;
    }
  });
  
  console.log('External periods lookup created for IDs:', Object.keys(periodsLookup));
  
  // Now check periods dataset - look for indicators with non-zero IDs
  const periodsData = periodsDataset.getRange(1, 1, Math.min(100, periodsDataset.getLastRow()), periodsDataset.getLastColumn()).getValues();
  const periodsHeaders = periodsData[0];
  const periodsRows = periodsData.slice(1);
  
  const periodsIdIndex = periodsHeaders.indexOf('id_indicador');
  const periodsTipoIndex = periodsHeaders.indexOf('tipo de dato');
  const periodsCodigoIndex = periodsHeaders.indexOf('codigo');
  const periodsMunicipioIndex = periodsHeaders.indexOf('municipio');
  
  console.log('Periods dataset ID column index:', periodsIdIndex);
  
  // Look for non-zero indicator IDs in periods dataset
  const periodsIndicatorIds = new Set();
  const sampleMatches = [];
  let zeroCount = 0;
  let nonZeroCount = 0;
  
  periodsRows.forEach((row, rowIndex) => {
    const rawIndicatorId = row[periodsIdIndex];
    const tipo = row[periodsTipoIndex];
    const codigo = row[periodsCodigoIndex];
    const municipio = row[periodsMunicipioIndex];
    
    // Same normalization as mapping function
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
    
    const normalizedIndicatorId = normalizeIndicatorId(rawIndicatorId);
    
    if (rawIndicatorId === 0 || rawIndicatorId === '0' || normalizedIndicatorId === '0') {
      zeroCount++;
    } else if (normalizedIndicatorId !== null) {
      nonZeroCount++;
      periodsIndicatorIds.add(normalizedIndicatorId);
      
      // Check if this ID exists in external lookup
      if (periodsLookup[normalizedIndicatorId] && sampleMatches.length < 5) {
        sampleMatches.push({
          id: normalizedIndicatorId,
          tipo: tipo,
          municipio: municipio,
          codigo: codigo,
          rowNum: rowIndex + 2
        });
      }
      
      if (rowIndex < 10) {
        console.log(`Periods row ${rowIndex + 2}: id="${rawIndicatorId}" -> "${normalizedIndicatorId}", tipo="${tipo}"`);
      }
    }
  });
  
  // Find actual matches
  const actualMatches = [];
  periodsIndicatorIds.forEach(id => {
    if (periodsLookup[id]) {
      actualMatches.push(id);
    }
  });
  
  // Build report
  let report = `MAPPING LOGIC DEBUG RESULTS:\n\n`;
  report += `EXTERNAL SHEET:\n`;
  report += `• Found ${Object.keys(periodsLookup).length} indicators\n`;
  report += `• IDs: ${Object.keys(periodsLookup).join(', ')}\n\n`;
  
  report += `PERIODS DATASET (first 100 rows analyzed):\n`;
  report += `• Rows with id_indicador = 0: ${zeroCount} (correct - for manual entry)\n`;
  report += `• Rows with non-zero id_indicador: ${nonZeroCount}\n`;
  report += `• Unique non-zero IDs: ${Array.from(periodsIndicatorIds).join(', ')}\n\n`;
  
  report += `MATCHING RESULTS:\n`;
  report += `• External IDs available: ${Object.keys(periodsLookup).join(', ')}\n`;
  report += `• Periods IDs found: ${Array.from(periodsIndicatorIds).join(', ')}\n`;
  report += `• Actual matches: ${actualMatches.join(', ')}\n`;
  report += `• Total matches: ${actualMatches.length}\n\n`;
  
  if (actualMatches.length === 0) {
    report += `❌ PROBLEM IDENTIFIED:\n`;
    if (periodsIndicatorIds.size === 0) {
      report += `No non-zero indicator IDs found in first 100 rows of periods dataset.\n`;
      report += `This suggests all rows analyzed are components with id_indicador = 0.\n\n`;
      report += `SOLUTION: Look further in the dataset for indicator rows, or\n`;
      report += `verify that your periods dataset contains both components AND indicators.`;
    } else {
      report += `Your periods dataset has IDs: ${Array.from(periodsIndicatorIds).join(', ')}\n`;
      report += `But external sheet has IDs: ${Object.keys(periodsLookup).join(', ')}\n`;
      report += `These don't match - this is a data synchronization issue.`;
    }
  } else {
    report += `✅ MATCHES FOUND!\n`;
    report += `Sample matching rows:\n`;
    sampleMatches.forEach(match => {
      report += `• Row ${match.rowNum}: ID ${match.id}, ${match.tipo}, ${match.municipio}\n`;
    });
    report += `\nThe mapping should work for these ${actualMatches.length} indicators.`;
  }
  
  ui.alert('Mapping Logic Debug Results', report, ui.ButtonSet.OK);
  console.log('=== MAPPING DEBUG COMPLETE ===');
}