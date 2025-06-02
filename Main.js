/**
 * Main.js
 * Main entry point and menu system for Municipality Tools
 * Dependencies: All other modules
 */

/**
 * @NotOnlyCurrentDoc
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Municipality Tools')
    .addItem('🔍 Test Document Access', 'testDocumentAccess')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Dataset Management')
      .addItem('Generate Components Dataset', 'showGenerateComponentsDialog')
      .addItem('Refresh Components Dataset', 'refreshDatasetData')
      .addItem('Backup Components Dataset', 'backupCurrentDataset')
      .addSeparator()
      .addItem('Generate Indicators Dataset', 'showGenerateIndicatorsDialog')
      .addItem('Refresh Indicators Dataset', 'refreshIndicatorsDataset')
      .addSeparator()
      .addItem('Preview Column Mapping', 'previewColumnMapping')
      .addItem('Combine Components & Indicators', 'concatAndOrderByIndicator')
      .addItem('Refresh Combined Dataset', 'refreshConcatenatedDataset')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Evaluation Periods')
      .addItem('Create Evaluation Periods Dataset', 'createEvaluationPeriodsDataset')
      .addSeparator()
      .addItem('⚡ Smart Refresh All Periods Data', 'refreshAllPeriodsData')
      .addItem('⚙️ Configure Smart Refresh', 'configureSmartRefresh')
      .addSeparator()
      .addItem('Map Periods (Flexible - Choose Sheets)', 'mapPeriodsFromExternalTableFlexible')
      .addItem('Map Periods (Quick - Preset Names)', 'mapPeriodsFromExternalTableUltraFast')
      .addSeparator()
      .addItem('Generate Calculations from Periods Dataset', 'generateCalculationsFromPeriodsDataset')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🧮 Calculations & Formulas')
      .addItem('🔧 Configure Components Data Sheet', 'configureComponentsDataSheet')
      .addSeparator()
      .addItem('Create Formula Builder', 'createFormulaBuilder')
      .addItem('Setup Auto-Refresh (Every 5 Minutes)', 'createDataRefreshTrigger')
      .addItem('Refresh Calculations Now', 'manualRefreshCalculations')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Testing & Debug')
      .addItem('Test External Document Access', 'testExternalDocumentAccess')
      .addItem('Test Component Value', 'testComponentValue')
      .addItem('Test Component Sum', 'testComponentSum')
      .addItem('🔬 Debug Formula Generation', 'debugFormulaGeneration')
      .addSeparator()
      .addItem('Fast Periods Dataset Diagnostic', 'fastPeriodsDatasetDiagnostic')
      .addItem('Detailed Periods Dataset Analysis', 'detailedPeriodsDatasetDiagnostic')
      .addItem('Debug Periods Mapping', 'debugPeriodsMapping')
    )
    .addToUi();
}

// =============================================
// CORE HELPER FUNCTIONS
// =============================================

/**
 * Helper functions that are used across multiple modules
 */
function getSheetNames() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => sheet.getName());
}

/**
 * Test external document access with flexible document ID
 */
function testExternalDocumentAccess() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask for document ID
  const docIdResult = ui.prompt(
    'Test External Document Access',
    'Enter the external Google Sheets document ID to test:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (docIdResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceDocId = docIdResult.getResponseText().trim();
  
  // Ask for sheet name
  const sheetNameResult = ui.prompt(
    'External Sheet Name',
    'Enter the sheet name to test (e.g., "Fechas_evaluación"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetNameResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const externalSheetName = sheetNameResult.getResponseText().trim();
  
  try {
    ui.alert('Testing...', 'Testing access to your external document. Please wait.', ui.ButtonSet.OK);
    
    // Test access
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceDocId);
    const spreadsheetName = sourceSpreadsheet.getName();
    
    // Test sheet access
    const externalSheet = sourceSpreadsheet.getSheetByName(externalSheetName);
    if (!externalSheet) {
      ui.alert('Sheet Not Found', `The sheet "${externalSheetName}" was not found in the external document.`, ui.ButtonSet.OK);
      return;
    }
    
    // Test data reading
    const data = externalSheet.getDataRange().getValues();
    const headers = data[0];
    const rowCount = data.length - 1;
    
    // Find ID column using CommonHelpers
    const idColumnIndex = findIdColumn(headers);
    
    let message = `✅ External Document Access Test SUCCESSFUL!\n\n`;
    message += `Document Details:\n`;
    message += `• Document ID: ${sourceDocId}\n`;
    message += `• Document Name: "${spreadsheetName}"\n`;
    message += `• Sheet: "${externalSheetName}"\n`;
    message += `• Total rows: ${rowCount}\n`;
    message += `• Headers found: ${headers.length}\n`;
    message += `• Headers: ${headers.slice(0, 5).join(', ')}${headers.length > 5 ? '...' : ''}\n\n`;
    
    if (idColumnIndex >= 0) {
      message += `ID Column Analysis:\n`;
      message += `• Found ID column: "${headers[idColumnIndex]}" at index ${idColumnIndex}\n`;
      
      // Show first few ID values
      const sampleIds = [];
      for (let i = 1; i <= Math.min(5, rowCount); i++) {
        const id = data[i][idColumnIndex];
        sampleIds.push(`"${id}"`);
      }
      message += `• Sample IDs: ${sampleIds.join(', ')}\n\n`;
      message += `✅ Your external document is ready for periods mapping!`;
    } else {
      message += `⚠️ Warning: Could not find ID Indicador column.\n`;
      message += `Available columns: ${headers.join(', ')}\n\n`;
      message += `Make sure your external sheet has a column with "ID" and "Indicador" in the name.`;
    }
    
    ui.alert('External Document Test Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    let errorMessage = `❌ External Document Access FAILED!\n\n`;
    errorMessage += `Error: ${error.message}\n\n`;
    errorMessage += `Common solutions:\n`;
    errorMessage += `• Make sure the document is shared with you\n`;
    errorMessage += `• Check that you have at least "Viewer" access\n`;
    errorMessage += `• Verify the document ID is correct\n`;
    errorMessage += `• Ensure the document hasn't been deleted\n\n`;
    errorMessage += `Document ID tested: ${sourceDocId}`;
    
    ui.alert('External Document Test Failed', errorMessage, ui.ButtonSet.OK);
  }
}

// =============================================
// DIAGNOSTIC FUNCTIONS (FROM PERIODS MANAGER)
// =============================================

/**
 * FAST DIAGNOSTIC: Check only first 50 rows to quickly identify issues
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
  
  // Find columns using CommonHelpers
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
    
    // Track unique IDs using CommonHelpers
    const normalizedId = normalizeIndicatorId(id);
    if (normalizedId !== null && normalizedId !== '0') {
      foundIds.add(String(normalizedId));
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
 * DETAILED DIAGNOSTIC: Analyze periods dataset thoroughly
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
  
  // Analyze ID distribution using CommonHelpers
  const idCounts = {};
  const sampleByIdAndType = {};
  
  rows.forEach((row, index) => {
    const id = row[idIndex];
    const tipo = row[tipoIndex];
    const municipio = row[municipioIndex];
    const codigo = row[codigoIndex];
    
    // Count IDs using normalized format
    const normalizedId = normalizeIndicatorId(id);
    if (normalizedId !== null) {
      const idStr = String(normalizedId);
      idCounts[idStr] = (idCounts[idStr] || 0) + 1;
      
      // Sample data for each ID and type
      const key = `${idStr}_${tipo}`;
      if (!sampleByIdAndType[key]) {
        sampleByIdAndType[key] = {
          id: idStr,
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
    report += `This means all id_indicador values are empty, null, or invalid.\n\n`;
    
    // Show first few rows to see what's actually there
    report += `First 5 rows of actual data:\n`;
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const row = rows[i];
      report += `Row ${i + 2}: id_indicador="${row[idIndex]}", tipo="${row[tipoIndex]}", municipio="${row[municipioIndex]}"\n`;
    }
  } else {
    report += `Found ${sortedIds.length} unique indicator IDs:\n`;
    
    // Show ID counts
    sortedIds.slice(0, 15).forEach(id => {
      report += `• ID ${id}: ${idCounts[id]} rows\n`;
    });
    
    if (sortedIds.length > 15) {
      report += `... and ${sortedIds.length - 15} more IDs\n`;
    }
    
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
  const zeroCount = rows.filter(row => {
    const normalizedId = normalizeIndicatorId(row[idIndex]);
    return normalizedId === '0';
  }).length;
  
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
 * Debug periods mapping functionality
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
  
  // Find ID columns using CommonHelpers
  const sourceIdIndex = findIdColumn(sourceHeaders);
  const periodsIdIndex = periodsHeaders.indexOf('id_indicador');
  
  console.log('Source ID column:', sourceHeaders[sourceIdIndex], 'at index', sourceIdIndex);
  console.log('Periods ID column: id_indicador at index', periodsIdIndex);
  
  // Analyze ID values using CommonHelpers
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
 * Debug function for formula generation issues
 */
function debugFormulaGeneration() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ask for the evaluation periods dataset
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  const periodsDatasetResult = ui.prompt(
    'Debug Formula Generation - Select Periods Dataset',
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
  
  // Ask for the period to debug
  const periodResult = ui.prompt(
    'Debug Formula Generation - Select Period',
    `Available periods:\n${EVALUATION_PERIOD_COLUMNS.join('\n')}\n\nEnter the period name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const selectedPeriod = periodResult.getResponseText().trim();
  
  try {
    // Get data
    const periodsData = periodsDataset.getDataRange().getValues();
    const periodsHeaders = periodsData[0];
    const periodsRows = periodsData.slice(1);
    
    console.log('=== FORMULA GENERATION DEBUG ===');
    console.log('Periods dataset:', periodsDatasetName);
    console.log('Selected period:', selectedPeriod);
    console.log('Headers:', periodsHeaders);
    
    // Check if period exists
    const selectedPeriodIndex = periodsHeaders.indexOf(selectedPeriod);
    console.log('Selected period index:', selectedPeriodIndex);
    
    if (selectedPeriodIndex === -1) {
      ui.alert(
        'Period Not Found',
        `Period "${selectedPeriod}" not found in periods dataset.\n\n` +
        `Available periods: ${periodsHeaders.filter(h => EVALUATION_PERIOD_COLUMNS.includes(h)).join(', ')}`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Find column indices
    const periodsIndices = getColumnIndices(periodsHeaders, ['municipio', 'id_indicador', 'tipo de dato', 'codigo']);
    console.log('Column indices:', periodsIndices);
    
    if (!periodsIndices.valid) {
      ui.alert('Error', 'Required columns not found in periods dataset.', ui.ButtonSet.OK);
      return;
    }
    
    // Create periods lookup and analyze first 10 rows
    const periodsLookup = new Map();
    const analysisData = [];
    
    for (let i = 0; i < Math.min(10, periodsRows.length); i++) {
      const row = periodsRows[i];
      const key = `${row[periodsIndices.municipio]}_${row[periodsIndices.id_indicador]}_${row[periodsIndices.tipo_de_dato]}_${row[periodsIndices.codigo]}`;
      const periodValue = row[selectedPeriodIndex];
      
      periodsLookup.set(key, periodValue);
      
      analysisData.push({
        rowNum: i + 2,
        municipio: row[periodsIndices.municipio],
        id_indicador: row[periodsIndices.id_indicador],
        tipo: row[periodsIndices.tipo_de_dato],
        codigo: row[periodsIndices.codigo],
        period: periodValue,
        key: key
      });
      
      console.log(`Row ${i + 2}: ${key} -> "${periodValue}"`);
    }
    
    // Build detailed report
    let report = `FORMULA GENERATION DEBUG REPORT:\n\n`;
    report += `Periods Dataset: "${periodsDatasetName}"\n`;
    report += `Selected Period: "${selectedPeriod}"\n`;
    report += `Period Column Index: ${selectedPeriodIndex}\n`;
    report += `Total Rows: ${periodsRows.length}\n\n`;
    
    report += `FIRST 10 ROWS ANALYSIS:\n`;
    analysisData.forEach(data => {
      report += `Row ${data.rowNum}: ${data.municipio} | ${data.tipo} | ${data.codigo}\n`;
      report += `  Period Value: "${data.period}"\n`;
      report += `  Lookup Key: ${data.key}\n\n`;
    });
    
    report += `COMPONENTS DATA SHEET CONFIGURATION:\n`;
    const scriptProperties = PropertiesService.getScriptProperties();
    const componentsDataSheet = scriptProperties.getProperty('COMPONENTS_DATA_SHEET_NAME');
    if (componentsDataSheet) {
      report += `✅ Configured: "${componentsDataSheet}"\n`;
      
      // Test if sheet exists
      const testSheet = ss.getSheetByName(componentsDataSheet);
      if (testSheet) {
        report += `✅ Sheet exists and accessible\n`;
      } else {
        report += `❌ Sheet not found - reconfigure needed\n`;
      }
    } else {
      report += `❌ Not configured - run "Configure Components Data Sheet"\n`;
    }
    
    report += `\nNEXT STEPS:\n`;
    report += `1. Verify period values are not empty\n`;
    report += `2. Check components data sheet configuration\n`;
    report += `3. Try running formula generation again\n`;
    
    ui.alert('Debug Results', report, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Debug error:', error);
    ui.alert('Debug Error', `Error during debug: ${error.message}`, ui.ButtonSet.OK);
  }
}