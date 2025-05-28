/**
 * @NotOnlyCurrentDoc
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Municipality Tools')
    .addItem('Test Document Access', 'testDocumentAccess')
    .addSeparator()
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
    .addSeparator()
    .addItem('Create Formula Builder', 'createFormulaBuilder')
    .addSeparator()
    .addItem('Setup Auto-Refresh (Every 5 Minutes)', 'createDataRefreshTrigger')
    .addItem('Refresh Calculations Now', 'manualRefreshCalculations')
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Evaluation Periods')
      .addItem('Create Evaluation Periods Dataset', 'createEvaluationPeriodsDataset')
      .addSeparator()
      .addItem('⚡ Refresh All Periods Data (Smart)', 'refreshAllPeriodsData')
      .addItem('⚙️ Configure Smart Refresh', 'configureSmartRefresh')
      .addSeparator()
      .addItem('Map Periods (Flexible - Choose Sheets)', 'mapPeriodsFromExternalTableFlexible')
      .addItem('Map Periods (Quick - Preset Names)', 'mapPeriodsFromExternalTableUltraFast')
      .addItem('Map Periods from Local Template', 'mapPeriodsFromLocalTemplate')
      .addSeparator()
      .addItem('Generate Calculations from Periods Dataset', 'generateCalculationsFromPeriodsDataset')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Debug Tools')
      .addItem('Test External Document Access', 'testExternalDocumentAccess')
      .addItem('Test Mapping (First 200 Rows)', 'testMappingFirst200Rows')
    )
    .addToUi();
}

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
    
    // Find ID column
    const idColumnIndex = headers.findIndex(h => 
      h.toLowerCase().includes('id') && h.toLowerCase().includes('indicador')
    );
    
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