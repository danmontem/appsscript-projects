
/**
 * Configure smart refresh settings
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

// =============================================
// SMART REFRESH ALL PERIODS DATA
// =============================================

