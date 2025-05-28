/**
 * UI Helper functions for testing document access and getting sheet information
 */

function testDocumentAccess() {
  const ui = SpreadsheetApp.getUi();
  
  // Get document ID from user
  const result = ui.prompt(
    "Test Document Access",
    "Enter the document ID you want to test:",
    ui.ButtonSet.OK_CANCEL
  );
  
  // Check if user clicked "OK"
  if (result.getSelectedButton() == ui.Button.OK) {
    const docId = result.getResponseText().trim();
    
    try {
      // Try to open the document
      const doc = SpreadsheetApp.openById(docId);
      const name = doc.getName();
      const sheets = doc.getSheets().map(sheet => sheet.getName()).join(", ");
      
      // Success - show document info
      ui.alert(
        "Success!",
        `Document accessed successfully!\n\nDocument name: ${name}\nSheets: ${sheets}\n\nYour account: ${Session.getEffectiveUser().getEmail()}`,
        ui.ButtonSet.OK
      );
    } catch (e) {
      // Failed - show error
      ui.alert(
        "Error",
        `Could not access document: ${e.message}\n\nPlease check:\n1. Document ID is correct\n2. Document is shared with your account: ${Session.getEffectiveUser().getEmail()}\n3. You've authorized the script to access external documents`,
        ui.ButtonSet.OK
      );
    }
  }
}

function getSourceDocumentSheets(sourceDocumentId) {
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceDocumentId);
    const sheets = sourceSpreadsheet.getSheets();
    return {
      success: true,
      sheets: sheets.map(sheet => sheet.getName())
    };
  } catch (e) {
    return {
      success: false,
      message: "Cannot access the source document. Please check the Document ID and ensure it's shared with you."
    };
  }
}

/**
 * Helper function that creates a formula builder for indicator calculations
 */
function createFormulaBuilder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let builderSheet = ss.getSheetByName("Formula Builder");
  
  if (!builderSheet) {
    builderSheet = ss.insertSheet("Formula Builder");
  } else {
    builderSheet.clear();
  }
  
  // Try to get components from BD_componentes sheet
  const componentsSheet = ss.getSheetByName("BD_componentes");
  
  // Extract unique components and municipalities
  const components = new Map();
  const municipalities = new Set();
  
  if (componentsSheet) {
    const data = componentsSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const municipio = row[0]; // Column A is municipio
      const componentCode = row[3]; // Column D is codigo_componente
      const componentName = row[7]; // Column H is nombre_componente
      
      municipalities.add(municipio);
      
      if (componentCode && !components.has(componentCode)) {
        components.set(componentCode, componentName);
      }
    }
  } else {
    // If BD_componentes sheet doesn't exist, use active sheet
    const datasetSheet = ss.getActiveSheet();
    const data = datasetSheet.getDataRange().getValues();
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const municipio = row[0]; // Column A is municipio
      const componentCode = row[3]; // Column D is codigo_componente
      const componentName = row[7]; // Column H is nombre_componente
      
      municipalities.add(municipio);
      
      if (componentCode && !components.has(componentCode)) {
        components.set(componentCode, componentName);
      }
    }
  }
  
  // Create builder sections
  builderSheet.getRange(1, 1).setValue("INDICATOR FORMULA BUILDER");
  builderSheet.getRange(1, 1).setFontWeight("bold").setFontSize(14);
  
  // Operation section
  builderSheet.getRange(3, 1).setValue("1. Basic Operations:");
  builderSheet.getRange(4, 1, 1, 5).setValues([["SUM", "AVG", "SINGLE", "MAX", "MIN"]]);
  
  // Component section
  builderSheet.getRange(6, 1).setValue("2. Component Codes:");
  let rowIndex = 7;
  for (const [code, name] of components.entries()) {
    builderSheet.getRange(rowIndex, 1, 1, 2).setValues([[code, name]]);
    rowIndex++;
  }
  
  // Date range section
  builderSheet.getRange(6, 4).setValue("3. Date Range Shortcuts:");
  builderSheet.getRange(7, 4, 6, 2).setValues([
    ["OCT24-MAR25", "Oct 2024 - Mar 2025"],
    ["ABR25-SEP25", "Apr 2025 - Sep 2025"],
    ["OCT25-MAR26", "Oct 2025 - Mar 2026"],
    ["ABR26-SEP26", "Apr 2026 - Sep 2026"],
    ["MAR25", "Single month: Mar 2025"],
    ["SEP25", "Single month: Sep 2025"]
  ]);
  
  // Municipality section
  builderSheet.getRange(15, 4).setValue("4. Municipalities:");
  rowIndex = 16;
  for (const municipio of municipalities) {
    builderSheet.getRange(rowIndex, 4).setValue(municipio);
    rowIndex++;
  }
  
  // Complex formula section
  const lastRow = Math.max(rowIndex, 30);
  builderSheet.getRange(lastRow, 1).setValue("COMPLEX FORMULA EXAMPLES:");
  builderSheet.getRange(lastRow, 1).setFontWeight("bold");
  
  // Example formulas for complex calculations
  builderSheet.getRange(lastRow + 2, 1).setValue("Basic formulas:");
  builderSheet.getRange(lastRow + 3, 1).setValue("Single component, single month:");
  builderSheet.getRange(lastRow + 3, 2).setValue('=CALCULATE_INDICATOR("SINGLE(COMP-1:MAR25)", "Monterrey")');
  
  builderSheet.getRange(lastRow + 4, 1).setValue("Sum of components over a period:");
  builderSheet.getRange(lastRow + 4, 2).setValue('=CALCULATE_INDICATOR("SUM(COMP-1:OCT24-MAR25, COMP-2:OCT24-MAR25)", "Monterrey")');
  
  builderSheet.getRange(lastRow + 6, 1).setValue("Complex formulas with arithmetic:");
  
  builderSheet.getRange(lastRow + 7, 1).setValue("Rate per 100,000 population:");
  builderSheet.getRange(lastRow + 7, 2).setValue('=CALCULATE_INDICATOR("SUM(HTA:OCT24-MAR25)/POP:MAR25*100000", "Monterrey")');
  
  builderSheet.getRange(lastRow + 8, 1).setValue("Percentage change:");
  builderSheet.getRange(lastRow + 8, 2).setValue('=CALCULATE_INDICATOR("(COMP-1:MAR25-COMP-1:MAR24)/COMP-1:MAR24*100", "Monterrey")');
  
  builderSheet.getRange(lastRow + 9, 1).setValue("Complex rate with multiple components:");
  builderSheet.getRange(lastRow + 9, 2).setValue('=CALCULATE_INDICATOR("SUM(HTA:OCT24-MAR25,HTP:OCT24-MAR25)/(POP:MAR25/100000)", "Monterrey")');
  
  builderSheet.getRange(lastRow + 10, 1).setValue("Average with scaling factor:");
  builderSheet.getRange(lastRow + 10, 2).setValue('=CALCULATE_INDICATOR("AVG(COMP-1:OCT24-MAR25,COMP-2:OCT24-MAR25)*10", "Monterrey")');
  
  // Format sheet
  builderSheet.autoResizeColumns(1, 5);
  
  SpreadsheetApp.getUi().alert("Formula Builder created! Use it to help construct your indicator formulas.");
}