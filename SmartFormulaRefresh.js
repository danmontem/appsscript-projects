/**
 * SmartFormulaRefresh.js
 * Intelligent formula refresh system for period changes
 * Dependencies: CommonHelpers.js, PeriodsManager.js (for some helper functions)
 * 
 * SOLVES: Period data changes (SINGLE ↔ SUM) not updating calculation formulas
 * PRESERVES: Manual formula entries while updating auto-generated ones
 */

/**
 * Main function: Smart refresh calculation formulas when periods change
 * Updates only formulas that need changes, preserves manual entries
 */
function smartRefreshCalculationFormulas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('=== SMART FORMULA REFRESH START ===');
  
  // Ask user for required sheets
  const sheets = getRequiredSheets(ui, ss);
  if (!sheets) return; // User cancelled
  
  // Ask which evaluation period to refresh
  const selectedPeriod = selectEvaluationPeriod(ui);
  if (!selectedPeriod) return; // User cancelled
  
  // Confirm before proceeding
  if (!confirmRefresh(ui, sheets, selectedPeriod)) return;
  
  // Perform the smart refresh
  performSmartFormulaRefresh(sheets, selectedPeriod, ui);
}

/**
 * Get required sheets from user
 */
function getRequiredSheets(ui, ss) {
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  // Get periods dataset (source of truth)
  const periodsResult = ui.prompt(
    'Select Periods Dataset (SOURCE)',
    `Available: ${sheetsText}\n\nEnter periods dataset name:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (periodsResult.getSelectedButton() !== ui.Button.OK) return null;
  
  const periodsName = periodsResult.getResponseText().trim();
  const periodsSheet = ss.getSheetByName(periodsName);
  if (!periodsSheet) {
    ui.alert('Error', `Sheet "${periodsName}" not found.`, ui.ButtonSet.OK);
    return null;
  }
  
  // Get calculations sheet (target to update)
  const calcResult = ui.prompt(
    'Select Calculations Sheet (TARGET)',
    `Available: ${sheetsText}\n\nEnter calculations sheet name:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (calcResult.getSelectedButton() !== ui.Button.OK) return null;
  
  const calcName = calcResult.getResponseText().trim();
  const calcSheet = ss.getSheetByName(calcName);
  if (!calcSheet) {
    ui.alert('Error', `Sheet "${calcName}" not found.`, ui.ButtonSet.OK);
    return null;
  }
  
  // Get components sheet (for references)
  const compResult = ui.prompt(
    'Select Components Dataset',
    `Available: ${sheetsText}\n\nEnter components dataset name:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (compResult.getSelectedButton() !== ui.Button.OK) return null;
  
  const compName = compResult.getResponseText().trim();
  const compSheet = ss.getSheetByName(compName);
  if (!compSheet) {
    ui.alert('Error', `Sheet "${compName}" not found.`, ui.ButtonSet.OK);
    return null;
  }
  
  return {
    periods: periodsSheet,
    periodsName: periodsName,
    calculations: calcSheet,
    calculationsName: calcName,
    components: compSheet,
    componentsName: compName
  };
}

/**
 * Select evaluation period to refresh
 */
function selectEvaluationPeriod(ui) {
  const periodResult = ui.prompt(
    'Select Evaluation Period',
    `Available periods:\n${EVALUATION_PERIOD_COLUMNS.join('\n')}\n\nEnter period name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (periodResult.getSelectedButton() !== ui.Button.OK) return null;
  return periodResult.getResponseText().trim();
}

/**
 * Confirm before proceeding
 */
function confirmRefresh(ui, sheets, selectedPeriod) {
  const confirmResult = ui.alert(
    'Confirm Smart Formula Refresh',
    `Smart refresh for:\n\n` +
    `• Periods: "${sheets.periodsName}"\n` +
    `• Calculations: "${sheets.calculationsName}"\n` +
    `• Components: "${sheets.componentsName}"\n` +
    `• Period: "${selectedPeriod}"\n\n` +
    `✅ Updates formulas when period types change (SINGLE ↔ SUM)\n` +
    `✅ Preserves manual formula entries\n` +
    `✅ Updates date ranges when periods change\n\n` +
    `Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  return confirmResult === ui.Button.YES;
}

/**
 * Core smart refresh logic
 */
function performSmartFormulaRefresh(sheets, selectedPeriod, ui) {
  const startTime = new Date();
  
  try {
    // Get data from sheets
    const periodsData = sheets.periods.getDataRange().getValues();
    const periodsHeaders = periodsData[0];
    
    const calcData = sheets.calculations.getDataRange().getValues();
    const calcHeaders = calcData[0];
    const calcRows = calcData.slice(1);
    
    // FIXED: Get formulas instead of values for the formula column
    const calcFormulas = sheets.calculations.getDataRange().getFormulas();
    const calcFormulaRows = calcFormulas.slice(1);
    
    // Find formula column
    const formulaColumnName = `Calculo_${selectedPeriod}`;
    const formulaColumnIndex = calcHeaders.indexOf(formulaColumnName);
    
    if (formulaColumnIndex === -1) {
      ui.alert('Error', `Formula column "${formulaColumnName}" not found.`, ui.ButtonSet.OK);
      return;
    }
    
    console.log(`Found formula column: ${formulaColumnName}`);
    
    // Build current periods lookup
    const periodsLookup = buildCurrentPeriodsLookup(periodsData, periodsHeaders, selectedPeriod);
    console.log(`Built periods lookup: ${periodsLookup.size} entries`);
    
    // FIXED: Analyze formulas using the formula data, not the calculated values
    const analysis = analyzeFormulasForUpdates(calcRows, calcHeaders, formulaColumnIndex, periodsLookup, calcFormulaRows);
    
    console.log(`Analysis: ${analysis.total} total, ${analysis.needsUpdate.length} need updates, ${analysis.manual} manual`);
    
    if (analysis.needsUpdate.length === 0) {
      ui.alert('No Updates Needed', `All ${analysis.total} formulas are up to date!`, ui.ButtonSet.OK);
      return;
    }
    
    // Update the formulas that need changes
    const updateResults = updateFormulasSelectively(sheets, analysis.needsUpdate, periodsLookup, formulaColumnIndex);
    
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    
    // Success message
    ui.alert(
      'Smart Refresh Complete!',
      `✅ Successfully updated formulas!\n\n` +
      `• Total formulas: ${analysis.total}\n` +
      `• Manual entries preserved: ${analysis.manual}\n` +
      `• Formulas updated: ${updateResults.updated}\n` +
      `• Period type changes: ${updateResults.typeChanges}\n` +
      `• Date range changes: ${updateResults.dateChanges}\n\n` +
      `⏱️ Time: ${executionTime.toFixed(1)}s`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Smart refresh error:', error);
    ui.alert('Error', `Smart refresh failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Build lookup of current periods
 */
function buildCurrentPeriodsLookup(periodsData, periodsHeaders, selectedPeriod) {
  const lookup = new Map();
  const rows = periodsData.slice(1);
  
  const municipioIndex = periodsHeaders.indexOf('municipio');
  const idIndex = periodsHeaders.indexOf('id_indicador');
  const tipoIndex = periodsHeaders.indexOf('tipo de dato');
  const codigoIndex = periodsHeaders.indexOf('codigo');
  const periodIndex = periodsHeaders.indexOf(selectedPeriod);
  
  if (periodIndex === -1) {
    throw new Error(`Period "${selectedPeriod}" not found in periods dataset`);
  }
  
  rows.forEach(row => {
    const municipio = row[municipioIndex];
    const id = row[idIndex];
    const tipo = row[tipoIndex];
    const codigo = row[codigoIndex];
    const period = row[periodIndex];
    
    if (municipio && codigo && period) {
      const key = `${municipio}_${id}_${tipo}_${codigo}`;
      lookup.set(key, period);
    }
  });
  
  return lookup;
}

/**
 * Analyze which formulas need updates
 */
function analyzeFormulasForUpdates(calcRows, calcHeaders, formulaColumnIndex, periodsLookup, calcFormulaRows) {
  const analysis = {
    total: 0,
    manual: 0,
    needsUpdate: []
  };
  
  const municipioIndex = calcHeaders.indexOf('municipio');
  const idIndex = calcHeaders.indexOf('id_indicador');
  const tipoIndex = calcHeaders.indexOf('tipo de dato');
  const codigoIndex = calcHeaders.indexOf('codigo');
  
  console.log(`Starting analysis with formula column at index: ${formulaColumnIndex}`);
  
  calcRows.forEach((row, rowIndex) => {
    const actualRowNum = rowIndex + 2;
    const municipio = row[municipioIndex];
    const id = row[idIndex];
    const tipo = row[tipoIndex];
    const codigo = row[codigoIndex];
    
    // FIXED: Get the actual formula from the formula array, not the calculated value
    const currentFormula = calcFormulaRows[rowIndex] ? calcFormulaRows[rowIndex][formulaColumnIndex] : '';
    
    // ENHANCED DEBUG: Log first few rows with safer logging and actual formula inspection
    if (rowIndex < 3) {
      const safeValue = currentFormula === null ? 'null' : 
                       currentFormula === undefined ? 'undefined' : 
                       typeof currentFormula === 'string' ? `"${currentFormula}"` : 
                       String(currentFormula);
      console.log(`Row ${actualRowNum}: formula type = ${typeof currentFormula}, formula = ${safeValue}`);
      
      // Also show what the old method was returning
      const oldValue = row[formulaColumnIndex];
      console.log(`Row ${actualRowNum}: old method returned = ${typeof oldValue}, value = ${oldValue}`);
    }
    
    // Only check components
    if (!codigo || tipo !== 'Componente') return;
    
    analysis.total++;
    
    // Get current period for this component
    const lookupKey = `${municipio}_${id}_${tipo}_${codigo}`;
    const currentPeriod = periodsLookup.get(lookupKey);
    
    if (!currentPeriod) {
      console.log(`No period found for: ${lookupKey}`);
      return;
    }
    
    // Check if formula needs update
    const updateInfo = checkIfFormulaNeedsUpdate(currentFormula, currentPeriod, codigo);
    
    if (updateInfo.isManual) {
      analysis.manual++;
      // DEBUG: Log manual entries for first few
      if (analysis.manual <= 3) {
        const safeValue = currentFormula === null ? 'null' : 
                         currentFormula === undefined ? 'undefined' : 
                         typeof currentFormula === 'string' ? `"${currentFormula}"` : 
                         String(currentFormula);
        console.log(`Manual entry detected at row ${actualRowNum}: ${codigo}, formula type: ${typeof currentFormula}, formula: ${safeValue}`);
      }
    } else if (updateInfo.needsUpdate) {
      analysis.needsUpdate.push({
        rowIndex: rowIndex,
        actualRowNum: actualRowNum,
        municipio: municipio,
        codigo: codigo,
        currentPeriod: currentPeriod,
        changeType: updateInfo.changeType,
        lookupKey: lookupKey
      });
    }
  });
  
  return analysis;
}

/**
 * Check if a single formula needs updating
 */
function checkIfFormulaNeedsUpdate(currentFormula, currentPeriod, codigo) {
  const result = {
    isManual: false,
    needsUpdate: false,
    changeType: null
  };
  
  // ENHANCED FIX: Handle all data types safely
  // Google Sheets can return strings, numbers, booleans, null, undefined, or Date objects
  if (currentFormula === null || 
      currentFormula === undefined || 
      currentFormula === '' || 
      typeof currentFormula === 'number' || 
      typeof currentFormula === 'boolean' || 
      currentFormula instanceof Date) {
    result.isManual = true;
    return result;
  }
  
  // Convert to string safely and trim
  let formulaString;
  try {
    formulaString = String(currentFormula).trim();
  } catch (error) {
    console.log(`Error converting formula to string for ${codigo}: ${error.message}`);
    result.isManual = true;
    return result;
  }
  
  // Check if it's a manual entry or not a valid formula
  if (!formulaString || 
      !formulaString.startsWith('=CALCULATE_INDICATOR') || 
      !formulaString.includes(codigo)) {
    result.isManual = true;
    return result;
  }
  
  try {
    // Convert period to expected format
    const expectedDateRange = convertNormalizedPeriodToFormula(currentPeriod);
    const expectedType = expectedDateRange.includes('-') ? 'SUM' : 'SINGLE';
    
    // Extract current formula info
    const currentType = formulaString.includes('"SINGLE(') ? 'SINGLE' : 'SUM';
    const dateMatch = formulaString.match(new RegExp(`${codigo}:([^"]+)`));
    const currentDateRange = dateMatch ? dateMatch[1] : null;
    
    // Check for changes
    if (currentType !== expectedType) {
      result.needsUpdate = true;
      result.changeType = 'TYPE_CHANGE';
      console.log(`Type change: ${codigo} ${currentType} → ${expectedType}`);
    } else if (currentDateRange && currentDateRange !== expectedDateRange) {
      result.needsUpdate = true;
      result.changeType = 'DATE_CHANGE';
      console.log(`Date change: ${codigo} ${currentDateRange} → ${expectedDateRange}`);
    }
    
  } catch (error) {
    console.log(`Error analyzing ${codigo}: ${error.message}`);
    result.isManual = true; // Treat errors as manual to preserve them
  }
  
  return result;
}

/**
 * Update only the formulas that need changes
 */
function updateFormulasSelectively(sheets, needsUpdate, periodsLookup, formulaColumnIndex) {
  const results = {
    updated: 0,
    typeChanges: 0,
    dateChanges: 0,
    errors: 0
  };
  
  // Build component row lookup
  const compData = sheets.components.getDataRange().getValues();
  const compHeaders = compData[0];
  const componentRowLookup = new Map();
  
  const municipioIndex = compHeaders.indexOf('municipio');
  const codigoIndex = compHeaders.findIndex(h => h.includes('codigo'));
  
  if (municipioIndex !== -1 && codigoIndex !== -1) {
    compData.slice(1).forEach((row, index) => {
      const municipio = row[municipioIndex];
      const codigo = row[codigoIndex];
      if (municipio && codigo) {
        componentRowLookup.set(`${municipio}_${codigo}`, index + 2);
      }
    });
  }
  
  console.log(`Component lookup: ${componentRowLookup.size} entries`);
  
  // Update each formula
  needsUpdate.forEach(updateInfo => {
    try {
      const { codigo, municipio, currentPeriod, actualRowNum, changeType } = updateInfo;
      
      // Generate new formula
      const dateRange = convertNormalizedPeriodToFormula(currentPeriod);
      const municipioCell = `A${actualRowNum}`;
      
      const componentRowNum = componentRowLookup.get(`${municipio}_${codigo}`);
      if (!componentRowNum) {
        console.log(`Component not found: ${municipio}_${codigo}`);
        results.errors++;
        return;
      }
      
      const componentRef = `${sheets.componentsName}!${componentRowNum}:${componentRowNum}`;
      
      let newFormula;
      if (dateRange.includes('-')) {
        newFormula = `=CALCULATE_INDICATOR("SUM(${codigo}:${dateRange})", ${municipioCell}, ${componentRef})`;
      } else {
        newFormula = `=CALCULATE_INDICATOR("SINGLE(${codigo}:${dateRange})", ${municipioCell}, ${componentRef})`;
      }
      
      // Update the cell
      sheets.calculations.getRange(actualRowNum, formulaColumnIndex + 1).setFormula(newFormula);
      results.updated++;
      
      if (changeType === 'TYPE_CHANGE') results.typeChanges++;
      if (changeType === 'DATE_CHANGE') results.dateChanges++;
      
      console.log(`Updated ${codigo}: ${newFormula}`);
      
    } catch (error) {
      console.log(`Error updating ${updateInfo.codigo}: ${error.message}`);
      results.errors++;
    }
  });
  
  return results;
}
