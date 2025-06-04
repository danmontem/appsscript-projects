/**
 * DataCombiner.js
 * Cleaned and optimized functions for combining components and indicators datasets
 * Dependencies: CommonHelpers.js
 * Updated with formula-preserving refresh capability
 */

/**
 * Enhanced function to concatenate indicators and their components with user-selected sheets
 */
function concatAndOrderByIndicator() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get list of available sheets
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  // Ask user for components sheet name
  const componentsResult = ui.prompt(
    'Select Components Sheet',
    `Available sheets: ${sheetsText}\n\nEnter the name of the COMPONENTS sheet:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (componentsResult.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  const componentsSheetName = componentsResult.getResponseText().trim();
  
  // Ask user for indicators sheet name
  const indicatorsResult = ui.prompt(
    'Select Indicators Sheet',
    `Available sheets: ${sheetsText}\n\nEnter the name of the INDICATORS sheet:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (indicatorsResult.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  const indicatorsSheetName = indicatorsResult.getResponseText().trim();
  
  // Validate sheet names
  const hoja1 = ss.getSheetByName(componentsSheetName);
  const hoja2 = ss.getSheetByName(indicatorsSheetName);
  
  if (!hoja1) {
    ui.alert('Error', `Components sheet "${componentsSheetName}" not found.\n\nAvailable sheets: ${sheetsText}`, ui.ButtonSet.OK);
    return;
  }
  
  if (!hoja2) {
    ui.alert('Error', `Indicators sheet "${indicatorsSheetName}" not found.\n\nAvailable sheets: ${sheetsText}`, ui.ButtonSet.OK);
    return;
  }
  
  // Ask for destination sheet (current sheet is default)
  const currentSheetName = ss.getActiveSheet().getName();
  const destinationResult = ui.prompt(
    'Select Destination Sheet',
    `Current sheet: "${currentSheetName}"\n\nPress OK to use current sheet, or enter a different sheet name:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (destinationResult.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  let hojaDestino;
  const destinationSheetName = destinationResult.getResponseText().trim();
  
  if (destinationSheetName === "" || destinationSheetName === currentSheetName) {
    hojaDestino = ss.getActiveSheet();
  } else {
    hojaDestino = ss.getSheetByName(destinationSheetName);
    if (!hojaDestino) {
      ui.alert('Error', `Destination sheet "${destinationSheetName}" not found.\n\nAvailable sheets: ${sheetsText}`, ui.ButtonSet.OK);
      return;
    }
  }
  
  // Confirm before proceeding
  const confirmResult = ui.alert(
    'Confirm Selection',
    `Components: "${componentsSheetName}"\nIndicators: "${indicatorsSheetName}"\nDestination: "${hojaDestino.getName()}"\n\nThis will clear the destination sheet. Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    return; // User chose not to continue
  }
  
  // Progress indicator
  ui.alert('Processing...', 'Combining and sorting data. Please wait.', ui.ButtonSet.OK);

  // Use standardized column mapping from CommonHelpers
  const columnasDestino = Object.keys(STANDARD_COLUMN_MAPPING);

  try {
    // Get data from both sheets
    const datos1 = extraerDatosMapeados(hoja1, "hoja1", "Componente");
    const datos2 = extraerDatosMapeados(hoja2, "hoja2", "Indicador");

    const datosCombinados = datos1.concat(datos2);

    // Get column indices for sorting
    const indiceId = columnasDestino.indexOf("id_indicador");
    const indiceMunicipio = columnasDestino.indexOf("municipio");
    const indiceTipo = columnasDestino.indexOf("tipo de dato");

    // Enhanced sorting: municipio (alfabético) → id_indicator (numérico) → tipo (componentes primero)
    datosCombinados.sort((a, b) => {
      // 1. Sort by municipality (alphabetical)
      const compMun = a[indiceMunicipio].toString().localeCompare(b[indiceMunicipio].toString(), 'es');
      if (compMun !== 0) return compMun;
      
      // 2. Sort by indicator ID (numerical)
      const compId = getValueOrZero(a[indiceId]) - getValueOrZero(b[indiceId]);
      if (compId !== 0) return compId;
      
      // 3. Sort components before indicators within same municipality and ID
      if (a[indiceTipo] === "Componente" && b[indiceTipo] === "Indicador") return -1;
      if (a[indiceTipo] === "Indicador" && b[indiceTipo] === "Componente") return 1;
      
      // 4. If same type, maintain original order
      return 0;
    });

    // Clear destination sheet and write results
    hojaDestino.clear();
    
    // Store configuration for future refreshes
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('CONCAT_COMPONENTS_SHEET', componentsSheetName);
    scriptProperties.setProperty('CONCAT_INDICATORS_SHEET', indicatorsSheetName);
    scriptProperties.setProperty('CONCAT_DESTINATION_SHEET', hojaDestino.getName());
    scriptProperties.setProperty('CONCAT_ORIGINAL_COLUMN_ORDER', JSON.stringify(columnasDestino));
    
    // Set headers
    hojaDestino.getRange(1, 1, 1, columnasDestino.length).setValues([columnasDestino]);
    
    // Set data
    if (datosCombinados.length > 0) {
      hojaDestino.getRange(2, 1, datosCombinados.length, columnasDestino.length).setValues(datosCombinados);
    }
    
    // Apply basic formatting to headers
    const headerRange = hojaDestino.getRange(1, 1, 1, columnasDestino.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    // Auto-resize columns
    hojaDestino.autoResizeColumns(1, columnasDestino.length);
    
    // Success message with summary
    const totalComponentes = datos1.length;
    const totalIndicadores = datos2.length;
    const totalRows = datosCombinados.length;
    
    ui.alert(
      'Success!', 
      `Data combined successfully!\n\n` +
      `Source sheets:\n` +
      `• Components: "${componentsSheetName}" (${totalComponentes} rows)\n` +
      `• Indicators: "${indicatorsSheetName}" (${totalIndicadores} rows)\n` +
      `• Destination: "${hojaDestino.getName()}"\n\n` +
      `Total combined rows: ${totalRows}\n\n` +
      `Data sorted by:\n` +
      `1. Municipality (alphabetical)\n` +
      `2. Indicator ID (numerical)\n` +
      `3. Type (Components before Indicators)\n\n` +
      `You can now manually reorder columns if desired.\nThe refresh function will respect your custom order!`, 
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    // Enhanced error handling
    console.error('Error in concatAndOrderByIndicator:', error);
    ui.alert(
      'Error', 
      `An error occurred while processing the data:\n\n${error.message}\n\nPlease check:\n` +
      `1. Both sheets contain data\n` +
      `2. Column names match the expected format:\n` +
      `   Components: municipio, id_indicador, eje, tema, nombre_componente, codigo_componente\n` +
      `   Indicators: municipio, id_indicador, eje, tema, nombre_indicador, calculo_indicador\n` +
      `3. Data is properly formatted`, 
      ui.ButtonSet.OK
    );
  }
}

/**
 * Function to extract mapped data with enhanced error handling
 * Uses CommonHelpers functions for consistency
 */
function extraerDatosMapeados(hoja, fuente, etiquetaTipo) {
  try {
    const datos = hoja.getDataRange().getValues();
    const encabezado = datos[0];
    const cuerpo = datos.slice(1);

    // Get all column names in the correct order
    const allColumnNames = Object.keys(STANDARD_COLUMN_MAPPING);
    
    // Map each column to its source column index (except "tipo de dato" which is manual)
    const columnMappings = allColumnNames.map(col => {
      if (col === "tipo de dato") {
        return { name: col, index: -1, manual: true }; // Manual column
      } else {
        const nombreCol = STANDARD_COLUMN_MAPPING[col][fuente];
        const idx = encabezado.indexOf(nombreCol);
        if (idx === -1) {
          throw new Error(`Column "${nombreCol}" not found in sheet "${hoja.getName()}"`);
        }
        return { name: col, index: idx, manual: false };
      }
    });

    return cuerpo.map(fila => {
      // Build row data in the correct column order
      const datosFila = columnMappings.map(mapping => {
        if (mapping.manual) {
          // This is the "tipo de dato" column - set the label
          return etiquetaTipo;
        } else {
          const value = fila[mapping.index];
          
          // Special handling for id_indicador: convert empty/blank to 0 using CommonHelpers
          if (mapping.name === "id_indicador") {
            return getValueOrZero(value);
          }
          
          return getValueOrEmpty(value);
        }
      });
      
      return datosFila;
    });
  } catch (error) {
    throw new Error(`Error processing sheet "${hoja.getName()}": ${error.message}`);
  }
}

/**
 * ENHANCED: Smart refresh function that preserves formulas in calculation columns
 * Detects and preserves formula columns while refreshing core data
 */
function refreshConcatenatedDataset() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Get stored properties
  const componentsSheetName = scriptProperties.getProperty('CONCAT_COMPONENTS_SHEET');
  const indicatorsSheetName = scriptProperties.getProperty('CONCAT_INDICATORS_SHEET');
  const destinationSheetName = scriptProperties.getProperty('CONCAT_DESTINATION_SHEET');
  const originalColumnOrder = JSON.parse(scriptProperties.getProperty('CONCAT_ORIGINAL_COLUMN_ORDER') || '[]');
  
  if (!componentsSheetName || !indicatorsSheetName || !destinationSheetName) {
    ui.alert(
      'Error', 
      'No concatenated dataset configuration found.\n\nPlease use "Combine Components & Indicators" first to create the initial dataset.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Validate that all sheets still exist
  const componentsSheet = ss.getSheetByName(componentsSheetName);
  const indicatorsSheet = ss.getSheetByName(indicatorsSheetName);
  const destinationSheet = ss.getSheetByName(destinationSheetName);
  
  if (!componentsSheet || !indicatorsSheet || !destinationSheet) {
    ui.alert('Error', 'One or more source sheets not found. Please check sheet names.', ui.ButtonSet.OK);
    return;
  }
  
  // ENHANCED: Analyze current destination sheet structure
  const currentData = destinationSheet.getDataRange().getValues();
  const currentHeaders = currentData[0];
  const currentRows = currentData.slice(1);
  
  console.log('Analyzing current sheet structure...');
  console.log('Current headers:', currentHeaders);
  
  // Identify core columns vs calculation columns using CommonHelpers
  const coreColumns = Object.keys(STANDARD_COLUMN_MAPPING);
  const coreColumnIndices = [];
  const calculationColumns = [];
  
  currentHeaders.forEach((header, index) => {
    if (coreColumns.includes(header)) {
      coreColumnIndices.push({
        name: header,
        index: index,
        originalIndex: coreColumns.indexOf(header)
      });
    } else {
      // This is likely a calculation column - check if it contains formulas
      calculationColumns.push({
        name: header,
        index: index,
        hasFormulas: false // We'll check this below
      });
    }
  });
  
  // ENHANCED: Detect which columns contain formulas
  console.log('Detecting formula columns...');
  for (let calcCol of calculationColumns) {
    // Check first few non-empty cells to see if they contain formulas
    const columnRange = destinationSheet.getRange(2, calcCol.index + 1, Math.min(10, currentRows.length), 1);
    const formulas = columnRange.getFormulas();
    
    // ENHANCED FIX: Safe string check before using startsWith
    const hasAnyFormulas = formulas.some(row => 
      row[0] && typeof row[0] === 'string' && row[0].startsWith('='));
    calcCol.hasFormulas = hasAnyFormulas;
    
    console.log(`Column "${calcCol.name}": ${hasAnyFormulas ? 'HAS FORMULAS' : 'has values only'}`);
  }
  
  // Store formulas from calculation columns before refresh
  const preservedFormulas = {};
  calculationColumns.forEach(calcCol => {
    if (calcCol.hasFormulas) {
      console.log(`Preserving formulas from column: ${calcCol.name}`);
      const columnRange = destinationSheet.getRange(2, calcCol.index + 1, currentRows.length, 1);
      preservedFormulas[calcCol.name] = {
        formulas: columnRange.getFormulas(),
        index: calcCol.index
      };
    }
  });
  
  // Confirm before proceeding
  const formulaColumnsCount = Object.keys(preservedFormulas).length;
  const confirmMessage = `This will refresh the concatenated dataset:\n\n` +
    `• Source Components: "${componentsSheetName}"\n` +
    `• Source Indicators: "${indicatorsSheetName}"\n` +
    `• Destination: "${destinationSheetName}"\n\n` +
    `✅ Core data will be refreshed\n` +
    `✅ Custom column order will be preserved\n` +
    `✅ ${formulaColumnsCount} formula column(s) will be preserved\n\n` +
    `Continue?`;
    
  const confirmResult = ui.alert('Smart Refresh Dataset', confirmMessage, ui.ButtonSet.YES_NO);
  
  if (confirmResult !== ui.Button.YES) {
    return;
  }
  
  // Progress indicator
  ui.alert('Refreshing...', 'Updating dataset while preserving formulas. Please wait.', ui.ButtonSet.OK);
  
  try {
    // Get fresh data from both sheets
    const datos1 = extraerDatosMapeados(componentsSheet, "hoja1", "Componente");
    const datos2 = extraerDatosMapeados(indicatorsSheet, "hoja2", "Indicador");
    const datosCombinados = datos1.concat(datos2);

    // Sort data using correct column indices from STANDARD_COLUMN_MAPPING
    const columnOrder = Object.keys(STANDARD_COLUMN_MAPPING);
    const municipioIndex = columnOrder.indexOf("municipio");
    const idIndex = columnOrder.indexOf("id_indicador");
    const tipoIndex = columnOrder.indexOf("tipo de dato");
    
    datosCombinados.sort((a, b) => {
      const compMun = a[municipioIndex].toString().localeCompare(b[municipioIndex].toString(), 'es');
      if (compMun !== 0) return compMun;
      
      const compId = getValueOrZero(a[idIndex]) - getValueOrZero(b[idIndex]);
      if (compId !== 0) return compId;
      
      if (a[tipoIndex] === "Componente" && b[tipoIndex] === "Indicador") return -1;
      if (a[tipoIndex] === "Indicador" && b[tipoIndex] === "Componente") return 1;
      
      return 0;
    });

    // ENHANCED: Build final data preserving current column order and formulas
    const targetColumnOrder = [...currentHeaders]; // Use current order
    const finalData = [targetColumnOrder]; // Headers first
    
    // Create lookup for preserved formulas based on row key
    const formulaLookup = buildFormulaLookup(currentRows, currentHeaders, preservedFormulas);
    
    // Build final rows
    datosCombinados.forEach(newRowData => {
      const key = buildRowKey(newRowData, currentHeaders);
      
      // Build final row respecting target column order
      const finalRow = targetColumnOrder.map(colName => {
        if (coreColumns.includes(colName)) {
          // Use fresh data for core columns
          const coreIndex = coreColumns.indexOf(colName);
          const value = newRowData[coreIndex];
          
          // FIXED: Don't use || "" because 0 is falsy but valid
          // Apply special handling for id_indicador to ensure 0 is preserved
          if (colName === "id_indicador") {
            return getValueOrZero(value);
          } else {
            return getValueOrEmpty(value);
          }
        } else {
          // For calculation columns, preserve any existing data/formulas
          return ""; // Will be filled with formulas later
        }
      });
      
      finalData.push(finalRow);
    });

    // Clear and update destination sheet with core data
    destinationSheet.clear();
    
    // Write all data
    if (finalData.length > 0) {
      destinationSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
    }
    
    // ENHANCED: Restore formulas to calculation columns
    restoreFormulasToColumns(destinationSheet, preservedFormulas, formulaLookup, currentHeaders, finalData);
    
    // Apply formatting
    const headerRange = destinationSheet.getRange(1, 1, 1, targetColumnOrder.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    // Auto-resize columns
    destinationSheet.autoResizeColumns(1, targetColumnOrder.length);
    
    // Success message
    const totalComponentes = datos1.length;
    const totalIndicadores = datos2.length;
    const totalRows = datosCombinados.length;
    
    ui.alert(
      'Smart Refresh Complete!',
      `Dataset refreshed successfully with formula preservation!\n\n` +
      `Updated data:\n` +
      `• Components: "${componentsSheetName}" (${totalComponentes} rows)\n` +
      `• Indicators: "${indicatorsSheetName}" (${totalIndicadores} rows)\n` +
      `• Total rows: ${totalRows}\n\n` +
      `✅ Core data refreshed from source sheets\n` +
      `✅ Custom column order preserved\n` +
      `✅ ${formulaColumnsCount} calculation column(s) with formulas preserved\n\n` +
      `Your calculation formulas are intact and ready to use!`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error in smart refresh:', error);
    ui.alert(
      'Error', 
      `An error occurred during refresh:\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

// =============================================
// HELPER FUNCTIONS FOR FORMULA PRESERVATION
// =============================================

/**
 * Build lookup for preserved formulas based on row key
 */
function buildFormulaLookup(currentRows, currentHeaders, preservedFormulas) {
  const formulaLookup = {};
  
  currentRows.forEach((row, index) => {
    const key = buildRowKey(row, currentHeaders);
    formulaLookup[key] = {};
    
    // Store formulas for this row
    Object.entries(preservedFormulas).forEach(([colName, colInfo]) => {
      if (colInfo.formulas[index] && colInfo.formulas[index][0]) {
        formulaLookup[key][colName] = colInfo.formulas[index][0];
      }
    });
  });
  
  return formulaLookup;
}

/**
 * Build row key for formula lookup
 */
function buildRowKey(row, headers) {
  // Use proper column indices from headers or fallback to STANDARD_COLUMN_MAPPING order
  const columnOrder = Object.keys(STANDARD_COLUMN_MAPPING);
  
  const municipioIndex = headers.indexOf("municipio") !== -1 ? headers.indexOf("municipio") : columnOrder.indexOf("municipio");
  const idIndex = headers.indexOf("id_indicador") !== -1 ? headers.indexOf("id_indicador") : columnOrder.indexOf("id_indicador");
  const tipoIndex = headers.indexOf("tipo de dato") !== -1 ? headers.indexOf("tipo de dato") : columnOrder.indexOf("tipo de dato");
  const codigoIndex = headers.indexOf("codigo") !== -1 ? headers.indexOf("codigo") : columnOrder.indexOf("codigo");
  
  const municipio = row[municipioIndex];
  const indicadorId = getValueOrZero(row[idIndex]);
  const tipo = row[tipoIndex];
  const codigo = row[codigoIndex];
  
  return `${municipio}_${indicadorId}_${tipo}_${codigo}`;
}

/**
 * Restore formulas to calculation columns
 */
function restoreFormulasToColumns(destinationSheet, preservedFormulas, formulaLookup, currentHeaders, finalData) {
  console.log('Restoring formulas to calculation columns...');
  
  const updatedRows = finalData.slice(1); // Skip headers
  
  Object.entries(preservedFormulas).forEach(([colName, colInfo]) => {
    console.log(`Restoring formulas to column: ${colName}`);
    
    const colIndex = currentHeaders.indexOf(colName);
    if (colIndex === -1) return;
    
    // Build array of formulas to restore
    const formulasToRestore = [];
    
    updatedRows.forEach((row, rowIndex) => {
      const key = buildRowKey(row, currentHeaders);
      
      if (formulaLookup[key] && formulaLookup[key][colName]) {
        formulasToRestore.push([formulaLookup[key][colName]]);
      } else {
        formulasToRestore.push(['']); // Empty if no formula found
      }
    });
    
    // Write formulas back
    if (formulasToRestore.length > 0) {
      const range = destinationSheet.getRange(2, colIndex + 1, formulasToRestore.length, 1);
      range.setFormulas(formulasToRestore);
    }
  });
}

/**
 * Enhanced preview function that also asks for sheet names
 */
function previewColumnMapping() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get list of available sheets
  const availableSheets = ss.getSheets().map(sheet => sheet.getName());
  const sheetsText = availableSheets.join(', ');
  
  // Ask user for components sheet name
  const componentsResult = ui.prompt(
    'Select Components Sheet for Preview',
    `Available sheets: ${sheetsText}\n\nEnter the name of the COMPONENTS sheet:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (componentsResult.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  const componentsSheetName = componentsResult.getResponseText().trim();
  
  // Ask user for indicators sheet name
  const indicatorsResult = ui.prompt(
    'Select Indicators Sheet for Preview',
    `Available sheets: ${sheetsText}\n\nEnter the name of the INDICATORS sheet:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (indicatorsResult.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }
  
  const indicatorsSheetName = indicatorsResult.getResponseText().trim();

  const hoja1 = ss.getSheetByName(componentsSheetName);
  const hoja2 = ss.getSheetByName(indicatorsSheetName);

  if (!hoja1) {
    ui.alert('Error', `Components sheet "${componentsSheetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  
  if (!hoja2) {
    ui.alert('Error', `Indicators sheet "${indicatorsSheetName}" not found.`, ui.ButtonSet.OK);
    return;
  }

  const encabezado1 = hoja1.getDataRange().getValues()[0];
  const encabezado2 = hoja2.getDataRange().getValues()[0];

  let preview = `Column Mapping Preview:\n\n`;
  preview += `Components Sheet: "${componentsSheetName}"\n`;
  preview += `Indicators Sheet: "${indicatorsSheetName}"\n\n`;
  
  Object.keys(STANDARD_COLUMN_MAPPING).forEach(finalCol => {
    const map = STANDARD_COLUMN_MAPPING[finalCol];
    preview += `"${finalCol}":\n`;
    
    if (map.hoja1 && map.hoja1 !== "Manual") {
      const found1 = encabezado1.includes(map.hoja1);
      preview += `  ${componentsSheetName}: "${map.hoja1}" ${found1 ? '✓' : '✗'}\n`;
    }
    
    if (map.hoja2 && map.hoja2 !== "Manual") {
      const found2 = encabezado2.includes(map.hoja2);
      preview += `  ${indicatorsSheetName}: "${map.hoja2}" ${found2 ? '✓' : '✗'}\n`;
    }
    
    if (map.hoja1 === "Manual" || map.hoja2 === "Manual") {
      preview += `  Generated automatically\n`;
    }
    
    preview += `\n`;
  });

  ui.alert('Column Mapping Preview', preview, ui.ButtonSet.OK);
}