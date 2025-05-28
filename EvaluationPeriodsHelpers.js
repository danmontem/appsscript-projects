/**
 * EvaluationPeriodsHelpers.gs
 * Helper functions for evaluation periods management
 */

/**
 * Helper function to get value or empty string (for most columns)
 */
function getValueOrEmpty(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value;
}

/**
 * Helper function to get value or 0 (specifically for id_indicador)
 */
function getValueOrZero(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  return value;
}

/**
 * Alternative: Manual upload option when external access fails
 */
function showManualUploadOption(periodsDataset, ui) {
  const instructions = `Alternative: Manual Data Entry\n\n` +
    `Since external document access failed, you can:\n\n` +
    `Option 1: Copy and paste data\n` +
    `• Open your external periods table\n` +
    `• Copy the relevant data\n` +
    `• Paste it into a new sheet in this document\n` +
    `• Run the mapping function again using the local sheet\n\n` +
    `Option 2: Upload CSV/Excel file\n` +
    `• Export your periods table as CSV or Excel\n` +
    `• Import it into this spreadsheet\n` +
    `• Use the local copy for mapping\n\n` +
    `Option 3: Manual entry\n` +
    `• Enter periods manually in the dataset\n\n` +
    `Would you like instructions for Option 1?`;
  
  const helpResult = ui.alert('Alternative Options', instructions, ui.ButtonSet.YES_NO);
  
  if (helpResult === ui.Button.YES) {
    createLocalPeriodsTemplate(periodsDataset, ui);
  }
}

/**
 * Create a template for local periods data
 */
function createLocalPeriodsTemplate(periodsDataset, ui) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create a template sheet for manual data entry
  const templateName = 'Periodos_Template';
  let templateSheet = ss.getSheetByName(templateName);
  
  if (templateSheet) {
    templateSheet.clear();
  } else {
    templateSheet = ss.insertSheet(templateName);
  }
  
  // Create headers for the template
  const templateHeaders = [
    'ID_Indicador',
    'Periodo_Base_EVA_Mayo',
    'Periodo_Base_EVA_Noviembre', 
    'Diagnostico_Mayo_25',
    'Evaluacion_1_Noviembre_25',
    'Evaluacion_2_Mayo_26',
    'Evaluacion_3_Noviembre_26',
    'Evaluacion_4_Mayo_27',
    'Evaluacion_5_Octubre_27'
  ];
  
  templateSheet.getRange(1, 1, 1, templateHeaders.length).setValues([templateHeaders]);
  
  // Format the template
  const headerRange = templateSheet.getRange(1, 1, 1, templateHeaders.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8f4fd');
  
  // Add some example data
  const exampleData = [
    [1, 'diciembre 2024', 'junio 2025', 'mayo 2025', 'noviembre 2025', 'mayo 2026', 'noviembre 2026', 'mayo 2027', 'octubre 2027'],
    [2, 'enero 2025', 'julio 2025', 'mayo 2025', 'noviembre 2025', 'mayo 2026', 'noviembre 2026', 'mayo 2027', 'octubre 2027'],
    ['...', '...', '...', '...', '...', '...', '...', '...', '...']
  ];
  
  templateSheet.getRange(2, 1, exampleData.length, templateHeaders.length).setValues(exampleData);
  
  // Auto-resize columns
  templateSheet.autoResizeColumns(1, templateHeaders.length);
  
  // Add instructions
  templateSheet.getRange('A1').setNote(
    'Instructions:\n\n' +
    '1. Replace example data with your actual periods data\n' +
    '2. Use formats like "mayo 2025" or "enero 2025 - diciembre 2025"\n' +
    '3. After filling, run "Map Periods from Local Template" function\n' +
    '4. Delete this template when done'
  );
  
  ui.alert(
    'Template Created',
    `Template sheet "${templateName}" created!\n\n` +
    `Next steps:\n` +
    `1. Copy your periods data from external source\n` +
    `2. Paste into the template (replace example data)\n` +
    `3. Run "Map Periods from Local Template" when ready\n\n` +
    `The template includes example formatting.`,
    ui.ButtonSet.OK
  );
}

/**
 * Normalize period text to standard format
 */
function normalizePeriodText(text) {
  const originalText = text.toLowerCase().trim();
  
  // Handle NA, N/A, and similar cases (but NOT empty strings)
  if (originalText === 'na' || originalText === 'n/a' || originalText === '-') {
    return 'NA';
  }
  
  // Handle empty/blank - return as is (preserve the distinction from NA)
  if (originalText === '') {
    return '';
  }
  
  // Handle specific patterns
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
      // Check if it's "month-month year" format
      const yearMatch = originalText.match(/202\d/);
      if (yearMatch) {
        const year = yearMatch[0];
        const firstMonth = parts[0].trim();
        const secondPart = parts[1].trim();
        
        // Extract second month (remove year if present)
        const secondMonth = secondPart.replace(/202\d/, '').trim();
        
        // Validate months
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
  
  // Already correct formats
  if (originalText.includes(' - ') || /^[a-z]+ 202\d$/.test(originalText)) {
    return originalText;
  }
  
  return `ERROR: Cannot normalize "${text}"`;
}

/**
 * Convert normalized period to formula date format
 */
function convertNormalizedPeriodToFormula(normalizedPeriod) {
  if (!normalizedPeriod || normalizedPeriod.includes('ERROR:')) {
    return normalizedPeriod;
  }
  
  const monthCodeMap = {
    'enero': 'ENE', 'febrero': 'FEB', 'marzo': 'MAR', 'abril': 'ABR',
    'mayo': 'MAY', 'junio': 'JUN', 'julio': 'JUL', 'agosto': 'AGO',
    'septiembre': 'SEP', 'octubre': 'OCT', 'noviembre': 'NOV', 'diciembre': 'DIC'
  };
  
  if (normalizedPeriod.includes(' - ')) {
    const parts = normalizedPeriod.split(' - ').map(p => p.trim());
    const startCode = convertSingleDateToCode(parts[0], monthCodeMap);
    const endCode = convertSingleDateToCode(parts[1], monthCodeMap);
    return `${startCode}-${endCode}`;
  } else {
    return convertSingleDateToCode(normalizedPeriod, monthCodeMap);
  }
}

/**
 * Convert single date to code format
 */
function convertSingleDateToCode(dateText, monthCodeMap) {
  const parts = dateText.trim().split(' ');
  if (parts.length === 2) {
    const month = parts[0].toLowerCase();
    const year = parts[1];
    const monthCode = monthCodeMap[month];
    if (monthCode && year.length === 4) {
      return `${monthCode}${year.substring(2)}`;
    }
  }
  return dateText;
}

/**
 * Generate indicator formula using cell references
 */
function generateIndicatorCellFormula(formulaText, currentRow, calculationColumn, sheet, rows, headers) {
  const componentCodes = extractComponentCodes(formulaText);
  
  if (componentCodes.length === 0) {
    return '';
  }
  
  const municipioIndex = headers.indexOf('municipio');
  const indicadorIdIndex = headers.indexOf('id_indicador');
  const tipoIndex = headers.indexOf('tipo de dato');
  const codigoIndex = headers.indexOf('codigo');
  
  const currentRowData = rows[currentRow - 2];
  const currentIndicatorId = currentRowData[indicadorIdIndex];
  const currentMunicipio = currentRowData[municipioIndex];
  
  const componentRefs = {};
  
  componentCodes.forEach(code => {
    const componentRowIndex = rows.findIndex(row => 
      row[municipioIndex] === currentMunicipio &&
      row[indicadorIdIndex] === currentIndicatorId &&
      row[tipoIndex] === 'Componente' &&
      row[codigoIndex] === code
    );
    
    if (componentRowIndex >= 0) {
      const componentRowNum = componentRowIndex + 2;
      componentRefs[code] = `${calculationColumn}${componentRowNum}`;
    }
  });
  
  let finalFormula = formulaText;
  
  Object.entries(componentRefs).forEach(([code, cellRef]) => {
    const regex = new RegExp(`\\b${code}\\b`, 'g');
    finalFormula = finalFormula.replace(regex, cellRef);
  });
  
  if (!finalFormula.startsWith('=')) {
    finalFormula = '=' + finalFormula;
  }
  
  return finalFormula;
}

/**
 * Extract component codes from formula
 */
function extractComponentCodes(formulaText) {
  const codePattern = /\b[A-Za-z]{2,}[A-Za-z0-9]*\b/g;
  const matches = formulaText.match(codePattern) || [];
  const excludeTerms = ['and', 'or', 'not', 'sum', 'avg', 'max', 'min', 'if', 'then', 'else'];
  
  return [...new Set(matches.filter(match => 
    !excludeTerms.includes(match.toLowerCase()) && 
    match.length >= 2
  ))];
}

/**
 * Get column letter from number
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    columnNumber--;
    letter = String.fromCharCode(65 + (columnNumber % 26)) + letter;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return letter;
}