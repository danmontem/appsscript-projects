/**
 * CommonHelpers.js
 * Consolidated utility functions used across the entire project
 * This file eliminates duplicates and provides a single source of truth
 */

// =============================================
// DATA VALIDATION HELPERS
// =============================================

/**
 * Helper function to get value or empty string (for most columns)
 * Replaces multiple duplicate versions across files
 */
function getValueOrEmpty(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value;
}

/**
 * Helper function to get value or 0 (specifically for id_indicador)
 * Replaces multiple duplicate versions across files
 */
function getValueOrZero(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  return value;
}

/**
 * Normalize indicator ID for consistent comparison
 * Consolidated from multiple files (EvaluationPeriodsMapping, EvaluationPeriodsManager, etc.)
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

// =============================================
// PERIOD TEXT NORMALIZATION
// =============================================

/**
 * Normalize period text to standard format
 * Consolidated from EvaluationPeriodsHelpers.js and EvaluationPeriodsSmartRefresh.js
 */
function normalizePeriodText(text) {
  if (!text || text === '') return '';
  
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
 * From EvaluationPeriodsHelpers.js
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
 * Helper for convertNormalizedPeriodToFormula
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

// =============================================
// MONTH/DATE CONVERSION UTILITIES
// =============================================

/**
 * Helper function to convert month codes to month names
 * Consolidated from FormulaCalculator.js
 */
function convertMonthCode(code) {
  // If the input is already a full month name (contains a space), return it as is
  if (code.includes(" ")) {
    return code.toLowerCase(); // Ensure it's lowercase for consistency
  }

  // Otherwise, handle the code format (OCT24 or OCT2024)
  let monthPart, yearPart;
  
  if (code.length === 5) {
    // Format is OCT24
    monthPart = code.substring(0, 3).toUpperCase();
    yearPart = code.substring(3);
  } else if (code.length === 7) {
    // Format is OCT2024
    monthPart = code.substring(0, 3).toUpperCase();
    yearPart = code.substring(3);
  } else {
    throw new Error(`Invalid month code format: ${code}. Use format OCT24 or OCT2024`);
  }
  
  // Map month codes to names
  const monthMap = {
    "ENE": "enero",
    "FEB": "febrero",
    "MAR": "marzo",
    "ABR": "abril",
    "MAY": "mayo",
    "JUN": "junio",
    "JUL": "julio",
    "AGO": "agosto",
    "SEP": "septiembre", 
    "OCT": "octubre",
    "NOV": "noviembre",
    "DIC": "diciembre"
  };
  
  const monthName = monthMap[monthPart];
  if (!monthName) {
    throw new Error(`Invalid month code: ${monthPart}`);
  }
  
  // Handle both 2-digit and 4-digit year formats
  if (yearPart.length === 2) {
    return `${monthName} 20${yearPart}`;
  } else {
    return `${monthName} ${yearPart}`;
  }
}

// =============================================
// COLUMN FINDING UTILITIES
// =============================================

/**
 * Find ID column with flexible matching
 * Consolidated from multiple evaluation files
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
 * Consolidated from EvaluationPeriodsMapping.js
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

// =============================================
// SPREADSHEET UTILITIES
// =============================================

/**
 * Get column letter from number
 * From EvaluationPeriodsHelpers.js
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

/**
 * Get column indices for required columns
 * From EvaluationPeriodsMapping.js
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

// =============================================
// FORMULA PROCESSING UTILITIES
// =============================================

/**
 * Extract component codes from formula
 * From EvaluationPeriodsHelpers.js
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

// =============================================
// CONFIGURATION CONSTANTS
// =============================================

/**
 * Standard evaluation period column names
 * Centralized configuration
 */
const EVALUATION_PERIOD_COLUMNS = [
  'Periodo_Base_EVA_Mayo',
  'Periodo_Base_EVA_Noviembre', 
  'Diagnostico_Mayo_25',
  'Evaluacion_1_Noviembre_25',
  'Evaluacion_2_Mayo_26',
  'Evaluacion_3_Noviembre_26',
  'Evaluacion_4_Mayo_27',
  'Evaluacion_5_Octubre_27'
];

/**
 * Standard dataset column mappings
 * Centralized configuration
 */
const STANDARD_COLUMN_MAPPING = {
  "municipio": { hoja1: "municipio", hoja2: "municipio" },
  "id_indicador": { hoja1: "id_indicador", hoja2: "id_indicador" },
  "eje": { hoja1: "eje", hoja2: "eje" },
  "tema": { hoja1: "tema", hoja2: "tema" },
  "nombre": { hoja1: "nombre_componente", hoja2: "nombre_indicador" },
  "codigo": { hoja1: "codigo_componente", hoja2: "calculo_indicador"},
  "tipo de dato": { hoja1: null, hoja2: null }
};

/**
 * Month name to number mapping
 * Centralized configuration
 */
const MONTH_ORDER = {
  "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
  "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
};