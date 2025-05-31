/**
 * FormulaCalculator.js
 * Cleaned and optimized formula calculation functions
 * Dependencies: CommonHelpers.js
 */

// =============================================
// MAIN CALCULATION FUNCTIONS
// =============================================

/**
 * Enhanced COMPONENT_VALUE function that handles string values like "SD"
 * 
 * @param {string} componentCode - The code of the component
 * @param {string} month - The month in format "month year" (e.g., "marzo 2025")
 * @param {string} municipio - The municipality name
 * @return The value of the component for the specified month (numeric or string)
 * @customfunction
 */
function COMPONENT_VALUE(componentCode, month, municipio) {
  // Use the correct sheet name where your components are stored
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_componentes");
  
  // Check if sheet exists
  if (!sheet) {
    throw new Error("Sheet 'BD_componentes' not found. Please specify the correct sheet name.");
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Convert month code to actual month name (using CommonHelpers)
  const fullMonth = convertMonthCode(month);
  
  // Find column indices
  const municipioCol = 0; // Column A (municipio)
  const componentCodeCol = 3; // Column D (codigo_componente)
  const monthCol = headers.indexOf(fullMonth.toLowerCase());
  
  if (monthCol === -1) {
    throw new Error(`Month "${fullMonth}" not found in headers. Available headers: ${headers.join(", ")}`);
  }
  
  // Search for the row with matching component code and municipality
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[municipioCol] === municipio && row[componentCodeCol] === componentCode) {
      const value = row[monthCol];
      
      // Return the value as-is (whether it's numeric, "SD", "información reservada", etc.)
      return value;
    }
  }
  
  throw new Error(`Component "${componentCode}" not found for municipality "${municipio}"`);
}

/**
 * Enhanced COMPONENT_SUM function that handles string values
 * 
 * @param {string} componentCode - The code of the component
 * @param {string} startMonth - The start month in format "month year"
 * @param {string} endMonth - The end month in format "month year"
 * @param {string} municipio - The municipality name
 * @return The sum of numeric values, or a string indicating data issues
 * @customfunction
 */
function COMPONENT_SUM(componentCode, startMonth, endMonth, municipio) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_componentes");
  
  if (!sheet) {
    throw new Error("Sheet 'BD_componentes' not found.");
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const municipioCol = 0;
  const componentCodeCol = 3;
  
  // Parse start and end months (using CommonHelpers MONTH_ORDER)
  const [startMonthName, startYear] = startMonth.toLowerCase().split(" ");
  const [endMonthName, endYear] = endMonth.toLowerCase().split(" ");
  
  const startDate = new Date(parseInt(startYear), MONTH_ORDER[startMonthName] - 1);
  const endDate = new Date(parseInt(endYear), MONTH_ORDER[endMonthName] - 1);
  
  // Find relevant month columns within the range
  const monthColumns = [];
  for (let j = 0; j < headers.length; j++) {
    const header = headers[j];
    
    if (!header.includes(" 20")) continue;
    
    const [headerMonthName, headerYear] = header.split(" ");
    const headerDate = new Date(parseInt(headerYear), MONTH_ORDER[headerMonthName] - 1);
    
    if (headerDate >= startDate && headerDate <= endDate) {
      monthColumns.push(j);
    }
  }
  
  if (monthColumns.length === 0) {
    throw new Error(`No months found between "${startMonth}" and "${endMonth}"`);
  }
  
  // Search for the row with matching component code and municipality
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[municipioCol] === municipio && row[componentCodeCol] === componentCode) {
      
      // Collect all values in the date range
      const values = [];
      const stringValues = [];
      
      for (const colIndex of monthColumns) {
        const value = row[colIndex];
        
        if (value === "" || value === null || value === undefined) {
          // Skip empty cells
          continue;
        } else if (typeof value === "string" && isNaN(value)) {
          // It's a non-numeric string (like "SD", "información reservada")
          stringValues.push(value);
        } else if (!isNaN(value)) {
          // It's a numeric value
          values.push(parseFloat(value));
        }
      }
      
      // Decision logic for what to return
      if (stringValues.length > 0 && values.length > 0) {
        // Mixed data: some strings, some numbers - this indicates an error condition
        throw new Error(`Mixed data types found for ${componentCode} in ${municipio}: numeric values and strings (${stringValues.join(", ")})`);
      } else if (stringValues.length > 0 && values.length === 0) {
        // All non-empty values are strings - return the first string found
        const uniqueStrings = [...new Set(stringValues)];
        if (uniqueStrings.length === 1) {
          return uniqueStrings[0]; // All strings are the same (e.g., all "SD")
        } else {
          return `MIXED: ${uniqueStrings.join(", ")}`; // Different types of strings
        }
      } else if (values.length > 0) {
        // All values are numeric - return the sum
        return values.reduce((sum, val) => sum + val, 0);
      } else {
        // All cells are empty
        return 0;
      }
    }
  }
  
  throw new Error(`Component "${componentCode}" not found for municipality "${municipio}"`);
}

/**
 * Custom function to calculate an indicator with complex formulas
 * Enhanced CALCULATE_INDICATOR function that handles string values appropriately
 * 
 * @param {string} formula - The formula to calculate
 * @param {string} municipio - The municipality name
 * @param {number=} timestamp - Optional timestamp parameter for auto-refresh
 * @return The calculated indicator value or a string indicating data issues
 * @customfunction
 */
function CALCULATE_INDICATOR(formula, municipio, timestamp) {
  if (timestamp === undefined) {
    timestamp = new Date().getTime();
  }
  
  console.log("CALCULATE_INDICATOR called with formula: " + formula);
  console.log("Municipality: " + municipio);
  
  if (!formula || !municipio) {
    throw new Error("Formula and municipality must be provided");
  }
  
  try {
    // Process the formula with the enhanced formula processor
    return processComplexFormula(formula, municipio);
  } catch (error) {
    // If there's an error and it mentions data issues, return a descriptive message
    if (error.message.includes("Mixed data types") || 
        error.message.includes("SD") || 
        error.message.includes("información reservada")) {
      return error.message;
    } else {
      throw error; // Re-throw other types of errors
    }
  }
}

// =============================================
// FORMULA PROCESSING ENGINE
// =============================================

/**
 * Processes a complex formula with operators like +, -, *, /, etc.
 * Enhanced processComplexFormula that handles string values intelligently
 * 
 * @param {string} formula - The formula to process
 * @param {string} municipio - The municipality name
 * @return The calculated result
 */
function processComplexFormula(formula, municipio) {
  if (!formula) {
    throw new Error("Formula cannot be undefined or empty");
  }
  
  console.log("Processing formula: " + formula);
  
  // Check if the entire formula is already a string value
  if (!/[+\-*/():]|SUM|AVG|SINGLE|MAX|MIN/.test(formula)) {
    return formula;
  }
  
  // Handle nested function calls like SUM() or SINGLE()
  const functionRegex = /(SUM|SINGLE|AVG|MAX|MIN)\(([^()]+)\)/g;
  let modifiedFormula = formula;
  let functionMatch;
  
  while ((functionMatch = functionRegex.exec(formula)) !== null) {
    const [fullMatch, functionName, content] = functionMatch;
    console.log(`Found function: ${functionName}(${content})`);
    
    let result;
    try {
      result = processFunction(functionName, content, municipio);
    } catch (e) {
      return `ERROR: ${e.message}`;
    }
    
    console.log(`Function ${functionName} result: ${result}`);
    
    // If function returns a string, return it immediately
    if (typeof result === "string" && isNaN(result)) {
      return result;
    }
    
    // Replace the function call with its result
    modifiedFormula = modifiedFormula.replace(fullMatch, result);
    functionRegex.lastIndex = 0;
    formula = modifiedFormula;
  }
  
  // Handle remaining component references
  modifiedFormula = processComponentReferences(modifiedFormula, municipio);
  
  // If we got a string result from component processing, return it
  if (typeof modifiedFormula === "string" && !/^[0-9.+\-*/()E\s]*$/.test(modifiedFormula)) {
    return modifiedFormula;
  }
  
  console.log("Final formula to evaluate: " + modifiedFormula);
  
  // Clean the formula for mathematical evaluation
  modifiedFormula = modifiedFormula.replace(/\s+/g, "");
  
  // Final check before eval - make sure it only contains mathematical operators
  if (!/^[0-9.+\-*/()E]*$/.test(modifiedFormula)) {
    return `ERROR: Non-mathematical characters found: ${modifiedFormula}`;
  }
  
  // Evaluate the resulting mathematical expression
  try {
    const result = eval(modifiedFormula);
    return result;
  } catch (e) {
    return `ERROR: Mathematical evaluation failed - ${e.message}`;
  }
}

/**
 * Process a specific function (SUM, AVG, etc.) with enhanced string handling
 */
function processFunction(functionName, content, municipio) {
  const components = content.split(",").map(c => c.trim());
  const values = components.map(component => evaluateComponentPart(component, municipio));
  
  // Separate string and numeric values
  const stringValues = values.filter(v => typeof v === "string" && isNaN(v));
  const numericValues = values.filter(v => typeof v === "number" || !isNaN(v)).map(v => Number(v));
  
  // Handle mixed data types
  if (stringValues.length > 0 && numericValues.length > 0) {
    return `ERROR: Mixed data types - ${stringValues.join(", ")}`;
  } else if (stringValues.length > 0) {
    const uniqueStrings = [...new Set(stringValues)];
    return uniqueStrings.length === 1 ? uniqueStrings[0] : `MIXED: ${uniqueStrings.join(", ")}`;
  }
  
  // Process numeric values based on function type
  switch (functionName) {
    case "SUM":
      return numericValues.reduce((sum, val) => sum + val, 0);
    case "AVG":
      return numericValues.reduce((sum, val) => sum + val, 0) / numericValues.length;
    case "MAX":
      return Math.max(...numericValues);
    case "MIN":
      return Math.min(...numericValues);
    case "SINGLE":
      // For SINGLE, handle the special case
      const [code, monthCode] = content.split(":");
      const month = convertMonthCode(monthCode);
      return COMPONENT_VALUE(code, month, municipio);
    default:
      throw new Error(`Unknown function: ${functionName}`);
  }
}

/**
 * Process component references in formula
 */
function processComponentReferences(formula, municipio) {
  const componentRefRegex = /([A-Za-z0-9-_]+):([A-Za-z0-9-]+(-[A-Za-z0-9-]+)?)/g;
  let modifiedFormula = formula;
  let componentMatch;
  
  while ((componentMatch = componentRefRegex.exec(modifiedFormula)) !== null) {
    const fullMatch = componentMatch[0];
    console.log("Found component reference: " + fullMatch);
    
    try {
      const componentValue = evaluateComponentPart(fullMatch, municipio);
      console.log("Component value: " + componentValue);
      
      // If it's a string, return it immediately - don't try to continue processing
      if (typeof componentValue === "string" && isNaN(componentValue)) {
        return componentValue;
      }
      
      modifiedFormula = modifiedFormula.replace(fullMatch, componentValue);
      componentRefRegex.lastIndex = 0;
    } catch (e) {
      return `ERROR: ${e.message}`;
    }
  }
  
  return modifiedFormula;
}

/**
 * Evaluates a specific component part from the formula
 * 
 * @param {string} componentPart - The component part of the formula (e.g., "HTA:OCT24-MAR25")
 * @param {string} municipio - The municipality name
 * @return The evaluated value
 */
function evaluateComponentPart(componentPart, municipio) {
  // Parse component format
  if (componentPart.includes(":")) {
    // Format is CODE:DATERANGE
    const [code, dateRange] = componentPart.split(":");
    
    if (dateRange.includes("-")) {
      // Format is CODE:STARTMONTH-ENDMONTH (date range)
      const [startMonthCode, endMonthCode] = dateRange.split("-");
      const startMonth = convertMonthCode(startMonthCode);
      const endMonth = convertMonthCode(endMonthCode);
      return COMPONENT_SUM(code, startMonth, endMonth, municipio);
    } else {
      // Format is CODE:MONTH (single month)
      const month = convertMonthCode(dateRange);
      return COMPONENT_VALUE(code, month, municipio);
    }
  } else {
    // Try to parse as a number
    const numValue = parseFloat(componentPart);
    if (!isNaN(numValue)) {
      return numValue;
    }
    throw new Error(`Invalid component part format: ${componentPart}`);
  }
}

// =============================================
// HELPER AND UTILITY FUNCTIONS
// =============================================

/**
 * Helper function to get component data for change detection
 */
function getComponentData(componentCode, municipio) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_componentes");
  
  if (!sheet) {
    throw new Error("Sheet 'BD_componentes' not found.");
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const municipioCol = 0; // Column A (municipio)
  const componentCodeCol = 3; // Column D (codigo_componente)
  
  // Search for the row with matching component code and municipality
  for (let i = 1; i < data.length; i++) {
    if (data[i][municipioCol] === municipio && data[i][componentCodeCol] === componentCode) {
      // Get all monthly data from this row
      const monthValues = {};
      for (let j = 9; j < headers.length; j++) {
        // Only include columns that look like month names
        if (headers[j].includes(" 20")) {
          monthValues[headers[j]] = data[i][j];
        }
      }
      
      // Return data including row number for reference
      return {
        rowNumber: i + 1,
        months: monthValues
      };
    }
  }
  
  throw new Error(`Component "${componentCode}" not found for municipality "${municipio}"`);
}

// =============================================
// TESTING FUNCTIONS
// =============================================

/**
 * Test function for debugging component values
 */
function testComponentValue() {
  const result = COMPONENT_VALUE("MRez", "ENE24", "Apodaca");
  console.log("Result: " + result);
}

/**
 * Test sum function
 */
function testComponentSum() {
  const result = COMPONENT_SUM("MRez", "enero 2024", "diciembre 2024", "Apodaca");
  console.log("Sum result: " + result);
}