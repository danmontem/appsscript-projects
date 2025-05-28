/**
 * Formula Calculation Functions
 */

/**
 * Test function for debugging component values
 */
function testComponentValue() {
  const result = COMPONENT_VALUE("MRez", "ENE24", "Apodaca");
  console.log("Result: " + result);
}

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
  
  // Convert month code to actual month name
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
      // Let the calculation functions handle the logic for combining values
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
  
  // Map of month names to their numerical order
  const monthOrder = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
  };
  
  // Parse start and end months
  const [startMonthName, startYear] = startMonth.toLowerCase().split(" ");
  const [endMonthName, endYear] = endMonth.toLowerCase().split(" ");
  
  const startDate = new Date(parseInt(startYear), monthOrder[startMonthName] - 1);
  const endDate = new Date(parseInt(endYear), monthOrder[endMonthName] - 1);
  
  // Find relevant month columns within the range
  const monthColumns = [];
  for (let j = 0; j < headers.length; j++) {
    const header = headers[j];
    
    if (!header.includes(" 20")) continue;
    
    const [headerMonthName, headerYear] = header.split(" ");
    const headerDate = new Date(parseInt(headerYear), monthOrder[headerMonthName] - 1);
    
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
        // or a combined indicator if there are different types
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

/** Test sum function */
function testComponentSum() {
  const result = COMPONENT_SUM("MRez", "enero 2024", "diciembre 2024", "Apodaca");
  console.log("Sum result: " + result);
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

/**
 * Custom function to calculate an indicator with complex formulas
 * Modified to auto-refresh when data changes
 * Enhanced CALCULATE_INDICATOR function that handles string values appropriately
 * 
 * @param {string} formula - The formula to calculate
 * @param {string} municipio - The municipality name
 * @param {number=} timestamp - Optional timestamp parameter
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
    if (error.message.includes("Mixed data types") || error.message.includes("SD") || error.message.includes("información reservada")) {
      return error.message;
    } else {
      throw error; // Re-throw other types of errors
    }
  }
}

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

/**
 * Processes a complex formula with operators like +, -, *, /, etc.
 * 
 * @param {string} formula - The formula to process
 * @param {string} municipio - The municipality name
 * @return The calculated result
 * Enhanced processComplexFormula that handles string values intelligently
 */
function processComplexFormula(formula, municipio) {
  if (!formula) {
    throw new Error("Formula cannot be undefined or empty");
  }
  
  console.log("Processing formula: " + formula);
  
  // Check if the entire formula is already a string value (like "SD", "información reservada", etc.)
  // If it doesn't contain any operators or function calls, it might be a direct string
  if (!/[+\-*/():]|SUM|AVG|SINGLE|MAX|MIN/.test(formula)) {
    // It's likely a direct string value
    return formula;
  }
  
  // First handle any nested function calls like SUM() or SINGLE()
  const functionRegex = /(SUM|SINGLE|AVG|MAX|MIN)\(([^()]+)\)/g;
  let modifiedFormula = formula;
  let functionMatch;
  
  while ((functionMatch = functionRegex.exec(formula)) !== null) {
    const [fullMatch, functionName, content] = functionMatch;
    console.log(`Found function: ${functionName}(${content})`);
    
    let result;
    try {
      if (functionName === "SUM") {
        const components = content.split(",").map(c => c.trim());
        const values = components.map(component => evaluateComponentPart(component, municipio));
        
        // Check if any values are strings
        const stringValues = values.filter(v => typeof v === "string");
        const numericValues = values.filter(v => typeof v === "number");
        
        if (stringValues.length > 0 && numericValues.length > 0) {
          return `ERROR: Mixed data types - ${stringValues.join(", ")}`;
        } else if (stringValues.length > 0) {
          const uniqueStrings = [...new Set(stringValues)];
          return uniqueStrings.length === 1 ? uniqueStrings[0] : `MIXED: ${uniqueStrings.join(", ")}`;
        } else {
          result = numericValues.reduce((sum, val) => sum + val, 0);
        }
      } 
      else if (functionName === "SINGLE") {
        const [code, monthCode] = content.split(":");
        const month = convertMonthCode(monthCode);
        const value = COMPONENT_VALUE(code, month, municipio);
        
        // If SINGLE returns a string, return it immediately
        if (typeof value === "string") {
          return value;
        }
        result = value;
      }
      else if (functionName === "AVG") {
        const components = content.split(",").map(c => c.trim());
        const values = components.map(component => evaluateComponentPart(component, municipio));
        
        const stringValues = values.filter(v => typeof v === "string");
        const numericValues = values.filter(v => typeof v === "number");
        
        if (stringValues.length > 0 && numericValues.length > 0) {
          return `ERROR: Mixed data types - ${stringValues.join(", ")}`;
        } else if (stringValues.length > 0) {
          const uniqueStrings = [...new Set(stringValues)];
          return uniqueStrings.length === 1 ? uniqueStrings[0] : `MIXED: ${uniqueStrings.join(", ")}`;
        } else {
          result = numericValues.reduce((sum, val) => sum + val, 0) / numericValues.length;
        }
      }
      else if (functionName === "MAX" || functionName === "MIN") {
        const components = content.split(",").map(c => c.trim());
        const values = components.map(component => evaluateComponentPart(component, municipio));
        
        const stringValues = values.filter(v => typeof v === "string");
        const numericValues = values.filter(v => typeof v === "number");
        
        if (stringValues.length > 0) {
          const uniqueStrings = [...new Set(stringValues)];
          return uniqueStrings.length === 1 ? uniqueStrings[0] : `MIXED: ${uniqueStrings.join(", ")}`;
        } else {
          result = functionName === "MAX" ? Math.max(...numericValues) : Math.min(...numericValues);
        }
      }
    } catch (e) {
      return `ERROR: ${e.message}`;
    }
    
    console.log(`Function ${functionName} result: ${result}`);
    
    // Replace the function call with its result
    modifiedFormula = modifiedFormula.replace(fullMatch, result);
    functionRegex.lastIndex = 0;
    formula = modifiedFormula;
  }
  
  // Handle remaining component references
  const componentRefRegex = /([A-Za-z0-9-_]+):([A-Za-z0-9-]+(-[A-Za-z0-9-]+)?)/g;
  let componentMatch;
  
  while ((componentMatch = componentRefRegex.exec(modifiedFormula)) !== null) {
    const fullMatch = componentMatch[0];
    console.log("Found component reference: " + fullMatch);
    
    try {
      const componentValue = evaluateComponentPart(fullMatch, municipio);
      console.log("Component value: " + componentValue);
      
      // If it's a string, return it immediately - don't try to continue processing
      if (typeof componentValue === "string") {
        return componentValue;
      }
      
      modifiedFormula = modifiedFormula.replace(fullMatch, componentValue);
      componentRefRegex.lastIndex = 0;
    } catch (e) {
      return `ERROR: ${e.message}`;
    }
  }
  
  console.log("Final formula to evaluate: " + modifiedFormula);
  
  // At this point, if we still have a string that's not purely mathematical, return it
  if (typeof modifiedFormula === "string") {
    // Check if it's a string that contains non-mathematical characters
    if (!/^[0-9.+\-*/()E\s]*$/.test(modifiedFormula)) {
      return modifiedFormula;
    }
  }
  
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
 * Helper function to convert month codes to month names
 * This updated version handles both codes and full month names
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
