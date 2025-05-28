/**
 * Refresh Management Functions
 */

/**
 * Helper function to create a timer-based trigger to refresh calculation cells
 * This is called once to set up periodic refreshing
 */
function createDataRefreshTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'refreshCalculationCells') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Create a new trigger that runs every 5 minutes
  ScriptApp.newTrigger('refreshCalculationCells')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  SpreadsheetApp.getUi().alert('Auto-refresh trigger created. Calculation cells will update every 5 minutes.');
}

/**
 * Function that runs on a timer to force recalculation of formula cells
 */
function refreshCalculationCells() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Add a temporary timestamp to a hidden cell to force recalculation
  const timestampCell = sheet.getRange(1, sheet.getMaxColumns());
  timestampCell.setValue(new Date().getTime());
  
  // Clear the timestamp cell after a brief delay (to ensure formulas recalculate)
  Utilities.sleep(1000);
  timestampCell.clearContent();
}

/**
 * Menu item to manually refresh all calculation cells
 */
function manualRefreshCalculations() {
  refreshCalculationCells();
  SpreadsheetApp.getUi().alert('All calculation cells have been refreshed with current data.');
}
