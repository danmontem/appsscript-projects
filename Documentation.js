/**
 * Municipality Tools Documentation
 * ================================
 * 
 * This file serves as documentation for the Municipality Tools Apps Script project.
 * It does not contain executable code but provides an overview of the project,
 * its structure, and how to use and maintain it.
 * 
 * Project Overview
 * ----------------
 * This Apps Script project provides tools for managing municipal data,
 * generating datasets of components and indicators, and performing calculations
 * on this data. It is designed to work with a specific data structure and
 * includes functionality for refreshing data while preserving user inputs.
 * 
 * File Structure
 * --------------
 * The project is organized into the following files:
 * 
 * - Main.gs: Core setup and menu creation
 * - UIHelpers.gs: Dialog and UI helper functions
 * - ComponentsManager.gs: Components dataset generation and management
 * - IndicatorsManager.gs: Indicators dataset generation and management
 * - FormulaCalculator.gs: Calculation functions and utilities
 * - RefreshManager.gs: Functions for refreshing data and calculations
 * - Documentation.gs: This documentation file
 * 
 * HTML Files:
 * - GenerateComponentsDialog.html: User interface for generating components dataset
 * - GenerateIndicatorsDialog.html: User interface for generating indicators dataset
 * 
 * Key Functionality
 * ----------------
 * 
 * 1. Dataset Generation
 *    - Generate Components Dataset: Creates a dataset of municipal components
 *    - Generate Indicators Dataset: Creates a dataset of municipal indicators
 * 
 * 2. Dataset Refreshing
 *    - Refresh Components Dataset: Updates the components dataset while preserving user-entered values
 *    - Refresh Indicators Dataset: Updates the indicators dataset while preserving custom columns
 * 
 * 3. Calculation Functions
 *    - COMPONENT_VALUE(): Gets a component value for a specific month and municipality
 *    - COMPONENT_SUM(): Sums a component's values over a date range
 *    - CALCULATE_INDICATOR(): Calculates complex indicators using formulas
 * 
 * 4. Utilities
 *    - Formula Builder: Creates a helper sheet for building formulas
 *    - Auto-Refresh: Sets up automatic refreshing of calculations
 * 
 * Usage Instructions
 * -----------------
 * 
 * Initial Setup:
 * 1. Run the script for the first time to create the custom menu
 * 2. Use "Test Document Access" to verify connections to source documents
 * 
 * Creating Datasets:
 * 1. Select "Generate Components Dataset" from the menu
 * 2. Enter the source document ID and select target sheet
 * 3. Select municipalities to include
 * 4. Click "Generate" to create the dataset
 * 
 * Using Calculation Functions:
 * 1. Create a Formula Builder using the menu option
 * 2. Use the provided examples to build formulas
 * 3. Insert formulas into cells using the CALCULATE_INDICATOR() function
 * 
 * Setting Up Auto-Refresh:
 * 1. Select "Setup Auto-Refresh" from the menu to create a trigger
 * 2. Calculations will refresh every 5 minutes
 * 
 * Maintenance
 * -----------
 * 
 * Code Organization:
 * - Each file contains related functions to make maintenance easier
 * - When adding new features, create functions in the appropriate file
 * 
 * Debugging:
 * - Use the console.log() statements in the code to help debug issues
 * - Test functions are provided for key functionality
 * 
 * Data Structure Requirements:
 * - The components dataset expects specific source sheets and column structures
 * - The BD_componentes sheet is critical for calculations to work
 * 
 * Potential Future Enhancements
 * -----------------------------
 * 
 * 1. Add support for additional calculation types
 * 2. Implement data validation for user inputs
 * 3. Create visualization tools for indicator data
 * 4. Add export functionality to PDF or other formats
 * 
 * Version History
 * --------------
 * - 1.0.0 (Initial Version): Basic functionality for dataset generation and calculations
 * - 1.1.0: Added component code support and improved refresh reliability
 * - 1.2.0: Restructured code into multiple files for better organization
 * 
 * Last Updated: [Insert Date Here]
 * 
 * Developer Contact Information
 * ----------------------------
 * [Your Name/Team]
 * [Contact Email]
 */

// This file contains no executable code
function documentation() {
  // This function does nothing but serves as a placeholder
  // so that this file appears as a valid script file
}