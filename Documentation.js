/**
 * Municipality Tools Documentation
 * ================================
 * 
 * This file serves as technical documentation for the Municipality Tools Apps Script project.
 * It provides detailed information about the project structure, functionality, and maintenance.
 * 
 * Project Overview
 * ----------------
 * This Apps Script project provides comprehensive tools for managing municipal evaluation data,
 * generating datasets of components and indicators, performing complex calculations, and managing
 * evaluation periods. It is designed for municipal assessment workflows with automated data
 * processing and preservation of manual entries.
 * 
 * UPDATED File Structure (Post-Restructuring)
 * -------------------------------------------
 * The project has been reorganized for better maintainability and performance:
 * 
 * CORE SYSTEM:
 * - Main.js: Menu system, entry points, and diagnostic functions
 * - CommonHelpers.js: Centralized utility functions (eliminates duplicates)
 * 
 * DATA MANAGEMENT:
 * - ComponentsManager.js: Components dataset generation and management
 * - IndicatorsManager.js: Indicators dataset generation and management
 * - DataCombiner.js: Smart dataset combination with formula preservation
 * 
 * CALCULATIONS:
 * - FormulaCalculator.js: Core calculation engine with enhanced string handling
 * 
 * EVALUATION SYSTEM:
 * - PeriodsManager.js: Consolidated evaluation periods management system
 * 
 * USER INTERFACE:
 * - UIHelpers.js: Dialog and UI helper functions
 * - RefreshManager.js: Auto-refresh and trigger management
 * 
 * UTILITIES:
 * - createNamedRanges.js: Named range creation utility
 * - deleteNamedRanges.js: Named range cleanup utility
 * - Documentation.js: This documentation file
 * 
 * HTML INTERFACES:
 * - GenerateComponentsDialog.html: Components dataset generation UI
 * - GenerateIndicatorsDialog.html: Indicators dataset generation UI
 * 
 * CONFIGURATION:
 * - appsscript.json: Project configuration and permissions
 * 
 * Key Functionality
 * ----------------
 * 
 * 1. DATASET MANAGEMENT
 *    - Generate Components Dataset: Creates comprehensive municipal components datasets
 *    - Generate Indicators Dataset: Creates municipal indicators datasets
 *    - Smart Dataset Combination: Combines components and indicators with preservation
 *    - Intelligent Refresh: Updates datasets while preserving user modifications
 * 
 * 2. EVALUATION PERIODS SYSTEM
 *    - Create Evaluation Periods Dataset: Generates time-based evaluation structures
 *    - External Periods Mapping: Maps evaluation dates from external sources
 *    - Smart Refresh: Automated workflow for updating all periods data
 *    - Manual Entry Preservation: Protects user-entered data during updates
 * 
 * 3. ADVANCED CALCULATIONS
 *    - COMPONENT_VALUE(): Retrieves component values with string handling (SD, etc.)
 *    - COMPONENT_SUM(): Intelligent summing with mixed data type handling
 *    - CALCULATE_INDICATOR(): Complex indicator calculations with formula processing
 *    - Formula Generation: Automated formula creation based on evaluation periods
 * 
 * 4. SMART REFRESH SYSTEM
 *    - Configurable automated refresh workflows
 *    - Preservation of manual entries and custom formulas
 *    - Batch processing for performance optimization
 *    - Error recovery and detailed logging
 * 
 * 5. DEBUGGING AND DIAGNOSTICS
 *    - Fast diagnostic tools for troubleshooting
 *    - External document access testing
 *    - Periods dataset analysis and validation
 *    - Mapping logic debugging tools
 * 
 * Usage Instructions
 * -----------------
 * 
 * INITIAL SETUP:
 * 1. Open the spreadsheet and run the script to create the "Municipality Tools" menu
 * 2. Use "Test Document Access" to verify external document connections
 * 3. Configure your source documents and sharing permissions
 * 
 * BASIC WORKFLOW:
 * 1. Generate Components Dataset from external source
 * 2. Generate Indicators Dataset from external source
 * 3. Combine Components & Indicators into unified dataset
 * 4. Create Evaluation Periods Dataset for time-based analysis
 * 5. Configure Smart Refresh for automated updates
 * 
 * EVALUATION PERIODS WORKFLOW:
 * 1. Create base evaluation periods dataset
 * 2. Configure Smart Refresh with external periods document
 * 3. Map periods from external table (preserves manual entries)
 * 4. Generate calculation formulas based on periods
 * 5. Use automated refresh to keep data current
 * 
 * CALCULATION WORKFLOW:
 * 1. Create Formula Builder for reference
 * 2. Use COMPONENT_VALUE() for single month values
 * 3. Use COMPONENT_SUM() for date range calculations
 * 4. Use CALCULATE_INDICATOR() for complex formulas
 * 5. Set up auto-refresh triggers for live updates
 * 
 * Advanced Features
 * ----------------
 * 
 * SMART REFRESH SYSTEM:
 * - Configures complete refresh workflows
 * - Preserves manual entries during updates
 * - Handles mixed data types (numeric, "SD", "información reservada")
 * - Provides detailed execution logging
 * 
 * FORMULA PRESERVATION:
 * - Detects formula columns vs data columns
 * - Preserves complex calculations during dataset refreshes
 * - Maintains custom column ordering
 * - Handles cross-sheet references
 * 
 * ERROR HANDLING:
 * - Comprehensive error reporting with context
 * - Graceful handling of missing data
 * - Recovery mechanisms for interrupted processes
 * - Detailed diagnostic information
 * 
 * DATA VALIDATION:
 * - Automatic ID normalization and validation
 * - Period text standardization
 * - Column mapping verification
 * - Data type consistency checks
 * 
 * Configuration Options
 * --------------------
 * 
 * STANDARDIZED COLUMN MAPPINGS (in CommonHelpers.js):
 * - Components: municipio, id_indicador, eje, tema, nombre_componente, codigo_componente
 * - Indicators: municipio, id_indicador, eje, tema, nombre_indicador, calculo_indicador
 * 
 * EVALUATION PERIOD COLUMNS:
 * - Periodo_Base_EVA_Mayo, Periodo_Base_EVA_Noviembre
 * - Diagnostico_Mayo_25, Evaluacion_1_Noviembre_25
 * - Evaluacion_2_Mayo_26, Evaluacion_3_Noviembre_26
 * - Evaluacion_4_Mayo_27, Evaluacion_5_Octubre_27
 * 
 * AUTO-REFRESH SETTINGS:
 * - Default: Every 5 minutes for calculation updates
 * - Configurable triggers for dataset refreshes
 * - Batch processing optimization settings
 * 
 * Troubleshooting
 * --------------
 * 
 * COMMON ISSUES:
 * 1. External Document Access: Check sharing permissions and document IDs
 * 2. Column Mapping Errors: Verify source sheet column names match expectations
 * 3. Mixed Data Types: System handles "SD" and other string values automatically
 * 4. Performance Issues: Use batch processing and avoid frequent manual refreshes
 * 
 * DIAGNOSTIC TOOLS:
 * - Fast Periods Dataset Diagnostic: Quick analysis of first 50 rows
 * - Detailed Periods Dataset Analysis: Comprehensive data structure review
 * - Debug Periods Mapping: Step-by-step mapping logic verification
 * - Test External Document Access: Connection and permissions testing
 * 
 * DEBUGGING TECHNIQUES:
 * 1. Check browser console for detailed error logs
 * 2. Use diagnostic functions to isolate issues
 * 3. Test with small datasets before full processing
 * 4. Verify all dependencies are loaded correctly
 * 
 * Maintenance Guidelines
 * ---------------------
 * 
 * CODE ORGANIZATION:
 * - CommonHelpers.js contains all shared utility functions
 * - Each module has specific, focused responsibilities
 * - Dependencies are clearly documented in file headers
 * - All duplicate functions have been eliminated
 * 
 * ADDING NEW FEATURES:
 * 1. Add utility functions to CommonHelpers.js
 * 2. Add data management functions to appropriate manager files
 * 3. Add menu items to Main.js
 * 4. Update this documentation
 * 
 * PERFORMANCE OPTIMIZATION:
 * - Use batch operations for large datasets
 * - Implement caching for frequently accessed data
 * - Minimize API calls through intelligent batching
 * - Use progress indicators for long-running operations
 * 
 * TESTING:
 * - Test all functions after modifications
 * - Use diagnostic tools to verify data integrity
 * - Test refresh operations with preserved data
 * - Validate external document connections
 * 
 * Data Structure Requirements
 * --------------------------
 * 
 * SOURCE SHEETS (External Documents):
 * - Indicadores_Componentes: Component definitions and mappings
 * - Banco_indicadores_ACV: Indicator definitions and formulas
 * - Evaluation Periods Table: Time period definitions for assessments
 * 
 * INTERNAL SHEETS:
 * - BD_componentes: Core components data (critical for calculations)
 * - Combined datasets: Merged components and indicators
 * - Evaluation periods datasets: Time-based analysis structures
 * 
 * DATA FORMATS:
 * - Dates: "month year" format (e.g., "marzo 2025")
 * - Month codes: OCT24, MAR25 format for formula references
 * - IDs: Numeric indicator IDs with 0 for manual-entry components
 * - Periods: Flexible text with automatic normalization
 * 
 * API and Integration Points
 * -------------------------
 * 
 * GOOGLE APPS SCRIPT APIS USED:
 * - SpreadsheetApp: Core spreadsheet operations
 * - PropertiesService: Configuration storage
 * - HtmlService: User interface dialogs
 * - ScriptApp: Trigger management
 * - Utilities: Sleep and utility functions
 * 
 * EXTERNAL INTEGRATIONS:
 * - Cross-document data access via SpreadsheetApp.openById()
 * - Automated trigger-based refreshes
 * - HTML dialog interfaces for user input
 * 
 * Future Enhancement Roadmap
 * -------------------------
 * 
 * SHORT TERM:
 * 1. Add data export functionality (PDF, Excel)
 * 2. Implement user preference storage
 * 3. Add visualization components
 * 4. Enhance error reporting with email notifications
 * 
 * MEDIUM TERM:
 * 1. Integration with external APIs
 * 2. Advanced analytics and reporting
 * 3. Multi-language support
 * 4. Mobile-responsive interfaces
 * 
 * LONG TERM:
 * 1. Machine learning integration for predictions
 * 2. Real-time collaboration features
 * 3. Advanced workflow automation
 * 4. Integration with municipal systems
 * 
 * Version History
 * --------------
 * - 1.0.0 (Initial): Basic dataset generation and calculations
 * - 1.1.0: Added component code support and improved refresh reliability
 * - 1.2.0: Restructured code into multiple files for better organization
 * - 2.0.0 (Current): Major restructuring with consolidated utilities and enhanced features
 *   * Eliminated all duplicate functions
 *   * Consolidated evaluation periods system
 *   * Enhanced formula preservation
 *   * Improved error handling and diagnostics
 *   * Added smart refresh workflows
 *   * Optimized performance with batch processing
 * 
 * Technical Specifications
 * -----------------------
 * 
 * PERFORMANCE CHARACTERISTICS:
 * - Batch size: 50 rows for optimal processing
 * - Timeout handling: Built-in delays for large operations
 * - Memory optimization: Efficient data structure usage
 * - Error recovery: Graceful degradation and retry mechanisms
 * 
 * BROWSER COMPATIBILITY:
 * - Optimized for Google Apps Script V8 runtime
 * - Compatible with all modern browsers via Google Sheets interface
 * - HTML dialogs work across desktop and mobile platforms
 * 
 * SECURITY CONSIDERATIONS:
 * - Secure external document access with proper authorization
 * - Input validation and sanitization
 * - No sensitive data storage in properties
 * - Audit trail through console logging
 * 
 * Last Updated: December 2024
 * Project Status: Active Development
 * 
 * Support and Contact
 * ------------------
 * For technical support, feature requests, or bug reports:
 * - Review this documentation for common solutions
 * - Use built-in diagnostic tools for troubleshooting
 * - Check console logs for detailed error information
 * - Test with smaller datasets to isolate issues
 */

// This file contains no executable code - it serves as comprehensive documentation
function documentation() {
  // Placeholder function to make this file appear as a valid script file
  // The actual documentation is in the comments above
  return "Municipality Tools Documentation - Version 2.0.0";
}