# Municipality Tools - Google Apps Script

> 🏛️ **Comprehensive municipal evaluation data management system with automated calculations and smart refresh capabilities**

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8-blue.svg)](https://script.google.com)
[![Version](https://img.shields.io/badge/version-2.0.0-green.svg)](https://github.com/your-repo/municipality-tools)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

## 🌟 Features

- **📊 Smart Dataset Management** - Generate and combine municipal components and indicators
- **📅 Evaluation Periods System** - Time-based analysis with automated period mapping  
- **🧮 Advanced Calculations** - Handle complex formulas with mixed data types (numeric, "SD", etc.)
- **⚡ Smart Refresh** - Automated workflows that preserve manual entries and custom formulas
- **🔧 Diagnostic Tools** - Built-in troubleshooting and validation utilities
- **🔄 Formula Preservation** - Maintains complex calculations during dataset updates

## 🚀 Quick Start

### Prerequisites
- Google Apps Script project
- Access to source Google Sheets documents
- Basic understanding of Google Sheets formulas

### Installation

1. **Open Google Apps Script**
   ```
   https://script.google.com
   ```

2. **Create New Project or Open Existing**
   
3. **Add the Core Files** (copy and paste each):
   - `Main.js` - Menu system and entry points
   - `CommonHelpers.js` - Shared utility functions  
   - `FormulaCalculator.js` - Calculation engine
   - `DataCombiner.js` - Dataset combination logic
   - `PeriodsManager.js` - Evaluation periods system
   - Plus other supporting files (see [File Structure](#file-structure))

4. **Configure Permissions**
   - Enable external document access
   - Set up proper sharing permissions

5. **Test Installation**
   ```
   Municipality Tools → Test Document Access
   ```

## 📖 Basic Usage

### 1. Generate Datasets
```
Municipality Tools → Dataset Management → Generate Components Dataset
Municipality Tools → Dataset Management → Generate Indicators Dataset  
Municipality Tools → Dataset Management → Combine Components & Indicators
```

### 2. Set Up Evaluation Periods
```
Municipality Tools → Evaluation Periods → Create Evaluation Periods Dataset
Municipality Tools → Evaluation Periods → Configure Smart Refresh
Municipality Tools → Evaluation Periods → Map Periods (Flexible)
```

### 3. Use Calculation Functions
```javascript
// In your spreadsheet cells:
=COMPONENT_VALUE("MRez", "ENE24", "Apodaca")
=COMPONENT_SUM("MRez", "enero 2024", "diciembre 2024", "Apodaca") 
=CALCULATE_INDICATOR("SUM(HTA:OCT24-MAR25)/POP:MAR25*100000", "Monterrey")
```

### 4. Enable Auto-Refresh
```
Municipality Tools → Calculations & Formulas → Setup Auto-Refresh
```

## 🏗️ File Structure

```
📁 Municipality Tools/
├── 🔧 Core System
│   ├── Main.js                    # Menu system & diagnostics
│   └── CommonHelpers.js           # Shared utilities (eliminates duplicates)
│
├── 📊 Data Management  
│   ├── ComponentsManager.js       # Components dataset handling
│   ├── IndicatorsManager.js       # Indicators dataset handling
│   └── DataCombiner.js           # Smart dataset combination
│
├── 🧮 Calculations
│   └── FormulaCalculator.js       # Calculation engine
│
├── 📅 Evaluation System
│   └── PeriodsManager.js          # Consolidated periods management
│
├── 🎨 User Interface
│   ├── UIHelpers.js              # Dialog helpers
│   ├── RefreshManager.js         # Auto-refresh system
│   ├── GenerateComponentsDialog.html
│   └── GenerateIndicatorsDialog.html
│
├── 🛠️ Utilities
│   ├── createNamedRanges.js      # Named range utilities
│   ├── deleteNamedRanges.js      
│   └── Documentation.js          # Technical documentation
│
└── ⚙️ Configuration
    └── appsscript.json           # Project settings
```

## 🔧 Key Functions

### Dataset Management
- **Generate Components Dataset** - Creates municipal components from external sources
- **Smart Refresh** - Updates datasets while preserving user modifications
- **Formula Preservation** - Maintains complex calculations during updates

### Calculation Engine
- **COMPONENT_VALUE()** - Get single component values with string handling
- **COMPONENT_SUM()** - Intelligent summing across date ranges
- **CALCULATE_INDICATOR()** - Process complex formulas with mixed data types

### Evaluation Periods
- **Automated Period Mapping** - Map evaluation dates from external sources
- **Manual Entry Preservation** - Protect user-entered data during updates
- **Smart Refresh Workflows** - Complete automation of period data updates

## 🎯 Use Cases

### Municipal Assessment
- **Performance Indicators** - Track municipal KPIs over time
- **Component Analysis** - Analyze individual assessment components
- **Period Comparisons** - Compare performance across evaluation periods

### Data Management
- **Multi-Source Integration** - Combine data from multiple Google Sheets
- **Automated Workflows** - Set up hands-off data refresh processes  
- **Quality Assurance** - Built-in validation and diagnostic tools

### Calculation Processing
- **Mixed Data Types** - Handle numeric values alongside "SD", "información reservada"
- **Complex Formulas** - Process indicator calculations with component references
- **Batch Operations** - Efficiently process large datasets

## 🛠️ Advanced Features

### Smart Refresh System
```javascript
// Configure once, then automate everything:
Municipality Tools → Evaluation Periods → Configure Smart Refresh
Municipality Tools → Evaluation Periods → Smart Refresh All Periods Data
```

### Diagnostic Tools
```javascript
// Built-in troubleshooting:
Municipality Tools → Testing & Debug → Fast Periods Dataset Diagnostic
Municipality Tools → Testing & Debug → Test External Document Access
```

### Formula Preservation
- Automatically detects formula columns vs data columns
- Preserves complex calculations during dataset refreshes
- Maintains custom column ordering and references

## 📋 Requirements

### Google Sheets Structure
- **Components Sheet**: municipio, id_indicador, eje, tema, nombre_componente, codigo_componente
- **Indicators Sheet**: municipio, id_indicador, eje, tema, nombre_indicador, calculo_indicador
- **BD_componentes**: Core data sheet for calculations

### External Documents
- Proper sharing permissions (Viewer minimum)
- Consistent column naming conventions
- Valid Google Sheets document IDs

## 🐛 Troubleshooting

### Common Issues

**External Document Access**
```
Solution: Check sharing permissions and document IDs
Tool: Municipality Tools → Testing & Debug → Test External Document Access
```

**Column Mapping Errors**
```
Solution: Verify source sheet column names
Tool: Municipality Tools → Dataset Management → Preview Column Mapping
```

**Performance Issues**
```
Solution: Use batch processing, avoid frequent manual refreshes
Tool: Built-in batch processing (50 rows default)
```

### Diagnostic Tools
- **Fast Diagnostic**: Quick analysis of first 50 rows
- **Detailed Analysis**: Comprehensive data structure review
- **Debug Mapping**: Step-by-step mapping logic verification

## 🔄 Version History

### v2.0.0 (Current) - Major Restructuring
- ✅ Eliminated all duplicate functions across files
- ✅ Consolidated evaluation periods system (4 files → 1)
- ✅ Enhanced formula preservation during refreshes
- ✅ Improved error handling and diagnostics
- ✅ Added smart refresh workflows
- ✅ Optimized performance with batch processing

### v1.2.0 - Multi-file Organization
- Restructured code into multiple files
- Improved maintainability

### v1.1.0 - Enhanced Reliability  
- Added component code support
- Improved refresh reliability

### v1.0.0 - Initial Release
- Basic dataset generation and calculations

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Test thoroughly with diagnostic tools
4. Update documentation
5. Commit changes (`git commit -m 'Add AmazingFeature'`)
6. Push to branch (`git push origin feature/AmazingFeature`)
7. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- **Documentation**: Check `Documentation.js` for technical details
- **Diagnostic Tools**: Use built-in testing and debug functions
- **Console Logs**: Check browser console for detailed error information
- **Issue Tracker**: Report bugs and request features via GitHub issues

## 🙏 Acknowledgments

- Google Apps Script platform for serverless automation
- Municipal assessment teams for requirements and testing
- Open source community for inspiration and best practices

---

⭐ **Star this repository if it helps with your municipal data management!**