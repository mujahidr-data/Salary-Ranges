# Changelog

## [3.0.0] - 2025-11-13

### 🎉 Major Release: Consolidated Script

#### Added
- **Consolidated Script**: Combined all functionality into single `SalaryRangesCalculator.gs` file (~1900 lines)
- **Comprehensive Menu System**: Organized into Setup, Import, Build, Export, and Tools submenus
- **Improved Error Handling**: Better error messages and validation throughout
- **Enhanced Documentation**: Updated README, QUICKSTART, and SETUP guides

#### Changed
- **Architecture**: Merged `AppImports.gs`, `Helpers.gs`, and `RangeCalculator.gs` into one file
- **Menu Structure**: Reorganized menu with emoji icons and logical grouping
- **clasp Configuration**: Updated to only push consolidated script
- **Documentation**: Completely rewritten to reflect new structure

#### Improved
- **Code Organization**: Better function grouping and comments
- **Performance**: Optimized caching and sheet read operations
- **Maintainability**: Single file easier to manage and deploy
- **User Experience**: Clearer menu options and help dialogs

#### Archived
- Moved original individual scripts to `archive/` folder:
  - `AppImports.gs`
  - `Helpers.gs`
  - `RangeCalculator.gs`

### 📁 Project Structure Changes

**Before (v2.x)**:
```
├── AppImports.gs
├── Helpers.gs
├── RangeCalculator.gs
└── ExecMappingManager.html
```

**After (v3.0)**:
```
├── SalaryRangesCalculator.gs  # ⭐ All-in-one script
├── ExecMappingManager.html
└── archive/                   # Old scripts (reference)
    ├── AppImports.gs
    ├── Helpers.gs
    └── RangeCalculator.gs
```

### 🔧 Technical Details

- **Lines of Code**: ~1900 lines in consolidated script
- **Functions**: ~80+ functions organized by purpose
- **Menu Items**: 25+ menu options across 5 submenus
- **Caching**: 10-minute TTL for performance
- **API Integration**: HiBob API v1

### 🚀 Deployment

- Only `SalaryRangesCalculator.gs` and `ExecMappingManager.html` are pushed to Apps Script
- Old scripts archived but not deployed
- Faster deployment with single script

### 📊 Features Preserved

- ✅ All Bob data import functionality
- ✅ All Aon market data calculations
- ✅ All salary range formulas
- ✅ All mapping and configuration tools
- ✅ All helper functions
- ✅ Interactive calculator UI
- ✅ Multi-region support
- ✅ FX conversion
- ✅ Internal vs Market analytics

### 🔗 Data Source

Aon market data location documented:
https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-

---

## [2.0.0] - 2024-2025

### Previous Version
- Separate scripts for imports, helpers, and calculations
- Basic menu system
- Core functionality established

---

## Migration Guide (v2.x → v3.0)

### For Existing Users

1. **Pull the latest code**:
   ```bash
   git pull origin main
   ```

2. **Push the consolidated script**:
   ```bash
   npm run push
   ```

3. **No data migration needed**:
   - All your data and mappings remain intact
   - Sheet structure unchanged
   - Custom functions work identically

4. **New menu structure**:
   - Refresh your Google Sheet to see new menu
   - All functions available in reorganized menu

### Breaking Changes

- ❌ None! Fully backward compatible
- ✅ All custom functions preserved
- ✅ All sheet names unchanged
- ✅ All data structures intact

---

**Current Version**: 3.0.0  
**Status**: Stable  
**Last Updated**: November 13, 2025

