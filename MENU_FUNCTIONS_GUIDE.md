# 💰 Salary Ranges Calculator - Menu Functions Guide

**Version**: 3.3.0  
**Date**: 2025-11-27

Complete description of every function available in the menu system.

---

## 📋 Menu Structure

```
💰 Salary Ranges Calculator
├── ⚙️ Setup
├── 📥 Import Data
├── 🏗️ Build
├── 📤 Export
└── 🔧 Tools
```

---

## ⚙️ SETUP MENU

### ⚡ Quick Setup (Run Once)
**Function**: `quickSetup_()`

**Purpose**: One-click initialization of the entire system

**What it does (6 steps)**:
1. Creates all necessary tabs (Aon region tabs, mapping tabs)
2. Seeds executive job family mappings from Aon data
3. Fills job families in Aon region tabs
4. Builds the interactive calculator UI with dropdowns
5. Generates the Help & Instructions sheet
6. Enhances mapping sheets with formatting and validations

**When to use**: 
- ✅ First-time setup after pasting Aon data
- ✅ Complete reset of the system

**Prerequisites**:
- Aon data must be pasted into region tabs first
- Aon tabs must have: Job Code, Job Family, and percentile columns

**Time**: ~30-60 seconds

**Output**: Alert with next steps (configure HiBob API, import data)

---

### 📖 Generate Help Sheet
**Function**: `buildHelpSheet_()`

**Purpose**: Creates comprehensive help documentation in the spreadsheet

**What it does**:
- Creates "About & Help" sheet
- Adds step-by-step instructions for:
  - Quick Start workflow
  - Manual setup workflow
  - Regular usage workflow
  - Menu function descriptions
  - Calculation explanations
  - Mapping instructions
  - Tips and troubleshooting

**When to use**:
- ✅ Need in-spreadsheet documentation
- ✅ Training new users
- ✅ Quick reference guide

**Time**: <5 seconds

**Output**: Creates/updates "About & Help" sheet

---

### 🌍 Create Aon Region Tabs
**Function**: `createAonPlaceholderSheets_()`

**Purpose**: Creates placeholder sheets for Aon market data

**What it does**:
- Creates 3 sheets (if they don't exist):
  - "Aon India - 2025"
  - "Aon US - 2025"
  - "Aon UK - 2025"
- Adds header row with proper column names:
  - Job Code
  - Job Family
  - Market (43) CFY Fixed Pay: 10th Percentile
  - Market (43) CFY Fixed Pay: 25th Percentile
  - Market (43) CFY Fixed Pay: 50th Percentile
  - Market (43) CFY Fixed Pay: 62.5th Percentile
  - Market (43) CFY Fixed Pay: 75th Percentile
  - Market (43) CFY Fixed Pay: 90th Percentile
- Formats headers as bold
- Sets up number formatting for percentile columns
- Freezes header row

**When to use**:
- ✅ Initial setup
- ✅ Before pasting Aon data
- ✅ If tabs were accidentally deleted

**Time**: <5 seconds

**Output**: Toast notification confirming tab creation

---

### 🗺️ Create Mapping Tabs
**Function**: `createMappingPlaceholderSheets_()`

**Purpose**: Creates all required mapping sheets with proper structure

**What it does**:
- Creates/ensures 4 mapping sheets exist:

1. **Title Mapping**
   - Headers: Job title (live), Job title (Mapped), Job family
   - Purpose: Map job titles from HiBob to job families

2. **Job family Descriptions**
   - Headers: Aon Code, Job Family (Exec Description)
   - Purpose: Map Aon codes to user-friendly descriptions

3. **Employee Level Mapping**
   - Headers: Emp ID, Mapping, Status
   - Purpose: Map employees to specific job levels/families

4. **Aon Code Remap**
   - Headers: From Code, To Code
   - Default: EN.SOML → EN.AIML
   - Purpose: Handle Aon vendor code changes

**When to use**:
- ✅ Initial setup
- ✅ If mapping sheets are missing
- ✅ To reset mapping structure

**Time**: <5 seconds

**Output**: Toast notification confirming creation

---

### 📊 Build Calculator UI
**Function**: `buildCalculatorUI_()`

**Purpose**: Creates the interactive salary range calculator interface

**What it does**:
1. Creates/updates "Salary Ranges" sheet
2. Sets up control panel:
   - Row 2: Job Family dropdown (from mappings)
   - Row 3: Category dropdown (X0 or Y1)
   - Row 4: Region dropdown (US, UK, India)
3. Creates data table with headers:
   - Column A: Level (L2 IC through L9 Mgr)
   - Columns B-D: Range Start, Range Mid, Range End (market data)
   - Columns F-H: Min, Median, Max (internal data)
   - Column L: Emp Count (internal data)
4. Inserts formulas for all 16 levels:
   - Market ranges: SALARY_RANGE_MIN/MID/MAX functions
   - Internal stats: INTERNAL_STATS function
5. Applies currency formatting based on region

**When to use**:
- ✅ After creating mapping tabs
- ✅ After seeding job family mappings
- ✅ To rebuild the calculator interface
- ✅ If formulas are broken

**Time**: ~1 second (optimized with batch operations)

**Output**: 
- Interactive calculator sheet ready to use
- Toast notification on completion

---

### 🔧 Manage Exec Mappings
**Function**: `openExecMappingManager_()`

**Purpose**: Opens web interface for managing job family mappings

**What it does**:
- Opens sidebar with HTML interface
- Allows you to:
  - View all existing Aon Code → Exec Description mappings
  - Add new mappings
  - Edit existing mappings
  - Delete mappings
- Changes are immediately saved to "Job family Descriptions" sheet
- Clears cache after updates

**When to use**:
- ✅ Need to add custom job family descriptions
- ✅ Fix incorrect mappings
- ✅ Manage mappings with UI (easier than direct sheet editing)

**Time**: Instant (opens sidebar)

**Output**: Interactive sidebar panel

---

### ✅ Ensure Category Picker
**Function**: `ensureCategoryPicker_()`

**Purpose**: Ensures category dropdown is properly configured

**What it does**:
- Checks cell B3 in "Salary Ranges" sheet
- Creates/updates dropdown validation to show only X0 and Y1
- Sets default value to X0 if blank
- Converts old X1 values to X0

**When to use**:
- ✅ Category dropdown not working
- ✅ After upgrading from 3-category system
- ✅ Dropdown shows wrong values

**Time**: <1 second

**Output**: Category dropdown properly configured

---

### 🎨 Enhance Mapping Sheets
**Function**: `enhanceMappingSheets_()`

**Purpose**: Adds visual indicators and formulas to mapping sheets

**What it does**:

**For Employee Level Mapping**:
- Adds "Status" column with formula: Shows "Missing" when mapping is blank
- Adds "Missing Count" cell showing total unmapped employees
- Adds conditional formatting: Highlights unmapped rows in red
- Formula: `=IF(LEN(Mapping)=0,"Missing","")`

**For Title Mapping**:
- Adds "Status" column: Shows "Missing" for unmapped titles
- Adds "Missing Count" cell
- Red highlighting for unmapped job titles
- Formula: `=IF(LEN(Job family)=0,"Missing","")`

**When to use**:
- ✅ After syncing mappings from Bob
- ✅ To visualize which mappings are incomplete
- ✅ To see missing count at a glance

**Time**: <5 seconds

**Output**: Mapping sheets with visual indicators and counts

---

## 📥 IMPORT DATA MENU

### 🔄 Import All Bob Data
**Function**: `importAllBobData()`

**Purpose**: Imports all employee data from HiBob API in one action

**What it does (3 imports in sequence)**:
1. **Base Data** - Employee information
   - Employee ID, Name, Job Level, Job Title
   - Base salary, Employment Type
   - Site/Location, Start Date
   - Job Family Name, Active/Inactive status
   - Creates "Base Data" sheet

2. **Bonus History** - Latest bonus/commission per employee
   - Employee ID, Name
   - Effective date, Variable type
   - Commission/Bonus percentage, Amount, Currency
   - Creates "Bonus History" sheet

3. **Compensation History** - Latest comp change per employee
   - Employee ID, Name
   - Effective date, Base salary
   - Currency, Change reason
   - Creates "Comp History" sheet

**When to use**:
- ✅ Initial setup after configuring HiBob API credentials
- ✅ Daily/weekly data refresh
- ✅ Before rebuilding Full List
- ✅ After employees are hired/terminated

**Prerequisites**:
- HiBob API credentials configured in Script Properties:
  - BOB_ID
  - BOB_KEY

**Time**: 1-3 minutes (depending on employee count)

**Output**: 
- Creates/updates 3 sheets
- Success alert on completion
- Error alert if API fails

---

### 👥 Import Base Data Only
**Function**: `importBobDataSimpleWithLookup()`

**Purpose**: Imports only base employee data (faster)

**What it does**:
- Fetches Base Data report (ID: 31048356) from HiBob
- Filters for "Permanent" and "Regular Full-Time" employees only
- Formats Employee ID as text (for XLOOKUP compatibility)
- Formats Base Pay as currency
- Auto-resizes columns

**When to use**:
- ✅ Only need to refresh employee list
- ✅ Don't need bonus/comp history
- ✅ Faster than full import

**Time**: 30-60 seconds

**Output**: Updates "Base Data" sheet

---

### 💰 Import Bonus Only
**Function**: `importBobBonusHistoryLatest()`

**Purpose**: Imports only bonus/commission history

**What it does**:
- Fetches Bonus History report (ID: 31054302)
- Keeps only LATEST entry per employee
- Includes variable type, percentage, amount
- Formats dates, percentages, and amounts

**When to use**:
- ✅ Bonus structure changed
- ✅ Need to update Variable Type/% in Base Data
- ✅ Don't need full data refresh

**Time**: 30-60 seconds

**Output**: Updates "Bonus History" sheet

---

### 📈 Import Comp History Only
**Function**: `importBobCompHistoryLatest()`

**Purpose**: Imports only compensation change history

**What it does**:
- Fetches Compensation History report (ID: 31054312)
- Keeps only LATEST entry per employee
- Includes effective date, salary, currency, reason
- Useful for tracking last compensation change

**When to use**:
- ✅ Need comp change dates
- ✅ Analyzing compensation trends
- ✅ Don't need full data refresh

**Time**: 30-60 seconds

**Output**: Updates "Comp History" sheet

---

## 🏗️ BUILD MENU

### 📊 Rebuild Full List (with validation)
**Function**: `rebuildFullListTabsWithValidation_()`

**Purpose**: Generates comprehensive salary ranges combining Aon + Internal data

**What it does**:

**Step 1: Validation**
- Checks all Aon region tabs exist and have data
- Checks Lookup and Job family Descriptions exist
- Checks HiBob credentials configured
- Shows error if prerequisites missing

**Step 2: Build Full List**
- Reads Lookup table (CIQ Level → Aon Level mapping)
- Reads all 3 Aon region sheets
- Reads Base Data for internal statistics
- For each combination of:
  - Site (India, US, UK)
  - Region (India, US, UK)
  - Aon Code (EN.SODE, FI.FINA, etc.)
  - Job Family (Exec Description)
  - CIQ Level (L2 IC through L9 Mgr)
- Calculates:
  - Market percentiles: P10, P25, P40, P50, P62.5, P75, P90
  - Internal stats: Min, Median, Max, Employee Count
- Handles half-levels (L5.5, L6.5) by averaging neighbors

**Step 3: Create Coverage Summary** (if Base Data exists)
- Shows which job families have market data
- Shows which have internal data
- Counts levels with data vs expected

**Step 4: Create Employees (Mapped)** (if Base Data exists)
- Lists all mapped employees
- Shows their Aon Code, Level, Site, Salary
- Audit trail for employee mapping

**When to use**:
- ✅ After importing Bob data
- ✅ After updating Aon data
- ✅ After changing mappings
- ✅ Before using calculator
- ✅ Regular refresh (weekly/monthly)

**Prerequisites**:
- ✅ Aon region tabs with data
- ✅ Lookup table configured
- ✅ Job family Descriptions populated
- ✅ (Optional) Base Data for internal stats

**Time**: 30-90 seconds (depends on data size)

**Output**:
- Creates/updates "Full List" sheet (all combinations)
- Creates/updates "Coverage Summary" sheet
- Creates/updates "Employees (Mapped)" sheet
- Toast notification on completion

**Key Column in Full List**:
- "Key" column = "{ExecDescription}{CIQLevel}{Region}"
- Used for fast O(1) lookups by SALARY_RANGE functions

---

### 💵 Build Full List USD
**Function**: `buildFullListUsd_()`

**Purpose**: Creates USD-converted view of Full List for multi-region analysis

**What it does**:
- Reads "Full List" sheet
- Reads FX rates from "Lookup" sheet (Region → FX columns)
- For each row:
  - Gets region and FX rate
  - Multiplies all percentiles by FX rate:
    - P10, P25, P40, P50, P62.5, P75, P90
  - Multiplies internal stats by FX rate:
    - Internal Min, Median, Max
  - Rounds market percentiles to nearest $100
- Writes to "Full List USD" sheet

**When to use**:
- ✅ Comparing salaries across regions
- ✅ Global compensation analysis
- ✅ Need everything in one currency (USD)

**Prerequisites**:
- Full List must exist
- Lookup sheet must have FX rates configured

**FX Rates** (from Lookup sheet):
- US: 1.0 (base currency)
- UK: ~1.37 (GBP to USD)
- India: ~0.0125 (INR to USD)

**Time**: 5-15 seconds

**Output**: Creates/updates "Full List USD" sheet

---

### 🌱 Seed All Job Family Mappings
**Function**: `seedAllJobFamilyMappings_()`

**Purpose**: Automatically populates job family mappings from Aon data

**What it does (2 sub-functions)**:

**Part 1: Seed Exec Mappings**
- Scans all 3 Aon region sheets
- Extracts unique Aon Codes (e.g., EN.SODE, FI.FINA)
- Extracts Job Family names from Aon data
- Adds new mappings to "Job family Descriptions" sheet
- Skips codes that already exist

**Part 2: Fill Region Families**
- Scans Aon region sheets
- For each Job Code, fills in Job Family column
- Uses mappings from "Job family Descriptions"
- Highlights missing mappings in red
- Shows count of filled vs missing

**When to use**:
- ✅ First-time setup
- ✅ After adding new Aon data
- ✅ After Aon releases updated market data
- ✅ New job families added to organization

**Time**: 10-20 seconds

**Output**: 
- Populated "Job family Descriptions" sheet
- Filled Job Family columns in Aon sheets
- Toast with count of filled/missing

---

### 👥 Sync All Bob Mappings
**Function**: `syncAllBobMappings_()`

**Purpose**: Syncs employee-related mappings from HiBob data

**What it does (2 sub-functions)**:

**Part 1: Sync Employee Level Mapping**
- Reads "Base Data" sheet
- Extracts all active employees (Employee ID)
- For each employee:
  - Preserves existing mappings
  - Adds suggestion from Title Mapping (if available)
  - Adds to "Employee Level Mapping" sheet
- Adds Status column formula: Shows "Missing" for unmapped
- Adds Missing Count
- Red highlighting for blank mappings

**Part 2: Sync Title Mapping**
- Extracts all unique job titles from Base Data
- Adds new titles to "Title Mapping" sheet
- Preserves existing mappings
- Adds Status and Missing Count columns
- Red highlighting for unmapped titles

**When to use**:
- ✅ After importing Bob data
- ✅ New employees hired
- ✅ New job titles created
- ✅ Before rebuilding Full List

**Prerequisites**:
- Base Data must be imported

**Time**: 10-30 seconds

**Output**:
- Updated "Employee Level Mapping" sheet
- Updated "Title Mapping" sheet
- Toast with count of synced items

---

### 🗑️ Clear All Caches
**Function**: `clearAllCaches_()`

**Purpose**: Clears all cached data to force fresh calculations

**What it does**:
- Clears CacheService (Document Cache):
  - Internal stats cache (INT:*)
  - Aon value cache (AON:*)
  - Sheet data cache (SHEET_DATA:*)
  - Exec description map cache
  - Code remap cache
- Clears in-memory caches:
  - Lookup map cache
  - Header cache
- Forces all functions to re-read data on next call

**When to use**:
- ✅ Data looks stale or incorrect
- ✅ After major data updates
- ✅ Calculator showing old values
- ✅ Troubleshooting calculation issues
- ✅ After changing mappings

**Time**: <1 second

**Output**: Toast notification "All caches cleared"

**Note**: Caches automatically expire after 10 minutes (CACHE_TTL)

---

## 📤 EXPORT MENU

### 💼 Export Proposed Ranges
**Function**: `exportProposedSalaryRanges_()`

**Purpose**: Export salary range recommendations

**What it does**:
- Prompts for category (X0 or Y1)
- Rebuilds Full List to ensure fresh data
- Points user to Full List sheet for calculations
- (Note: Actual export to separate file trimmed in current version)

**When to use**:
- ✅ Need to share salary ranges with stakeholders
- ✅ Creating salary proposals
- ✅ Compensation planning

**Time**: Same as Rebuild Full List (~30-90 seconds)

**Output**: Alert to use Full List sheet

**Recommendation**: Use Full List sheet directly for most needs

---

## 🔧 TOOLS MENU

### 💱 Apply Currency Format
**Function**: `applyCurrency_()`

**Purpose**: Applies region-appropriate currency formatting to active sheet

**What it does**:
1. Detects region from sheet (reads Region cell or dropdown)
2. Detects currency from sheet (if explicitly labeled)
3. Selects appropriate format:
   - **India**: ₹#,##,##0 (Indian Rupee with lakhs/crores)
   - **US**: $#,##0 (US Dollar)
   - **UK**: £#,##0 (British Pound)
4. Finds header row (searches for Level, P62.5, P75 columns)
5. Applies currency format to:
   - Market range columns (Range Start, Mid, End)
   - Internal stats columns (Min, Median, Max)
6. Applies count format to Emp Count column
7. Uses special format to hide zeros: `$#,##0;$#,##0;;@`

**When to use**:
- ✅ Numbers showing without currency symbols
- ✅ After region change
- ✅ Format looks wrong
- ✅ Manual formatting needed

**Time**: <2 seconds

**Output**: Currency symbols appear on numbers

**Note**: Also called automatically by buildCalculatorUI_()

---

### ℹ️ Instructions & Help
**Function**: `showInstructions()`

**Purpose**: Shows quick-start instructions in a modal dialog

**What it does**:
- Opens modal dialog (600x600px)
- Displays:
  - First-time setup steps
  - Regular workflow
  - Custom function examples
  - Category definitions
  - Link to Aon data source
- HTML formatted for readability

**When to use**:
- ✅ New user onboarding
- ✅ Quick reference
- ✅ Forgot workflow steps
- ✅ Need function syntax

**Time**: Instant

**Output**: Modal dialog with instructions

**Tip**: For detailed help, use "Generate Help Sheet" instead

---

## 📊 HIDDEN/SUPPORT FUNCTIONS

These functions are called by menu functions but not directly exposed:

### buildTitleToFamilyMap_()
- Builds mapping from job titles to job families
- Used by employee level sync for suggestions

### _buildInternalIndex_()
- Pre-processes Base Data into indexed structure
- Used by Full List rebuild for internal stats
- Caches results for 10 minutes

### _getExecDescMap_()
- Loads Aon Code → Exec Description mappings
- Caches for 10 minutes
- Used throughout for friendly names

### getLookupMap_()
- Loads CIQ Level → Aon Level mappings
- Caches for 10 minutes
- Critical for level translations

### getRegionSheet_()
- Resolves region name to actual sheet
- Handles fallbacks and aliases
- Used by all Aon lookup functions

---

## 🎯 RECOMMENDED WORKFLOWS

### Initial Setup (New Sheet)
```
1. Setup → ⚡ Quick Setup (Run Once)
2. Configure HiBob API credentials (Script Properties)
3. Import Data → 🔄 Import All Bob Data
4. Build → 👥 Sync All Bob Mappings
5. Build → 📊 Rebuild Full List (with validation)
6. Start using calculator!
```

### Regular Refresh (Weekly/Monthly)
```
1. Import Data → 🔄 Import All Bob Data
2. Build → 📊 Rebuild Full List (with validation)
3. (Optional) Build → 💵 Build Full List USD
```

### Adding New Job Families
```
1. Paste new Aon data into region tabs
2. Build → 🌱 Seed All Job Family Mappings
3. Setup → 🔧 Manage Exec Mappings (customize descriptions)
4. Build → 📊 Rebuild Full List (with validation)
```

### Troubleshooting
```
1. Build → 🗑️ Clear All Caches
2. Import Data → 🔄 Import All Bob Data (refresh data)
3. Build → 📊 Rebuild Full List (with validation)
4. Setup → 🎨 Enhance Mapping Sheets (check for missing)
```

---

## 🔐 Script Properties Required

Set these in: **Extensions → Apps Script → Project Settings → Script Properties**

| Property | Purpose | Example |
|----------|---------|---------|
| `BOB_ID` | HiBob API Service Account ID | your_bob_id |
| `BOB_KEY` | HiBob API Service Account Key | your_secret_key |

**How to get HiBob credentials**:
1. Login to HiBob
2. Go to Settings → API → Service Users
3. Create/view service user
4. Copy ID and Key

---

## 📚 Custom Functions (For Formulas)

These can be used directly in Google Sheets cells:

### Salary Range Functions
```javascript
=SALARY_RANGE(category, region, family, level)
=SALARY_RANGE_MIN(category, region, family, level)
=SALARY_RANGE_MID(category, region, family, level)
=SALARY_RANGE_MAX(category, region, family, level)

// UI versions (read from calculator dropdowns):
=UI_SALARY_RANGE(region, family, level)
=UI_SALARY_RANGE_MIN(region, family, level)
```

### Aon Percentile Functions
```javascript
=AON_P10(region, family, level)   // 10th percentile
=AON_P25(region, family, level)   // 25th percentile
=AON_P40(region, family, level)   // 40th percentile
=AON_P50(region, family, level)   // 50th percentile (median)
=AON_P625(region, family, level)  // 62.5th percentile
=AON_P75(region, family, level)   // 75th percentile
=AON_P90(region, family, level)   // 90th percentile
```

### Internal Stats Function
```javascript
=INTERNAL_STATS(region, family, level)
// Returns: [Min, Median, Max, Employee Count]

// To use individual values:
=INDEX(INTERNAL_STATS("US", "EN.SODE", "L5 IC"), 1, 1)  // Min
=INDEX(INTERNAL_STATS("US", "EN.SODE", "L5 IC"), 1, 2)  // Median
=INDEX(INTERNAL_STATS("US", "EN.SODE", "L5 IC"), 1, 3)  // Max
=INDEX(INTERNAL_STATS("US", "EN.SODE", "L5 IC"), 1, 4)  // Count
```

---

## ⚡ Performance Notes

All functions are optimized with caching:
- **Cache TTL**: 10 minutes
- **Sheet data**: Cached to reduce reads
- **Lookups**: Indexed for O(1) access
- **Batch operations**: Used throughout

**Typical Execution Times**:
- Quick functions (<5s): Setup tabs, clear cache, help
- Medium functions (10-30s): Seed mappings, sync Bob
- Long functions (30-90s): Import Bob data, rebuild Full List

---

## 🐛 Common Issues & Solutions

### "Prerequisites Missing" Error
**Solution**: Run Setup → ⚡ Quick Setup (Run Once)

### "Sheet not found" Error
**Solution**: Setup → 🌍 Create Aon Region Tabs

### "Missing BOB_ID or BOB_KEY" Error
**Solution**: Configure Script Properties (see above)

### Calculator Shows Old Data
**Solution**: Build → 🗑️ Clear All Caches

### Missing Mappings Highlighted in Red
**Solution**: 
1. Setup → 🔧 Manage Exec Mappings (for job families)
2. Manually fill mapping sheets
3. Build → 📊 Rebuild Full List

### Formula Returns Blank
**Solution**:
1. Check if Full List has data for that combination
2. Check Coverage Summary to see data availability
3. Rebuild Full List if needed

---

## 📞 Support

For detailed help:
- **In-spreadsheet**: Setup → 📖 Generate Help Sheet
- **Quick start**: Tools → ℹ️ Instructions & Help
- **This guide**: MENU_FUNCTIONS_GUIDE.md

---

**Last Updated**: 2025-11-27  
**Version**: 3.3.0

