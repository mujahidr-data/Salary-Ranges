# 🎯 Simplified Workflow Plan

## Current State Review

### Sheets Currently Created:
1. **Aon India - 2025** ✅ KEEP - Market data source
2. **Aon US - 2025** ✅ KEEP - Market data source
3. **Aon UK - 2025** ✅ KEEP - Market data source
4. **Base Data** ✅ KEEP - Employee data from HiBob
5. **Bonus History** ❓ OPTIONAL - Can be trimmed if not needed
6. **Comp History** ❓ OPTIONAL - Can be trimmed if not needed
7. **Lookup** ✅ KEEP - CIQ Level → Aon Level mapping + FX rates
8. **Job family Descriptions** ✅ KEEP - Aon Code → Job Family mapping
9. **Employee Level Mapping** ✅ KEEP - Employee → Level/Family mapping
10. **Title Mapping** ✅ KEEP - Job Title → Job Family mapping
11. **Aon Code Remap** ✅ KEEP - Code change handling (EN.SOML → EN.AIML)
12. **Salary Ranges** ✅ KEEP - X0 Calculator (Engineering/Product)
13. **Salary Ranges (Y1)** ✅ KEEP - Y1 Calculator (Everyone Else)
14. **Full List** ✅ KEEP - Consolidated market data (local currency)
15. **Full List USD** ✅ KEEP - Consolidated market data (USD)
16. **Coverage Summary** ❌ REMOVE - Not needed per user
17. **Employees (Mapped)** ❌ REMOVE - Not needed per user
18. **About & Help** ❓ OPTIONAL - Documentation sheet

---

## 🎯 New Simplified Menu Structure

```
💰 Salary Ranges Calculator
├── 🏗️ Fresh Build (Create All Sheets)
├── 📥 Import Bob Data
├── 📊 Build Market Data (Full Lists)
└── 🔧 Tools
    ├── 💱 Apply Currency Format
    ├── 🗑️ Clear Caches
    └── ℹ️ Help
```

---

## ✨ Function 1: Fresh Build

**Name**: `freshBuild()`

**Purpose**: One-click setup of all required sheets and structure

**What it creates**:

### Data Source Sheets:
1. **Aon India - 2025** - Headers + formatting
2. **Aon US - 2025** - Headers + formatting  
3. **Aon UK - 2025** - Headers + formatting

### Mapping Sheets:
4. **Lookup** - Level mapping + FX rates with example data
5. **Job family Descriptions** - Aon Code → Family mapping
6. **Employee Level Mapping** - Employee → Level mapping
7. **Title Mapping** - Job Title → Family mapping
8. **Aon Code Remap** - Code change mapping (default: EN.SOML → EN.AIML)

### Calculator Sheets:
9. **Salary Ranges (X0)** - Engineering/Product calculator
10. **Salary Ranges (Y1)** - Everyone Else calculator

### Output Sheets (Placeholders):
11. **Full List** - Headers only (built by Function 3)
12. **Full List USD** - Headers only (built by Function 3)

**Steps**:
1. Create Aon region tabs with proper headers
2. Create mapping tabs with structure
3. Create Lookup tab with level mapping + FX rates
4. Create both calculator UIs (X0 and Y1)
5. Create Full List placeholder sheets
6. Show success message with next steps

**Time**: ~10 seconds

---

## 📥 Function 2: Import Bob Data

**Name**: `importBobData()`

**Purpose**: Import employee data from HiBob

**What it imports**:

### Required:
- **Base Data** - Employee list with:
  - Employee ID, Name, Job Level, Job Title
  - Base Salary, Employment Type
  - Site, Start Date, Job Family, Active status

### Optional (can be included or removed):
- **Bonus History** - Latest bonus/commission per employee
- **Comp History** - Latest compensation change per employee

**After Import**:
- Auto-syncs Employee Level Mapping (adds new employees)
- Auto-syncs Title Mapping (adds new job titles)
- Shows count of new employees/titles added

**Time**: 1-2 minutes

---

## 📊 Function 3: Build Market Data

**Name**: `buildMarketData()`

**Purpose**: Generate Full List and Full List USD from Aon data

**What it does**:

### Step 1: Validation
- Checks Aon sheets have data
- Checks Lookup table exists
- Checks Job family Descriptions populated
- Shows error if prerequisites missing

### Step 2: Build Full List
- Reads all Aon sheets (India, US, UK)
- Reads Employee Level Mapping
- Reads Base Data
- **Only includes combinations that have actual employees**
- For each employee:
  - Maps to Aon Code via Job family Descriptions
  - Maps to Level via Employee Level Mapping
  - Gets Region from Site
  - Pulls market percentiles (P10, P25, P40, P50, P62.5, P75, P90)
  - Calculates internal stats (Min, Median, Max, Count)
- Outputs to "Full List" sheet

### Step 3: Build Full List USD
- Reads Full List
- Reads FX rates from Lookup
- Converts all values to USD
- Outputs to "Full List USD" sheet

**Time**: 30-90 seconds (depends on employee count)

**Key Difference**: 
- ❌ OLD: Generated ALL possible combinations of region/family/level
- ✅ NEW: Only generates combinations for ACTUAL employees

This makes Full List much smaller and focused!

---

## 🔧 Tools Menu (Simplified)

### 💱 Apply Currency Format
- Current function (no change)

### 🗑️ Clear Caches
- Current function (no change)

### ℹ️ Help
- Show quick instructions modal

---

## ❌ Functions to Remove

1. ~~Quick Setup~~ → Replaced by Fresh Build
2. ~~Generate Help Sheet~~ → Optional, can keep if wanted
3. ~~Create Aon Region Tabs~~ → Part of Fresh Build
4. ~~Create Mapping Tabs~~ → Part of Fresh Build
5. ~~Build Calculator UI~~ → Part of Fresh Build
6. ~~Manage Exec Mappings~~ → Can keep in Tools if useful
7. ~~Ensure Category Picker~~ → Auto-handled by Fresh Build
8. ~~Enhance Mapping Sheets~~ → Auto-handled by Import Bob Data
9. ~~Import All Bob Data~~ → Renamed to "Import Bob Data"
10. ~~Import Base Data Only~~ → Removed (use main import)
11. ~~Import Bonus Only~~ → Removed (use main import)
12. ~~Import Comp History Only~~ → Removed (use main import)
13. ~~Rebuild Full List~~ → Renamed to "Build Market Data"
14. ~~Build Full List USD~~ → Part of Build Market Data
15. ~~Seed All Job Family Mappings~~ → Auto-handled by Fresh Build
16. ~~Sync All Bob Mappings~~ → Auto-handled by Import Bob Data
17. ~~Export Proposed Ranges~~ → Removed (not needed)

---

## 📋 Sheets to Remove

### During Rebuild Process:
1. **Coverage Summary** - Remove creation code
2. **Employees (Mapped)** - Remove creation code

These can be manually deleted from existing sheets.

---

## 🎯 Recommended New Workflow

### First Time Setup:
```
1. Run: 🏗️ Fresh Build
2. Paste Aon data into region tabs (India, US, UK)
3. Configure HiBob API credentials (Script Properties)
4. Run: 📥 Import Bob Data
5. Map employees in "Employee Level Mapping" sheet
6. Run: 📊 Build Market Data
7. Use calculators!
```

### Regular Refresh (Weekly/Monthly):
```
1. Run: 📥 Import Bob Data (get latest employees)
2. Update any new employee mappings
3. Run: 📊 Build Market Data (rebuild lists)
```

### After Aon Data Update:
```
1. Paste new Aon data into region tabs
2. Run: 📊 Build Market Data
```

---

## ✅ Implementation Checklist

- [ ] Create `freshBuild()` function
- [ ] Simplify `importBobData()` function  
- [ ] Modify `buildMarketData()` to only include actual employees
- [ ] Remove Coverage Summary creation code
- [ ] Remove Employees (Mapped) creation code
- [ ] Update `onOpen()` menu with new structure
- [ ] Remove/hide old functions
- [ ] Update calculator UI creation for both X0 and Y1
- [ ] Test full workflow

---

## 🤔 Questions for User

1. **Bonus/Comp History**: Do you want to keep importing these or remove them?
2. **Help Sheet**: Keep the "About & Help" sheet generation or remove it?
3. **Exec Mappings Manager**: Keep the web UI for managing mappings or remove?
4. **Full List Scope**: Confirm you want ONLY combinations for actual employees (not all possible combinations)?

---

**Next Steps**: Once approved, I'll implement this simplified workflow!

