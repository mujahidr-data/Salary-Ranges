# 💰 Salary Ranges Calculator

**v3.1.0** - Consolidated Google Apps Script for comprehensive salary range analysis

Combines HiBob employee data with Aon market data to calculate salary ranges across multiple regions, job families, and career levels.

**🆕 What's New in v3.1.0:**
- ✨ P10 and P25 percentile support
- ⚡ Quick Setup (one-click initialization)
- 🎯 Simplified menu (combined functions)
- ✅ Prerequisite validation

## 🎯 Features

### Data Integration
- ✅ **HiBob API**: Auto-import employee data, bonus, and compensation history
- ✅ **Aon Market Data**: **P10, P25, P40, P50, P62.5, P75, P90** percentiles
- ✅ **Multi-Region**: US, UK, India with FX conversion
- ✅ **Smart Mapping**: Job families, titles, and employee levels

### Salary Range Categories
- **X0**: P62.5 (min) / P75 (mid) / P90 (max) - *Top of market*
- **X1**: P50 (min) / P62.5 (mid) / P75 (max) - *Mid-market*
- **Y1**: P40 (min) / P50 (mid) / P62.5 (max) - *Entry-level*

### Analytics
- 📊 Internal vs Market comparison
- 📈 Coverage analysis  
- 👥 Employee distribution by level and family
- 💱 Multi-currency support (USD, GBP, INR)

## 📁 Project Structure

```
salary-ranges/
├── SalaryRangesCalculator.gs   # ⭐ MAIN CONSOLIDATED SCRIPT (1900+ lines)
├── ExecMappingManager.html      # Web UI for job family mappings
├── appsscript.json              # Apps Script manifest
├── .clasp.json                  # ⚠️ NEEDS YOUR SCRIPT ID
├── package.json                 # npm scripts
└── archive/                     # Old individual scripts (reference only)
    ├── AppImports.gs
    ├── Helpers.gs
    └── RangeCalculator.gs
```

## 🚀 Quick Start

### 1. Install & Login (One-time)

```bash
npm install -g @google/clasp
clasp login
```

Enable Apps Script API: https://script.google.com/home/usersettings

### 2. Create Your Project

```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space/salary-ranges"

# Create new sheet with script
clasp create --type sheets --title "Salary Ranges Calculator"
```

This automatically updates `.clasp.json` with your Script ID!

### 3. Push Code

```bash
npm run push
```

This pushes:
- ✅ `SalaryRangesCalculator.gs` (all functionality)
- ✅ `ExecMappingManager.html` (web UI)
- ✅ `appsscript.json` (manifest)

### 4. Configure HiBob API

In your Google Sheet:
1. **Extensions > Apps Script**
2. **⚙️ Project Settings > Script Properties**
3. Add:
   - `BOB_ID` = `your_bob_api_id`
   - `BOB_KEY` = `your_bob_api_key`

### 5. Load Aon Data

**Aon Data Source**: [Google Drive Folder](https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-)

1. Download Aon market data files from the Drive folder
2. In your sheet: **💰 Salary Ranges Calculator > ⚙️ Setup > 🌍 Create Aon Region Tabs**
3. Paste Aon data into the created tabs:
   - `Aon US Premium - 2025`
   - `Aon UK London - 2025`
   - `Aon India - 2025`

### 6. Initial Setup

In your Google Sheet menu:

```
1. 💰 Salary Ranges Calculator > ⚙️ Setup > 🗺️ Create Mapping Tabs
2. 💰 Salary Ranges Calculator > 🏗️ Build > 🌱 Seed Exec Mappings
3. 💰 Salary Ranges Calculator > ⚙️ Setup > 📊 Build Calculator UI
```

## 📊 Using the Calculator

### Menu System

Your Google Sheet now has a **💰 Salary Ranges Calculator** menu:

#### ⚙️ Setup
- Generate Help Sheet
- Create Aon Region Tabs
- Create Mapping Tabs
- Build Calculator UI
- Manage Exec Mappings
- Ensure Category Picker

#### 📥 Import Data
- Import Bob Base Data
- Import Bonus History
- Import Comp History
- **Import All Bob Data** ⭐

#### 🏗️ Build
- Rebuild Full List Tabs ⭐
- Build Full List USD
- Seed Exec Mappings
- Fill Job Families
- Sync Employee Level Mapping
- Sync Title Mapping
- Clear All Caches

#### 📤 Export
- Export Proposed Ranges

#### 🔧 Tools
- Apply Currency Format
- Instructions & Help

### Custom Functions

Use these in Google Sheets formulas:

```javascript
// Get salary ranges by category
=SALARY_RANGE(category, region, family, ciqLevel)
=SALARY_RANGE_MIN("X0", "US", "EN.SODE", "L5 IC")
=SALARY_RANGE_MID("X1", "India", "EN.SODE", "L6 IC")
=SALARY_RANGE_MAX("Y1", "UK", "FI.FINA", "L4 IC")

// Get market percentiles
=AON_P40("US", "EN.SODE", "L5 IC")
=AON_P50("UK", "SA.SALE", "L6 IC")
=AON_P625("India", "EN.AIML", "L5.5 IC")
=AON_P75("US", "EN.SODE", "L7 IC")
=AON_P90("UK", "FI.FINA", "L5 Mgr")

// Get internal statistics
=INTERNAL_STATS("US", "EN.SODE", "L5 IC")
// Returns: [Min, Median, Max, Employee Count]

// UI versions (reads from calculator sheet)
=UI_SALARY_RANGE(region, family, level)
```

### Interactive Calculator

After running **Build Calculator UI**, use the **Salary Ranges** sheet:

1. **Select Job Family** (dropdown in B2)
2. **Select Category** (X0/X1/Y1 in B3)
3. **Select Region** (US/UK/India in B4)
4. View calculated ranges for all levels

## 📈 Workflow

### Regular Use

```
1. Import Data → Import All Bob Data
   (Syncs employee data from HiBob)

2. Build → Rebuild Full List Tabs
   (Generates comprehensive salary ranges)

3. Use the Salary Ranges sheet or formulas
   (Analyze and calculate ranges)
```

### Updating Mappings

```
- Job Families:    Setup → Manage Exec Mappings
- Employee Levels: Build → Sync Employee Level Mapping
- Job Titles:      Build → Sync Title Mapping
```

## 💻 NPM Commands

```bash
# Push to Apps Script
npm run push

# Pull from Apps Script
npm run pull

# Open in browser
npm run open

# Auto-push on file changes (requires nodemon)
npm run watch

# Deploy to Apps Script + Git
npm run deploy

# View logs
npm run logs
```

## 📋 Required Sheets

### Source Data Sheets
- **Base Data** - Employee data from HiBob
- **Bonus History** - Bonus/commission data
- **Comp History** - Compensation changes
- **Aon US Premium - 2025** - US market data
- **Aon UK London - 2025** - UK market data
- **Aon India - 2025** - India market data

### Mapping Sheets
- **Lookup** - CIQ Level → Aon Level mapping + FX rates
- **Job family Descriptions** - Aon Code → Executive Description
- **Title Mapping** - Job titles → Job families
- **Employee Level Mapping** - Employee ID → Level mapping
- **Aon Code Remap** - Code aliases (e.g., EN.SOML → EN.AIML)

### Generated Sheets
- **Full List** - Consolidated market + internal data
- **Full List USD** - FX-converted view
- **Coverage Summary** - Data completeness report
- **Employees (Mapped)** - Audit of mapped employees
- **Salary Ranges** - Interactive calculator UI

## 🔧 Troubleshooting

### "YOUR_SCRIPT_ID_HERE" Error

Update `.clasp.json`:
```bash
clasp create --type sheets --title "Salary Ranges Calculator"
```

Or manually edit `.clasp.json` with your Script ID.

### "clasp: command not found"
```bash
npm install -g @google/clasp
```

### "User has not enabled the Apps Script API"
1. Go to https://script.google.com/home/usersettings
2. Enable "Google Apps Script API"

### "Access Not Granted or Expired"
```bash
clasp logout
clasp login
```

### Data Not Showing
1. Check Script Properties (BOB_ID, BOB_KEY)
2. Run **Build → Clear All Caches**
3. Run **Build → Rebuild Full List Tabs**

### Push Failed
```bash
# Check clasp status
clasp login --status

# Verify .clasp.json
cat .clasp.json

# Try manual push
clasp push
```

## 📊 Aon Data Structure

Your Aon tabs should have these columns:
- **Job Code** (e.g., EN.SODE.P5)
- **Job Family** (e.g., Engineering - Software Development)
- **Market (43) CFY Fixed Pay: 40th Percentile** (or P40)
- **Market (43) CFY Fixed Pay: 50th Percentile** (or P50)
- **Market (43) CFY Fixed Pay: 62.5th Percentile** (or P62.5)
- **Market (43) CFY Fixed Pay: 75th Percentile** (or P75)
- **Market (43) CFY Fixed Pay: 90th Percentile** (or P90)

## 🔗 Links

- **Aon Data**: https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-
- **Apps Script API Settings**: https://script.google.com/home/usersettings
- **Your Apps Script Projects**: https://script.google.com/home
- **clasp Documentation**: https://github.com/google/clasp
- **HiBob API Docs**: https://apidocs.hibob.com/

## 📝 Notes

- The consolidated script (`SalaryRangesCalculator.gs`) contains **all functionality** in one file (~1900 lines)
- Old individual scripts are archived in `archive/` for reference
- Only the consolidated script is pushed to Apps Script (see `.claspignore`)
- Built-in caching optimizes performance (10-minute TTL)
- Engineering families (EN.*) automatically use X0/X1 categories
- Other families default to Y1 unless explicitly set

## 📄 License

ISC

---

**Version**: 3.0.0 (Consolidated)  
**Last Updated**: 2025-11-13  
**Maintainer**: MR
