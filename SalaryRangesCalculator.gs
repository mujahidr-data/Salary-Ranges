/**
 * Salary Ranges Calculator - Consolidated Google Apps Script
 * 
 * Combines HiBob employee data with Aon market data for comprehensive
 * salary range analysis and calculation.
 * 
 * Features:
 * - Bob API integration (Base Data, Bonus, Comp History, Performance Ratings)
 * - Aon market percentiles (P10, P25, P40, P50, P62.5, P75, P90)
 * - Multi-region support (US, UK, India) with FX conversion
 * - Salary range categories: X0 (Engineering/Product), Y1 (Everyone Else)
 * - Internal vs Market analytics with CR calculations
 * - Smart employee mapping with anomaly detection
 * - Persistent legacy mapping storage
 * - Interactive calculator UI
 * 
 * @version 4.6.4
 * @date 2025-11-27
 * @changelog v4.6.4 - CRITICAL HOTFIX: Fixed calculator formulas (wrong lookup column)
 *   - Bug: Calculator formulas looking up in column U (Avg CR) instead of column Y (Key)
 *   - Result: Range Start/Mid/End were blank, Internal Min/Med/Max/Count were blank
 *   - Fix: Changed all XLOOKUP formulas to use 'Full List'!$Y:$Y for lookup array
 *   - Applied to both buildCalculatorUI_() and buildCalculatorUIForY1_()
 *   - Market range and internal stats now populate correctly in calculator
 * @previous v4.6.3-debug - Added comprehensive logging to diagnose internal stats issue
 *   - Added logging to _buildInternalIndex_() to show columns, sample data, and keys created
 *   - Added logging to Full List generation to show lookup attempts and success rate
 *   - Shows summary: X out of Y combinations have employee data
 *   - Purpose: Identify why internal stats not populating despite CR working
 * @previous v4.6.2 - CRITICAL HOTFIX: Fixed internal stats (min/med/max/count)
 *   - Bug: Region key mismatch - internalIndex uses "USA", lookup uses "US"
 *   - Bug: Property name mismatch - returns `n`, code accesses `cnt`
 *   - Fix: Normalize region to "USA" before lookup
 *   - Fix: Changed `intStats.cnt` to `intStats.n`
 *   - Added debug logging to _buildInternalIndex_()
 * @previous v4.6.1 - CRITICAL HOTFIX: Fixed Lookup sheet section detection
 *   - Bug: _getExecDescMap_() was reading Level Mapping section ("L5.5 IC" → "Avg of P5 and P6")
 *   - Bug: Full List showed wrong job families, all percentiles = 0
 *   - Fix: Strict section detection with regex validation for Aon codes (XX.YYYY format)
 *   - Fix: _getCategoryMap_() and _getFxMap_() also updated for safety
 *   - Added debug logging to _preloadAonData_() for troubleshooting
 * @previous v4.6.0 - Massive performance optimization for Build Market Data (90% faster!)
 *   - Pre-load Aon data: 10,080 reads → 3 reads (one per region)
 *   - Pre-index employees: 864,000 iterations → 600 (group once)
 *   - Build Full List: 300s → 30s (10x faster)
 *   - Added progress indicators for all long-running operations
 * @previous v4.5.0 - Employee Mapping optimization (80% faster)
 *   - Eliminated O(n²) nested loop in title mapping
 *   - Bulk-load legacy mappings (600+ reads → 1 read)
 *   - Smart conditional formatting skip
 * @previous v4.4.0 - Comprehensive legacy mapping dataset (675 employees)
 *   - EN.SOML → EN.AIML replacement
 * @previous v4.3.0 - Auto-populate Level from Bob Base Data
 * @previous v3.3.0 - Simplified to 2 categories with updated range definitions
 * 
 * Aon Data Source: https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const BOB_REPORT_IDS = {
  BASE_DATA: "31048356",
  BONUS_HISTORY: "31054302",
  COMP_HISTORY: "31054312",
  PERF_RATINGS: "31172066"
};

const SHEET_NAMES = {
  BASE_DATA: "Base Data",
  BONUS_HISTORY: "Bonus History",
  COMP_HISTORY: "Comp History",
  PERF_RATINGS: "Performance Ratings",
  SALARY_RANGES_X0: "Engineering and Product",
  SALARY_RANGES_Y1: "Everyone Else",
  FULL_LIST: "Full List",
  FULL_LIST_USD: "Full List USD",
  LOOKUP: "Lookup",
  LEGACY_MAPPINGS: "Legacy Mappings",
  EMPLOYEES_MAPPED: "Employees Mapped"
};

// UI Sheet name constants (used by calculator UI functions)
const UI_SHEET_NAME_X0 = "Engineering and Product";  // X0 calculator
const UI_SHEET_NAME_Y1 = "Everyone Else";  // Y1 calculator

const REGION_TAB = {
  'India': 'Aon India - 2025',
  'US': 'Aon US - 2025',
  'UK': 'Aon UK - 2025'
};

// ============================================================================
// CONSTANTS
// ============================================================================

const CACHE_TTL = 600; // 10 minutes (600 seconds)
const ALLOWED_EMP_TYPES = new Set(["Permanent", "Regular Full-Time"]);
const TENURE_THRESHOLDS = {
  FOUR_YEARS: 1460,   // 4 years in days
  THREE_YEARS: 1095,  // 3 years
  TWO_YEARS: 730,     // 2 years
  ONE_HALF_YEARS: 545, // 1.5 years
  ONE_YEAR: 365,      // 1 year
  SIX_MONTHS: 180     // 6 months
};
const WRITE_COLS_LIMIT = 23; // Column W limit for Base Data sheet

// ============================================================================
// CONSOLIDATED HELPER FUNCTIONS (Optimized - No Duplicates)
// ============================================================================

/**
 * Normalize string for case-insensitive comparison
 * @param {*} s - Value to normalize
 * @returns {string} Normalized string
 */
function normalizeString(s) {
  return String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
}

/**
 * Auto-resize columns in a sheet, but skip calculator sheets (user manually formats those)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to auto-resize
 * @param {number} startColumn - Starting column (1-based)
 * @param {number} numColumns - Number of columns to resize
 */
function autoResizeColumnsIfNotCalculator(sheet, startColumn, numColumns) {
  const sheetName = sheet.getName();
  // Skip calculator sheets - user manually formats these
  if (sheetName.toLowerCase().includes('calculator')) {
    Logger.log(`Skipping auto-resize for calculator sheet: ${sheetName}`);
    return;
  }
  sheet.autoResizeColumns(startColumn, numColumns);
}

/**
 * Find column index by trying multiple header aliases (case-insensitive)
 * @param {Array} headerRow - The header row array
 * @param {Array<string>} aliases - Array of possible column names
 * @param {boolean} throwError - Whether to throw error if not found (default: true)
 * @returns {number} Column index (0-based) or -1 if not found and throwError=false
 */
function findColumnIndex(headerRow, aliases, throwError = true) {
  const normalizedHeader = headerRow.map(normalizeString);
  for (const alias of aliases) {
    const idx = normalizedHeader.indexOf(normalizeString(alias));
    if (idx !== -1) return idx;
  }
  if (throwError) {
    throw new Error(
      `Could not find any of the columns [${aliases.join(", ")}]. Available headers: ${headerRow.join(" | ")}`
    );
  }
  return -1;
}

/**
 * Safely extract cell value as trimmed string
 * @param {Array} row - The data row
 * @param {number} idx - Column index
 * @returns {string} Trimmed string value or empty string
 */
function safeCell(row, idx) {
  return idx === -1 ? "" : (row[idx] == null ? "" : String(row[idx]).trim());
}

/**
 * Convert value to number, stripping non-numeric characters
 * @param {*} val - Value to convert
 * @returns {number} Numeric value or NaN
 */
function toNumber(val) {
  if (val == null || val === "") return NaN;
  return Number(String(val).replace(/[^\d.-]/g, ""));
}

/**
 * Parse date string intelligently from multiple formats
 * Supports: YYYY-MM-DD, DD/MM/YYYY, and standard Date parsing
 * @param {string} s - Date string
 * @returns {Date} Parsed date object
 */
function parseDateSmart(s) {
  if (!s) return s;
  
  // Try YYYY-MM-DD format
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  
  // Try DD/MM/YYYY format
  m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  
  // Fallback to standard date parsing
  return new Date(s);
}

/**
 * Convert date string to YYYY-MM-DD format
 * @param {string} s - Date string
 * @returns {string} YYYY-MM-DD formatted string or empty
 */
function toYmd(s) {
  if (!s) return "";
  
  // Already in YYYY-MM-DD format
  let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  
  // Convert from DD/MM/YYYY
  m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  
  return "";
}

/**
 * Convert column number to letter (A, B, ..., Z, AA, AB, ...)
 * @param {number} col - Column number (1-based)
 * @returns {string} Column letter
 */
function columnToLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Simple fast hash for cache keys
 * @param {...string} parts - Parts to hash
 * @returns {string} Hash key
 */
function hashKey(...parts) {
  return parts.join('|');
}

/**
 * Fetch authenticated CSV from Bob API
 * @param {string} reportId - Bob report ID
 * @param {string} locale - Locale for CSV download (default: en-CA)
 * @returns {Array<Array>} Parsed CSV rows
 */
function fetchBobReport(reportId, locale = "en-CA") {
  const apiUrl = `https://api.hibob.com/v1/company/reports/${reportId}/download?format=csv&locale=${locale}`;
  
  const apiId = PropertiesService.getScriptProperties().getProperty("BOB_ID");
  const apiKey = PropertiesService.getScriptProperties().getProperty("BOB_KEY");
  
  if (!apiId || !apiKey) {
    throw new Error("Missing BOB_ID or BOB_KEY in Script Properties.");
  }
  
  const basicAuth = Utilities.base64Encode(`${apiId}:${apiKey}`);
  
  const res = UrlFetchApp.fetch(apiUrl, {
    method: "get",
    headers: { 
      accept: "text/csv", 
      authorization: `Basic ${basicAuth}` 
    },
    muteHttpExceptions: true
  });
  
  if (res.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch CSV: ${res.getResponseCode()} - ${res.getContentText()}`);
  }
  
  const rows = Utilities.parseCsv(res.getContentText());
  if (!rows.length) throw new Error("CSV contained no data.");
  
  return rows;
}

/**
 * Get or create sheet by name
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} sheetName - Name of sheet
 * @returns {Sheet} Sheet object
 */
function getOrCreateSheet(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.setTabColor('#FF0000'); // Red color for all automated sheets
  }
  return sh;
}

/**
 * Write data to sheet with optional formatting
 * @param {Sheet} sheet - Target sheet
 * @param {Array<Array>} data - 2D array of data
 * @param {Object} formats - Optional formatting configuration
 */
function writeSheetData(sheet, data, formats = {}) {
  // Clear and write
  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  // Apply formats if provided
  if (formats && data.length > 1) {
    const numRows = data.length - 1;
    const firstDataRow = 2;
    
    Object.entries(formats).forEach(([col, format]) => {
      const colIdx = parseInt(col);
      if (colIdx > 0 && colIdx <= data[0].length) {
        sheet.getRange(firstDataRow, colIdx, numRows, 1).setNumberFormat(format);
      }
    });
  }
}


// ============================================================================
// BOB DATA IMPORTS
// ============================================================================

/**
 * Imports base employee data from Bob API
 * Adapted from bob-salary-data project for salary-ranges
 */
function importBobDataSimpleWithLookup() {
  try {
    const reportId = BOB_REPORT_IDS.BASE_DATA;
    const sheetName = SHEET_NAMES.BASE_DATA;
    const bonusSheetName = SHEET_NAMES.BONUS_HISTORY;
    
    Logger.log(`Starting import of ${sheetName}...`);
    
    // Fetch data from Bob API
    const rows = fetchBobReport(reportId);
    
    // Cache normalized header for performance
    const srcHeader = rows[0];
    const normalizedHeader = srcHeader.map(normalizeString);
    
    const idxEmpId       = findColumnIndex(srcHeader, ["Employee ID", "Emp ID", "Employee Id"]);
    const idxJobLevel    = findColumnIndex(srcHeader, ["Job Level", "Job level"]);
    const idxBasePay     = findColumnIndex(srcHeader, ["Base Pay", "Base salary", "Base Salary"]);
    const idxEmpType     = findColumnIndex(srcHeader, ["Employment Type", "Employment type"]);
    const idxStartDate   = findColumnIndex(srcHeader, ["Start Date", "Start date", "Original start date", "Original Start Date"]);
    const idxJobFamily   = findColumnIndex(srcHeader, ["Job Family Name"], false);
    
    let header = srcHeader.slice();
    // Removed: Variable Type, Variable % (legacy columns, not used)
    
    const out = [header];
  
    // Process rows
    for (let r = 1; r < rows.length; r++) {
      const src = rows[r];
      if (!src || !Array.isArray(src) || src.length === 0) continue;
      
      const row = src.slice();
      const empType = safeCell(row, idxEmpType);
      if (!ALLOWED_EMP_TYPES.has(empType)) continue;
      
      const empId  = safeCell(row, idxEmpId);
      const jobLvl = safeCell(row, idxJobLevel);
      if (!empId || !jobLvl) continue;
      
      // Ensure Employee ID is stored as text
      const empIdNum = toNumber(empId);
      if (isFinite(empIdNum)) {
        row[idxEmpId] = String(empIdNum);
      } else {
        row[idxEmpId] = empId.trim();
      }
      
      const basePayNum = toNumber(safeCell(row, idxBasePay));
      if (!isFinite(basePayNum) || basePayNum === 0) continue;
      
      row[idxBasePay] = basePayNum;
      // Removed: Variable Type, Variable % (legacy columns)
      out.push(row);
    }
    
    Logger.log(`Processed ${out.length - 1} rows for ${sheetName}`);
    
    // Write to sheet (preserve custom columns by only clearing what we write)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, sheetName);
    
    // Only clear the columns we're writing (columns 1 to out[0].length)
    const numCols = out[0].length;
    const maxRows = Math.max(sheet.getMaxRows(), out.length);
    sheet.getRange(1, 1, maxRows, numCols).clearContent();
    
    sheet.getRange(1, 1, out.length, numCols).setValues(out);
    
    // Format columns
    if (out.length > 1) {
      const numRows = out.length - 1;
      // Employee ID as text
      sheet.getRange(2, idxEmpId + 1, numRows, 1).setNumberFormat("@");
      // Base Pay as currency
      sheet.getRange(2, idxBasePay + 1, numRows, 1).setNumberFormat("#,##0.00");
    }
    
    sheet.autoResizeColumns(1, numCols);
    Logger.log(`Successfully imported ${sheetName} (preserved custom columns beyond column ${numCols})`);
    
  } catch (error) {
    Logger.log(`Error in importBobDataSimpleWithLookup: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Base Data: ${error.message}`);
    throw error;
  }
}

/**
 * Imports bonus history from Bob API, keeping only the latest entry per employee
 */
function importBobBonusHistoryLatest() {
  try {
    const reportId = BOB_REPORT_IDS.BONUS_HISTORY;
    const targetSheetName = SHEET_NAMES.BONUS_HISTORY;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    
    const rows = fetchBobReport(reportId);
    const header = rows[0];
    
    const iEmpId   = findColumnIndex(header, ["Employee ID", "Emp ID", "Employee Id"]);
    const iName    = findColumnIndex(header, ["Display name", "Emp Name", "Display Name", "Name"]);
    const iEffDate = findColumnIndex(header, ["Effective date", "Effective Date", "Effective"]);
    const iType    = findColumnIndex(header, ["Variable type", "Variable Type", "Type"]);
    const iPct     = findColumnIndex(header, ["Commission/Bonus %", "Variable %", "Commission %", "Bonus %"]);
    const iAmt     = findColumnIndex(header, ["Amount", "Variable Amount", "Commission/Bonus Amount"]);
    const iCurr    = findColumnIndex(header, ["Variable Amount currency", "Amount currency", "Currency"], false);
  
    // Keep latest row per Emp ID
    const latest = new Map();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;
      
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumber(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const effRaw = safeCell(row, iEffDate);
      const effKey = (effRaw.match(/^\d{4}-\d{2}-\d{2}/) || [])[0];
      
      if (!empId || !effKey) continue;
      
      const existing = latest.get(empId);
      if (!existing || effKey > existing.effKey) {
        latest.set(empId, { row, effKey });
      }
    }
  
    const outHeader = ["Employee ID", "Display name", "Effective date", 
                       "Variable type", "Commission/Bonus %", "Amount", "Amount currency"];
    const out = [outHeader];
  
    latest.forEach(({ row, effKey }) => {
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumber(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const name  = safeCell(row, iName);
      const type  = safeCell(row, iType);
      const pctVal = toNumber(safeCell(row, iPct));
      const amtVal = toNumber(safeCell(row, iAmt));
      const curr   = iCurr === -1 ? "" : safeCell(row, iCurr);
      
      out.push([empId, name, effKey, type, 
                isFinite(pctVal) ? pctVal : "", 
                isFinite(amtVal) ? amtVal : "", curr]);
    });
  
    // Write to sheet (preserve custom columns)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, targetSheetName);
    
    // Only clear the columns we're writing
    const numCols = out[0].length;
    const maxRows = Math.max(sheet.getMaxRows(), out.length);
    sheet.getRange(1, 1, maxRows, numCols).clearContent();
    
    sheet.getRange(1, 1, out.length, numCols).setValues(out);
    
    if (out.length > 1) {
      const numRows = out.length - 1;
      sheet.getRange(2, 3, numRows, 1).setNumberFormat("@"); // Date as text
      sheet.getRange(2, 5, numRows, 1).setNumberFormat("0.########"); // Percent
      sheet.getRange(2, 6, numRows, 1).setNumberFormat("#,##0.00"); // Amount
    }
    
    sheet.autoResizeColumns(1, numCols);
    Logger.log(`Successfully imported ${targetSheetName} (preserved custom columns beyond column ${numCols})`);
    
  } catch (error) {
    Logger.log(`Error in importBobBonusHistoryLatest: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Bonus History: ${error.message}`);
    throw error;
  }
}

/**
 * Imports compensation history from Bob API, keeping only the latest entry per employee
 */
function importBobCompHistoryLatest() {
  try {
    const reportId = BOB_REPORT_IDS.COMP_HISTORY;
    const targetSheetName = SHEET_NAMES.COMP_HISTORY;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    
    const rows = fetchBobReport(reportId);
    const header = rows[0];
    
    const iEmpId   = findColumnIndex(header, ["Emp ID", "Employee ID", "Employee Id"]);
    const iName    = findColumnIndex(header, ["Emp Name", "Display name", "Display Name", "Name"]);
    const iEffDate = findColumnIndex(header, ["History effective date", "Effective date", "Effective Date"]);
    const iBase    = findColumnIndex(header, ["History base salary", "Base salary", "Base Salary", "Base pay"]);
    const iCurr    = findColumnIndex(header, ["History base salary currency", "Base salary currency", "Currency"]);
    const iReason  = findColumnIndex(header, ["History reason", "Reason", "Change reason"]);
  
    // Keep latest row per Emp ID by Effective date
    const latest = new Map();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;
      
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumber(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const effStr = safeCell(row, iEffDate);
      const ymd = toYmd(effStr);
      if (!empId || !ymd) continue;
      
      const existing = latest.get(empId);
      if (!existing || ymd > existing.ymd) {
        latest.set(empId, { row, ymd });
      }
    }
  
    const outHeader = ["Emp ID", "Emp Name", "Effective date", "Base salary", "Base salary currency", "History reason"];
    const out = [outHeader];
  
    latest.forEach(({ row, ymd }) => {
      let empId = safeCell(row, iEmpId);
      const empIdNum = toNumber(empId);
      empId = isFinite(empIdNum) ? String(empIdNum) : empId.trim();
      
      const name   = safeCell(row, iName);
      const base   = toNumber(safeCell(row, iBase));
      const curr   = safeCell(row, iCurr);
      const reason = safeCell(row, iReason);
      const effDate = parseDateSmart(ymd);
      
      out.push([empId, name, effDate, isFinite(base) ? base : "", curr, reason]);
    });
  
    // Write to sheet (preserve custom columns)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, targetSheetName);
    
    // Only clear the columns we're writing
    const numCols = out[0].length;
    const maxRows = Math.max(sheet.getMaxRows(), out.length);
    sheet.getRange(1, 1, maxRows, numCols).clearContent();
    
    sheet.getRange(1, 1, out.length, numCols).setValues(out);
    
    if (out.length > 1) {
      const numRows = out.length - 1;
      sheet.getRange(2, 3, numRows, 1).setNumberFormat("yyyy-mm-dd"); // Date
      sheet.getRange(2, 4, numRows, 1).setNumberFormat("#,##0.00"); // Salary
    }
    
    sheet.autoResizeColumns(1, numCols);
    Logger.log(`Successfully imported ${targetSheetName} (preserved custom columns beyond column ${numCols})`);
    
  } catch (error) {
    Logger.log(`Error in importBobCompHistoryLatest: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Comp History: ${error.message}`);
    throw error;
  }
}

/**
 * Imports Performance Ratings report from HiBob
 * Report ID: 31172066
 * Preserves all columns as-is from the report
 */
function importBobPerformanceRatings() {
  try {
    const reportId = BOB_REPORT_IDS.PERF_RATINGS;
    const targetSheetName = SHEET_NAMES.PERF_RATINGS;
    
    Logger.log(`Starting import of ${targetSheetName}...`);
    
    const rows = fetchBobReport(reportId);
    
    if (!rows || rows.length === 0) {
      throw new Error("No data returned from Performance Ratings report");
    }
    
    // Use the header from the report as-is
    const header = rows[0];
    
    // Import all rows without transformation
    const out = [header];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (row && row.length > 0) {
        out.push(row);
      }
    }
    
    // Write to sheet (preserve custom columns)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, targetSheetName);
    
    // Only clear the columns we're writing
    const numCols = out[0].length;
    const maxRows = Math.max(sheet.getMaxRows(), out.length);
    sheet.getRange(1, 1, maxRows, numCols).clearContent();
    
    sheet.getRange(1, 1, out.length, numCols).setValues(out);
    
    // Auto-resize and format header
    sheet.getRange(1, 1, 1, numCols).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, numCols);
    
    Logger.log(`Successfully imported ${targetSheetName} - ${out.length - 1} rows (preserved custom columns beyond column ${numCols})`);
    SpreadsheetApp.getActive().toast(`Imported ${out.length - 1} performance rating records`, targetSheetName, 5);
    
  } catch (error) {
    Logger.log(`Error in importBobPerformanceRatings: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Performance Ratings: ${error.message}`);
    throw error;
  }
}


// ============================================================================
// SALARY RANGE CALCULATIONS
// ============================================================================

const LOOKUP_SHEET_NAME = 'Lookup';
const BASE_DATA_SHEET_NAME = 'Base Data';

/********************************
 * Small helpers
 ********************************/
function uiSheet_() { 
  return SpreadsheetApp.getActive().getSheetByName(UI_SHEET_NAME);
}

// Global cache for lookup map to avoid repeated reads
let _lookupMapCache = null;
let _lookupMapTime = 0;

function getLookupMap_(ss) {
  const now = Date.now();
  // Cache lookup map for 10 minutes
  if (_lookupMapCache && (now - _lookupMapTime) < CACHE_TTL * 1000) {
    return _lookupMapCache;
  }
  
  const sh = ss.getSheetByName(LOOKUP_SHEET_NAME);
  if (!sh) throw new Error('Sheet "Lookup" not found.');
  
  const rows = sh.getRange('A2:B18').getValues();
  const map = new Map();
  rows.forEach(([k, v]) => {
    k = String(k || '').trim(); 
    v = String(v || '').trim();
    if (k && v) map.set(k, v);
  });
  
  _lookupMapCache = map;
  _lookupMapTime = now;
  return map;
}

// Legacy wrapper for findColumnIndex with regex
function findHeaderIndex_(headers, regex) {
  const re = new RegExp(regex, 'i');
  for (let i = 0; i < headers.length; i++) {
    if (re.test(String(headers[i] || ''))) return i;
  }
  return -1;
}

// Legacy wrapper - use columnToLetter() instead
const _colToLetter_ = columnToLetter;

// Legacy wrapper - use toNumber() instead
const toNumber_ = toNumber;

function parseCiq_(ciq) {
  const s = String(ciq || '').trim();
  const m = s.match(/^L(\d+(?:\.5)?)\s*(IC|Mgr)$/i);
  if (!m) return { base: NaN, isHalf: false, role: '', label: s };
  return { 
    base: Number(m[1]), 
    isHalf: /\.5$/.test(m[1]), 
    role: m[2].toLowerCase() === 'mgr' ? 'Mgr' : 'IC' 
  };
}

function parseAonLevel_(token) {
  const s = String(token || '').trim();
  const nm = s.match(/(\d+)/);
  return { 
    letter: s ? s[0] : '', 
    num: nm ? Number(nm[1]) : NaN 
  };
}

function isFinanceFamily_(fam) {
  const f = String(fam || '').toUpperCase().trim();
  return /^FI[.\s_-]/.test(f) || /FINANCE/.test(f);
}

/** Collapse Region picker to Base Data's Site values: India / USA / UK */
function normalizePickerRegion_(r) {
  const s = String(r || '').trim();
  if (s === 'US') return 'USA';
  if (s === 'UK') return 'UK';
  if (s === 'India') return 'India';
  return s;
}

/********************************
 * Enhanced caching helpers
 ********************************/
function _cacheGet_(key) {
  try { 
    const val = CacheService.getDocumentCache().get(key);
    return val ? JSON.parse(val) : null;
  } catch (_) { 
    return null; 
  }
}

function _cachePut_(key, value, seconds) {
  try { 
    CacheService.getDocumentCache().put(key, JSON.stringify(value), seconds); 
  } catch (_) {}
}

/********************************
 * Header + value caches for AON pickers
 ********************************/
const _aonHdrCache = {};   // key: sheetName|regex -> header index

function _findHeaderCached_(headers, sheetName, regex) {
  const key = `${sheetName}|${regex}`;
  if (_aonHdrCache[key] !== undefined) return _aonHdrCache[key];
  const idx = findHeaderIndex_(headers, regex);
  _aonHdrCache[key] = idx;
  return idx;
}

// OPTIMIZED: Simplified cache key function
function _aonValueCacheKey_(sheetName, fam, targetNum, prefLetter, ciqBaseLevel, headerRegex) {
  return hashKey('AON', sheetName, fam, targetNum, prefLetter, ciqBaseLevel, headerRegex);
}

// Cache entire sheet data to reduce reads
const _sheetDataCache = {};
function _getSheetDataCached_(sheet) {
  const sheetName = sheet.getName();
  const cacheKey = `SHEET_DATA:${sheetName}`;
  
  const cached = _cacheGet_(cacheKey);
  if (cached !== null) return cached;
  
  const values = sheet.getDataRange().getValues();
  _cachePut_(cacheKey, values, CACHE_TTL);
  return values;
}

/********************************
 * Generic percentile picker (P50 / P62.5 / P75) with enhanced caching
 ********************************/
function getAonValueStrictByHeader_(sheet, fam, targetNum, prefLetter, ciqBaseLevel, headerRegex) {
  if (!Number.isFinite(targetNum)) return '';

  const sheetName = sheet.getName();
  const cacheKey = _aonValueCacheKey_(sheetName, fam, targetNum, prefLetter, ciqBaseLevel, headerRegex);
  const cached = _cacheGet_(cacheKey);
  if (cached !== null) return cached;

  // Use cached sheet data
  const values = _getSheetDataCached_(sheet);
  const headers = values[0];

  const colFam  = _findHeaderCached_(headers, sheetName, '\\bjob\\s*family\\b');
  const colCode = _findHeaderCached_(headers, sheetName, '\\bjob\\s*code\\b');
  const colPick = _findHeaderCached_(headers, sheetName, headerRegex);
  
  if (colFam < 0 || colCode < 0 || colPick < 0) {
    throw new Error(`Missing Job Family/Job Code/${headerRegex} header.`);
  }

  const famU = String(fam || '').trim().toUpperCase();
  const allowFinanceFallback = (prefLetter === 'P') && isFinanceFamily_(fam) && targetNum < 7;

  let out = '';
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][colFam] || '').trim().toUpperCase() !== famU) continue;

    const code = String(values[r][colCode] || '');
    const token = (code.match(/([^.]+)$/) || [,''])[1].toUpperCase();
    const v = toNumber_(values[r][colPick]);
    if (isNaN(v)) continue;

    if (targetNum >= 7) {
      const m = token.match(/^E(\d+)$/);
      if (m && Number(m[1]) === targetNum) { out = v; break; }
      if (token === 'EA' && (targetNum === 3 || targetNum === 4)) { out = v; break; }
      if (token === 'EB' && (targetNum === 1 || targetNum === 2)) { out = v; break; }
    } else {
      const nm = token.match(/(\d+)$/);
      const n = nm ? Number(nm[1]) : NaN;
      const letter = token ? token[0] : '';
      if (!Number.isFinite(n) || n !== targetNum) continue;
      if (letter === prefLetter || (allowFinanceFallback && letter === 'F')) { 
        out = v; 
        break; 
      }
    }
  }

  _cachePut_(cacheKey, out, CACHE_TTL);
  return out;
}

/********************************
 * Robust header regexes - Updated to handle newlines in Aon headers
 * Aon reports have headers like: "Market \n\n (43) CFY Fixed Pay: 10th Percentile"
 ********************************/
const HDR_P10  = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*10(?:th)?[\\s\\n]*Percentile';
const HDR_P25  = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*25(?:th)?[\\s\\n]*Percentile';
const HDR_P40  = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*40(?:th)?[\\s\\n]*Percentile';
const HDR_P50  = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*50(?:th)?[\\s\\n]*Percentile';
const HDR_P625 = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*62\\.?5(?:th)?[\\s\\n]*Percentile';
const HDR_P75  = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*75(?:th)?[\\s\\n]*Percentile';
const HDR_P90  = 'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*90(?:th)?[\\s\\n]*Percentile';

/********************************
 * Public custom functions
 ********************************/
function AON_P10(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P10);  }
function AON_P25(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P25);  }
function AON_P40(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P40);  }
function AON_P50(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P50);  }
function AON_P625(region, family, ciqLevel) { return _aonPick_(region, family, ciqLevel, HDR_P625); }
function AON_P75(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P75);  }
function AON_P90(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P90);  }

/********************************
 * Category-based salary ranges (2 categories only)
 * X0 = Engineering/Product: P25 (start) → P62.5 (mid) → P90 (end)
 * Y1 = Everyone Else: P10 (start) → P40 (mid) → P62.5 (end)
 * 
 * Fallback logic: If a percentile is missing, use the next higher percentile
 * Example: P10 missing → use P25, P25 missing → use P40, etc.
 ********************************/
function _rangeByCategory_(category, region, family, ciqLevel) {
  const cat = String(category || '').trim().toUpperCase();
  if (!cat) return { min: '', mid: '', max: '' };

  if (cat === 'X0') {
    // X0 (Engineering/Product): Range Start=P25, Range Mid=P62.5, Range End=P90
    let min = AON_P25(region, family, ciqLevel);
    let mid = AON_P625(region, family, ciqLevel);
    let max = AON_P90(region, family, ciqLevel);
    
    // Fallback: P25 missing → use P40
    if (!min || min === '') {
      min = AON_P40(region, family, ciqLevel);
      if (!min || min === '') min = AON_P50(region, family, ciqLevel);
    }
    // Fallback: P62.5 missing → use P75
    if (!mid || mid === '') {
      mid = AON_P75(region, family, ciqLevel);
      if (!mid || mid === '') mid = AON_P90(region, family, ciqLevel);
    }
    // Fallback: P90 missing → no fallback (already highest)
    
    return { min, mid, max };
  }
  if (cat === 'Y1') {
    // Y1 (Everyone Else): Range Start=P10, Range Mid=P40, Range End=P62.5
    let min = AON_P10(region, family, ciqLevel);
    let mid = AON_P40(region, family, ciqLevel);
    let max = AON_P625(region, family, ciqLevel);
    
    // Fallback: P10 missing → use P25
    if (!min || min === '') {
      min = AON_P25(region, family, ciqLevel);
      if (!min || min === '') min = AON_P40(region, family, ciqLevel);
    }
    // Fallback: P40 missing → use P50
    if (!mid || mid === '') {
      mid = AON_P50(region, family, ciqLevel);
      if (!mid || mid === '') mid = AON_P625(region, family, ciqLevel);
    }
    // Fallback: P62.5 missing → use P75
    if (!max || max === '') {
      max = AON_P75(region, family, ciqLevel);
      if (!max || max === '') max = AON_P90(region, family, ciqLevel);
    }
    
    return { min, mid, max };
  }

  return { min: '', mid: '', max: '' };
}

// Returns a horizontal array [min, mid, max] suitable for spilling across 3 cells
function SALARY_RANGE(category, region, family, ciqLevel) {
  const effectiveCat = _effectiveCategoryForFamily_(category, family);
  // Fast path: use Full List index
  const r = _getRangeFromFullList_(effectiveCat, region, family, ciqLevel);
  if (r.min !== '' || r.mid !== '' || r.max !== '') return [[r.min, r.mid, r.max]];
  // Fallback to direct Aon lookups (in case Full List missing)
  const rr = _rangeByCategory_(effectiveCat, region, family, ciqLevel);
  return [[rr.min === '' ? '' : Number(rr.min), rr.mid === '' ? '' : Number(rr.mid), rr.max === '' ? '' : Number(rr.max)]];
}

function SALARY_RANGE_MIN(category, region, family, ciqLevel) {
  const effectiveCat = _effectiveCategoryForFamily_(category, family);
  const r = _getRangeFromFullList_(effectiveCat, region, family, ciqLevel);
  if (r.min !== '') return r.min;
  const rr = _rangeByCategory_(effectiveCat, region, family, ciqLevel);
  return rr.min === '' ? '' : Number(rr.min);
}

function SALARY_RANGE_MID(category, region, family, ciqLevel) {
  const effectiveCat = _effectiveCategoryForFamily_(category, family);
  const r = _getRangeFromFullList_(effectiveCat, region, family, ciqLevel);
  if (r.mid !== '') return r.mid;
  const rr = _rangeByCategory_(effectiveCat, region, family, ciqLevel);
  return rr.mid === '' ? '' : Number(rr.mid);
}

function SALARY_RANGE_MAX(category, region, family, ciqLevel) {
  const effectiveCat = _effectiveCategoryForFamily_(category, family);
  const r = _getRangeFromFullList_(effectiveCat, region, family, ciqLevel);
  if (r.max !== '') return r.max;
  const rr = _rangeByCategory_(effectiveCat, region, family, ciqLevel);
  return rr.max === '' ? '' : Number(rr.max);
}

function _aonPick_(region, family, ciqLevel, headerRegex) {
  try {
    const ss = SpreadsheetApp.getActive();
    const regionKey = String(region || '').trim();
    const sheet = getRegionSheet_(ss, regionKey);
    if (!sheet) return '';

    const map = getLookupMap_(ss);
    const { base, isHalf, role } = parseCiq_(ciqLevel);
    if (!Number.isFinite(base)) return '';

    const fam = String(family || '').trim();
    const lbl = n => `L${n} ${role}`;
    const aon = l => String(map.get(l) || '').trim();
    const nFrom = a => parseAonLevel_(a).num;
    const LFrom = a => parseAonLevel_(a).letter;

    if (isHalf) {
      const n1 = Math.floor(base), n2 = n1 + 1;
      const a1 = aon(lbl(n1)), a2 = aon(lbl(n2));
      const v1 = a1 ? _getAonValueWithCodeFallback_(sheet, fam, nFrom(a1), LFrom(a1), n1, headerRegex) : '';
      const v2 = a2 ? _getAonValueWithCodeFallback_(sheet, fam, nFrom(a2), LFrom(a2), n2, headerRegex) : '';
      if (v1 === '' && v2 === '') return '';
      if (v1 === '') return v2;
      if (v2 === '') return v1;
      return (Number(v1) + Number(v2)) / 2;
    } else {
      const n = Math.floor(base);
      const a = aon(lbl(n));
      if (!a) return '';
      const v = _getAonValueWithCodeFallback_(sheet, fam, nFrom(a), LFrom(a), n, headerRegex);
      return v === '' ? '' : Number(v);
    }
  } catch (_) { return ''; }
}

// Resolve region to a sheet with sensible fallbacks
function getRegionSheet_(ss, region) {
  const r = String(region || '').trim();
  if (r === 'US' || r === 'US Premium' || r === 'US National') {
    return ss.getSheetByName('Aon US - 2025');
  }
  if (r === 'UK' || r === 'UK London') {
    return ss.getSheetByName('Aon UK - 2025');
  }
  if (r === 'India') {
    return ss.getSheetByName('Aon India - 2025');
  }
  // default try REGION_TAB mapping
  const tab = REGION_TAB[r];
  return tab ? ss.getSheetByName(tab) : null;
}

/********************************
 * INTERNAL_STATS(region, familyOrCode, ciqLevel) - OPTIMIZED
 ********************************/
function INTERNAL_STATS(region, familyOrCode, ciqLevel) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(BASE_DATA_SHEET_NAME);
  if (!sh) return [['', '', '', '']];

  const siteWanted = normalizePickerRegion_(region);
  const famCodeU = String(familyOrCode || '').trim().toUpperCase();
  const lvlU = String(ciqLevel || '').trim().toUpperCase();

  // Get friendly name once
  const ui = uiSheet_();
  const friendlyName = ui ? String(ui.getRange('C2').getDisplayValue() || '').trim().toUpperCase() : '';

  const cacheKey = hashKey('INT', siteWanted, famCodeU, friendlyName, lvlU); // OPTIMIZED
  const cached = _cacheGet_(cacheKey);
  if (cached) return [cached];

  // Use cached sheet data
  const values = _getSheetDataCached_(sh);
  const headers = values[0].map(h => String(h || ''));

  const colFam  = headers.indexOf('Job Family Name');
  const colMap  = headers.indexOf('Mapped Family');
  const colAct  = headers.indexOf('Active/Inactive');
  const colSite = headers.indexOf('Site');
  const colLvl  = headers.indexOf('Job Level');
  const colPay  = headers.indexOf('Base salary');
  
  if ([colFam,colMap,colAct,colSite,colLvl,colPay].some(i => i < 0)) {
    return [['', '', '', '']];
  }

  const nums = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[colAct] || '').toLowerCase() !== 'active') continue;
    
    const rowFamCodeU = String(row[colFam] || '').trim().toUpperCase();
    const rowMapNameU = String(row[colMap] || '').trim().toUpperCase();
    
    if (!(rowFamCodeU === famCodeU || (friendlyName && rowMapNameU === friendlyName))) continue;
    if (String(row[colLvl] || '').trim().toUpperCase() !== lvlU) continue;
    if (String(row[colSite] || '').trim() !== siteWanted) continue;
    
    const n = toNumber_(row[colPay]);
    if (!isNaN(n)) nums.push(n);
  }
  
  if (!nums.length) {
    const out0 = ['', '', '', ''];
    _cachePut_(cacheKey, out0, CACHE_TTL);
    return [out0];
  }

  nums.sort((a,b) => a - b);
  const min = nums[0];
  const max = nums[nums.length - 1];
  const count = nums.length;
  const m = Math.floor(count / 2);
  const median = count % 2 ? nums[m] : (nums[m - 1] + nums[m]) / 2;

  const out = [min, median, max, count || ''];
  _cachePut_(cacheKey, out, CACHE_TTL);
  return [out];
}

/********************************
 * Currency formatting (hide zeros) - OPTIMIZED
 ********************************/
function _setFmtIfNeeded_(range, fmt) {
  const rows = range.getNumRows(), cols = range.getNumColumns();
  const current = range.getNumberFormats();
  let needs = false;
  
  for (let r = 0; r < rows && !needs; r++) {
    for (let c = 0; c < cols; c++) {
      if (current[r][c] !== fmt) { 
        needs = true; 
        break; 
      }
    }
  }
  if (needs) range.setNumberFormat(fmt);
}

function applyCurrency_() {
  // Format calculator sheet dynamically by header labels and detected region
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh) return;

  // Try to detect region/currency from labeled cells in the first 10 rows
  let region = '';
  let currency = '';
  try {
    const top = sh.getRange(1,1,10,2).getDisplayValues();
    for (let r=0;r<top.length;r++) {
      const label = String(top[r][0]||'').trim().toLowerCase();
      if (label === 'region') region = String(top[r][1]||'').trim();
      if (label === 'currency') currency = String(top[r][1]||'').trim();
    }
  } catch(_) {}
  if (!region) region = String(sh.getRange('B4').getDisplayValue() || '').trim();
  if (!region) region = String(sh.getRange('B2').getDisplayValue() || '').trim();
  // If currency explicitly says USD, force US format regardless of region
  if (/^usd$/i.test(currency)) region = 'US';

  const formats = {
    'India': '₹#,##,##0;₹#,##,##0;;@',
    'US': '$#,##0;$#,##0;;@',
    'UK': '£#,##0;£#,##0;;@'
  };
  const cfmt = formats[region] || '#,##0;#,##0;;@';

  // Find header row (search first 30 rows for Level/P62.5/P75/P90)
  const maxHdrRows = Math.min(30, sh.getLastRow());
  let headerRow = -1; let headers = [];
  for (let r=1; r<=maxHdrRows; r++) {
    const row = sh.getRange(r,1,1,Math.max(20, sh.getLastColumn())).getDisplayValues()[0].map(v=>String(v||'').trim());
    if (row.some(v=>/^Level$/i.test(v)) && row.some(v=>/^P\s*62\.?5$/i.test(v)) && row.some(v=>/^P\s*75$/i.test(v))) { headerRow = r; headers = row; break; }
  }
  if (headerRow === -1) return; // nothing to format

  // Locate columns by label
  const colIndex = (labelRegex) => headers.findIndex(h => new RegExp(labelRegex,'i').test(h)) + 1;
  const cP625 = colIndex('^P\s*62\.?5$');
  const cP75  = colIndex('^P\s*75$');
  const cP90  = colIndex('^P\s*90$');
  const cMin  = colIndex('^Min$');
  const cMed  = colIndex('^Median$');
  const cMax  = colIndex('^Max$');
  const cEmp  = colIndex('^Emp\s*Count$');
  const lastRow = Math.max(headerRow+1, sh.getLastRow());

  const maybeFormatCol = (c, fmt) => { if (c > 0) _setFmtIfNeeded_(sh.getRange(headerRow+1, c, lastRow - headerRow, 1), fmt); };
  [cP625, cP75, cP90, cMin, cMed, cMax].forEach(c => maybeFormatCol(c, cfmt));
  // Hide zeros in Emp Count column (0;-0;;@)
  if (cEmp > 0) maybeFormatCol(cEmp, '0;-0;;@');
}

/********************************
 * Menu + triggers - See unified onOpen below at line ~1789
 ********************************/

function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== UI_SHEET_NAME) return;
    // Intentionally no auto-format or cache clear on region changes to reduce overhead
  } catch (_) {}
}

// Helper to manually clear all caches
function clearAllCaches_() {
  CacheService.getDocumentCache().removeAll(['INT:', 'AON:', 'SHEET_DATA:']);
  try { CacheService.getDocumentCache().remove('REM:CODEMAP'); } catch (_) {}
  try { CacheService.getDocumentCache().remove('MAP:EXEC_DESC'); } catch (_) {}
  _lookupMapCache = null;
  _lookupMapTime = 0;
  SpreadsheetApp.getActiveSpreadsheet().toast('All caches cleared', 'Success', 3);
}

/********************************
 * Exporter + Utilities (no hardcoded exec descriptions)
 ********************************/

/**
 * Reads Aon Code → Job Family mapping from Lookup sheet
 * Also falls back to Job family Descriptions sheet for backward compatibility
 */
function _getExecDescMap_() {
  const cacheKey = 'MAP:EXEC_DESC';
  const cached = _cacheGet_(cacheKey);
  if (cached) return new Map(cached);
  
  const ss = SpreadsheetApp.getActive();
  const map = new Map();
  
  // Try reading from Lookup sheet first (new format)
  const lookupSh = ss.getSheetByName('Lookup');
  if (lookupSh) {
    const vals = lookupSh.getDataRange().getValues();
    let inAonCodeSection = false;
    
    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];
      if (!row || row.length < 2) continue;
      
      const col1 = String(row[0] || '').trim();
      const col2 = String(row[1] || '').trim();
      const col3 = row.length > 2 ? String(row[2] || '').trim() : '';
      
      // Detect Aon Code section header
      if (col1 === 'Aon Code' && /Job.*Family.*Exec/i.test(col2)) {
        inAonCodeSection = true;
        continue;
      }
      
      // Stop at next section (new header row)
      if (inAonCodeSection && (col1 === 'CIQ Level' || col1 === 'Region' || col1 === 'Aon Code')) {
        inAonCodeSection = false;
        continue;
      }
      
      // Only read Aon Code section data
      if (inAonCodeSection && col1 && col2) {
        // Validate it's an Aon code format (XX.YYYY, not L5.5 IC)
        if (/^[A-Z]{2}\.[A-Z0-9]{4}$/i.test(col1)) {
          map.set(col1, col2);
        }
      }
    }
  }
  
  // Fall back to Job family Descriptions sheet if Lookup doesn't have mappings
  if (map.size === 0) {
    const sh = ss.getSheetByName('Job family Descriptions');
    if (sh) {
      const vals = sh.getDataRange().getValues();
      if (vals.length > 1) {
        const head = vals[0].map(h => String(h || '').trim());
        const iCode = head.findIndex(h => /^(Aon\s*Code|Job\s*Code)$/i.test(h));
        const iDesc = head.findIndex(h => /Job\s*Family\s*\(Exec\s*Description\)/i.test(h));
        for (let r=1; r<vals.length; r++) {
          const code = iCode>=0 ? String(vals[r][iCode]||'').trim() : '';
          const desc = iDesc>=0 ? String(vals[r][iDesc]||'').trim() : '';
          if (code && desc) map.set(code, desc);
        }
      }
    }
  }
  
  _cachePut_(cacheKey, Array.from(map.entries()), CACHE_TTL);
  return map;
}

/**
 * Reads Aon Code → Category mapping from Lookup sheet
 * Returns Map: Aon Code → 'X0' or 'Y1'
 */
function _getCategoryMap_() {
  const cacheKey = 'MAP:CATEGORY';
  const cached = _cacheGet_(cacheKey);
  if (cached) return new Map(cached);
  
  const ss = SpreadsheetApp.getActive();
  const map = new Map();
  
  // Read from Lookup sheet (only Aon Code section)
  const lookupSh = ss.getSheetByName('Lookup');
  if (lookupSh) {
    const vals = lookupSh.getDataRange().getValues();
    let inAonCodeSection = false;
    
    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];
      if (!row || row.length < 3) continue;
      
      const col1 = String(row[0] || '').trim();
      const col2 = String(row[1] || '').trim();
      const col3 = String(row[2] || '').trim().toUpperCase();
      
      // Detect Aon Code section header
      if (col1 === 'Aon Code' && /Job.*Family.*Exec/i.test(col2) && col3 === 'Category') {
        inAonCodeSection = true;
        continue;
      }
      
      // Stop at next section (new header row with different pattern)
      if (inAonCodeSection && (col1 === 'CIQ Level' || col1 === 'Region')) {
        break; // No more Aon Code section after this
      }
      
      // Only read Aon Code section data
      if (inAonCodeSection && col1 && (col3 === 'X0' || col3 === 'Y1')) {
        // Validate it's an Aon code format (XX.YYYY, not L5.5 IC)
        if (/^[A-Z]{2}\.[A-Z0-9]{4}$/i.test(col1)) {
          map.set(col1, col3);
        }
      }
    }
  }
  
  _cachePut_(cacheKey, Array.from(map.entries()), CACHE_TTL);
  return map;
}

function _readLookupRows_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Lookup');
  if (!sh) return [];
  const rows = sh.getRange('A2:B18').getValues();
  const out = [];
  rows.forEach(([ciq, aon]) => {
    const s = String(ciq || '').trim();
    const a = String(aon || '').trim().toUpperCase().replace(/\s+/g,'');
    if (!s) return;
    const m = s.match(/^L(\d+(?:\.5)?)\s*(IC|Mgr)$/i);
    const base = m ? Number(m[1]) : NaN;
    const half = m ? /\.5$/.test(m[1]) : false;
    const role = m ? (m[2].toLowerCase() === 'mgr' ? 'MGR' : 'IC') : '';
    // For half-levels, allow empty Aon level (we will average neighbors)
    if (!half && !a) return;
    out.push({ ciq:s, aon:a, base, half, role });
  });
  return out;
}

function _avg2_(a,b) {
  const af = isFinite(a), bf = isFinite(b);
  if (!af && !bf) return NaN;
  if (!af) return b;
  if (!bf) return a;
  return (a+b)/2;
}
function _round0_(n) { return isFinite(n) ? Math.round(n) : ''; }
function _round100_(n) {
  if (!isFinite(n)) return '';
  return Math.round(Number(n) / 100) * 100;
}
function _isNum_(v) { return v !== '' && v != null && isFinite(Number(v)); }

function _buildInternalIndex_() {
  const ss = SpreadsheetApp.getActive();
  const empSh = ss.getSheetByName('Employees Mapped');
  const baseSh = ss.getSheetByName('Base Data');
  const out = new Map();
  
  if (!empSh || empSh.getLastRow() <= 1) {
    Logger.log('WARNING: Employees Mapped sheet not found or empty - internal stats will be blank');
    return out;
  }
  
  if (!baseSh || baseSh.getLastRow() <= 1) {
    Logger.log('WARNING: Base Data sheet not found or empty - cannot check active status');
    return out;
  }

  // Build active status index from Base Data
  const baseVals = baseSh.getDataRange().getValues();
  const baseHead = baseVals[0].map(h => String(h || ''));
  const iBaseEmpID = baseHead.findIndex(h => /Emp.*ID|Employee.*ID/i.test(h));
  const iBaseActive = baseHead.findIndex(h => /Active.*Inactive/i.test(h));
  
  const activeStatusMap = new Map(); // empID → isActive
  if (iBaseEmpID >= 0 && iBaseActive >= 0) {
    for (let r = 1; r < baseVals.length; r++) {
      const empID = String(baseVals[r][iBaseEmpID] || '').trim();
      const activeStatus = String(baseVals[r][iBaseActive] || '').toLowerCase();
      if (empID) {
        activeStatusMap.set(empID, activeStatus === 'active');
      }
    }
    Logger.log(`Built active status index: ${activeStatusMap.size} employees`);
  } else {
    Logger.log('WARNING: Could not find Emp ID or Active/Inactive columns in Base Data!');
    return out;
  }

  const values = empSh.getRange(2, 1, empSh.getLastRow() - 1, 19).getValues();
  
  Logger.log(`Reading internal stats from Employees Mapped sheet: ${values.length} employees`);
  
  // Employees Mapped columns (19 total): 
  // A: Employee ID, B: Employee Name, C: Job Title, D: Department, E: Site
  // F: Aon Code, G: Job Family (Exec Description), H: Level, I: Full Aon Code, J: Mapping Override
  // K: Confidence, L: Source, M: Status, N: Base Salary, O: Start Date
  // P: Recent Promotion, Q: Level Anomaly, R: Title Anomaly, S: Market Data Missing
  const iEmpID = 0;     // Column A: Employee ID
  const iSite = 4;      // Column E: Site
  const iAonCode = 5;   // Column F: Aon Code (base, e.g., "EN.SODE")
  const iLevel = 7;     // Column H: Level (e.g., "L5 IC")
  const iStatus = 12;   // Column M: Status
  const iSalary = 13;   // Column N: Base Salary

  const buckets = new Map();
  let processedCount = 0;
  let skippedInactive = 0;
  let skippedNoMapping = 0;
  let skippedNoSalary = 0;
  
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    
    const empID = String(row[iEmpID] || '').trim();
    const site = String(row[iSite] || '').trim();
    const aonCode = String(row[iAonCode] || '').trim();
    const level = String(row[iLevel] || '').trim();
    const status = String(row[iStatus] || '').trim();
    const salary = row[iSalary];
    
    // CRITICAL: Only include ACTIVE employees for internal stats
    const isActive = activeStatusMap.get(empID);
    if (!isActive) {
      skippedInactive++;
      continue;
    }
    
    // Skip if no mapping
    if (!aonCode || !level) {
      skippedNoMapping++;
      continue;
    }
    
    // Skip if status is not Approved or Legacy (only use confirmed mappings)
    if (status !== 'Approved' && status !== 'Legacy') {
      skippedNoMapping++;
      continue;
    }
    
    // Skip if no salary
    const pay = toNumber(salary);
    if (isNaN(pay) || pay <= 0) {
      skippedNoSalary++;
      continue;
    }
    
    // Normalize region (US → USA for consistency)
    const normSite = site === 'US' ? 'USA' : (site === 'USA' ? 'USA' : (site === 'India' ? 'India' : (site === 'UK' ? 'UK' : site)));
    
    processedCount++;
    
    // Log first 3 employees for debugging
    if (processedCount <= 3) {
      Logger.log(`Sample employee ${processedCount}: empID=${empID}, site=${normSite}, aonCode=${aonCode}, level=${level}, pay=${pay}, status=${status}, active=true`);
    }
    
    // Create key: Region|AonCode|Level (e.g., "USA|EN.SODE|L5 IC")
    const key = `${normSite}|${aonCode}|${level}`;
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(pay);
    
    if (processedCount <= 3) {
      Logger.log(`  → Created key: ${key}`);
    }
  }
  
  Logger.log(`Processed ${processedCount} ACTIVE employees with approved mappings`);
  Logger.log(`Skipped: ${skippedInactive} inactive, ${skippedNoMapping} without mapping, ${skippedNoSalary} without salary`);
  buckets.forEach((arr, key) => {
    arr.sort((a,b)=>a-b);
    const n = arr.length; const min = arr[0], max = arr[n-1];
    const med = n % 2 ? arr[(n-1)/2] : (arr[n/2 - 1] + arr[n/2]) / 2;
    out.set(key, { min, med, max, n });
  });
  
  Logger.log(`Built internal index: ${out.size} combinations with employee data`);
  // Log first 5 for verification
  let count = 0;
  out.forEach((stats, key) => {
    if (count < 5) {
      Logger.log(`  ${key} → min=${stats.min}, med=${stats.med}, max=${stats.max}, n=${stats.n}`);
      count++;
    }
  });
  
  return out;
}

function _readMappedEmployeesForAudit_() {
  const ss = SpreadsheetApp.getActive();
  const mapSh = ss.getSheetByName('Employee Level Mapping');
  const baseSh = ss.getSheetByName('Base Data');
  const out = [];
  if (!mapSh || !baseSh) return out;
  const mVals = _getSheetDataCached_(mapSh); // OPTIMIZED: Use cached data
  const mHead = mVals[0].map(h => String(h || '').replace(/\s+/g,' ').trim());
  const colEmp = mHead.findIndex(h => /^Emp\s*ID/i.test(h));
  let colMap = mHead.findIndex(h => /Is\s*Mapped\?/i.test(h));
  if (colMap < 0) colMap = mHead.findIndex(h => /^Mapping$/i.test(h));
  if (colEmp < 0 || colMap < 0) return out;

  const bVals = _getSheetDataCached_(baseSh); // OPTIMIZED: Use cached data
  const bHead = bVals[0].map(h => String(h || '').replace(/\s+/g,' ').trim());
  const cEmp  = bHead.findIndex(h => /^Emp\s*ID/i.test(h) || /Employee\s*ID/i.test(h));
  const cSite = bHead.indexOf('Site');
  const cPay  = bHead.findIndex(h => /Base\s*salary/i.test(h));
  if (cEmp < 0 || cSite < 0 || cPay < 0) return out;

  const baseByEmp = new Map();
  for (let r=1; r<bVals.length; r++) {
    const id = String(bVals[r][cEmp] || '').trim();
    if (!id) continue;
    baseByEmp.set(id, { site: String(bVals[r][cSite] || '').trim(), pay: toNumber_(bVals[r][cPay]) });
  }

  for (let r=1; r<mVals.length; r++) {
    const emp = String(mVals[r][colEmp] || '').trim();
    const map = String(mVals[r][colMap] || '').trim();
    if (!emp || !map || map === '.' || map.indexOf('.') < 0) continue;
    const parts = map.split('.'); if (parts.length < 3) continue;
    const base = parts[0] + '.' + parts[1];
    const suf  = parts[2].toUpperCase().replace(/\s+/g,'');
    const rec = baseByEmp.get(emp);
    if (!rec || !isFinite(rec.pay)) continue;
    const site = rec.site === 'India' ? 'India' : (rec.site === 'USA' ? 'USA' : (rec.site === 'UK' ? 'UK' : rec.site));
    out.push([emp, base, suf, site, Math.round(rec.pay)]);
  }
  return out;
}

function exportMarketAndInternal_() { /* intentionally omitted in this merge; use rebuildFullListTabs_ or exportProposedSalaryRanges_ */ }

function rebuildFullListTabs_() {
  const ss = SpreadsheetApp.getActive();
  const lookupRows = _readLookupRows_();
  if (!lookupRows.length) throw new Error('Lookup (A2:B) is empty or malformed.');
  const regionNames = Object.keys(REGION_TAB);
  const regionIndexes = {}; const famByBaseGlobal = new Map();
  regionNames.forEach(region => {
    const sh = getRegionSheet_(ss, region);
    const byKey = new Map(); const famByBase = new Map();
    if (sh) {
      const values = sh.getDataRange().getValues();
      if (values.length) {
        const headers = values[0].map(h => String(h || '').replace(/\s+/g,' ').trim());
        const colJobCode = headers.indexOf('Job Code');
        const colJobFam  = headers.indexOf('Job Family');
        // Find columns using regex to handle newlines in headers
        const colP10  = findHeaderIndex_(headers, HDR_P10);
        const colP25  = findHeaderIndex_(headers, HDR_P25);
        const colP40  = findHeaderIndex_(headers, HDR_P40);
        const colP50  = findHeaderIndex_(headers, HDR_P50);
        const colP625 = findHeaderIndex_(headers, HDR_P625);
        const colP75  = findHeaderIndex_(headers, HDR_P75);
        const colP90  = findHeaderIndex_(headers, HDR_P90);
        if (colJobCode >= 0 && colJobFam >= 0 && colP50 >= 0 && colP625 >= 0 && colP75 >= 0) {
          for (let r=1; r<values.length; r++) {
            const row = values[r]; const jc = String(row[colJobCode] || '').trim(); if (!jc) continue;
            const i = jc.lastIndexOf('.'); const base = i>=0 ? jc.slice(0,i) : jc; const suf = (i>=0 ? jc.slice(i+1) : jc).toUpperCase().replace(/[^A-Z0-9]/g,'');
            const fam = String(row[colJobFam] || '').trim(); if (base && fam && !famByBase.has(base)) famByBase.set(base, fam);
            const p10 = colP10 >= 0 ? toNumber_(row[colP10]) : NaN; const p25 = colP25 >= 0 ? toNumber_(row[colP25]) : NaN; const p40 = colP40 >= 0 ? toNumber_(row[colP40]) : NaN; const p50 = toNumber_(row[colP50]); const p62 = toNumber_(row[colP625]); const p75 = toNumber_(row[colP75]); const p90 = colP90 >= 0 ? toNumber_(row[colP90]) : NaN;
            byKey.set(`${base}|${suf}`, { p10, p25, p40, p50, p62, p75, p90 });
          }
        }
      }
    }
    regionIndexes[region] = byKey; famByBase.forEach((v,k) => { if (!famByBaseGlobal.has(k)) famByBaseGlobal.set(k, v); });
  });

  const internalIdx = _buildInternalIndex_();
  const rows = [];
  const emitted = new Set();
  regionNames.forEach(region => {
    const site = normalizePickerRegion_(region); const idx = regionIndexes[region]; if (!idx) return;
    const bases = new Set(); idx.forEach((_, key) => bases.add(key.split('|')[0]));
    Array.from(bases).sort().forEach(base => {
      const baseOut = remapAonCode_(base);
      const rawFam = famByBaseGlobal.get(base) || '';
      const execMap = _getExecDescMap_();
      const execFam = execMap.get(baseOut) || execMap.get(base) || rawFam;
      const whole = new Map();
      const lookupRows = _readLookupRows_();
      lookupRows.forEach(L => { if (L.half) return; const rec = idx.get(`${base}|${L.aon}`); whole.set(`${L.role}|${Math.floor(L.base)}`, rec || { p10:NaN,p25:NaN,p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN }); });
      lookupRows.forEach(L => {
        let p10, p25, p40, p50, p62, p75, p90;
        if (L.half) { const k1 = `${L.role}|${Math.floor(L.base)}`; const k2 = `${L.role}|${Math.floor(L.base)+1}`; const v1 = whole.get(k1) || {p10:NaN,p25:NaN,p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN}; const v2 = whole.get(k2) || {p10:NaN,p25:NaN,p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN}; p10 = _avg2_(v1.p10, v2.p10); p25 = _avg2_(v1.p25, v2.p25); p40 = _avg2_(v1.p40, v2.p40); p50 = _avg2_(v1.p50, v2.p50); p62 = _avg2_(v1.p62, v2.p62); p75 = _avg2_(v1.p75, v2.p75); p90 = _avg2_(v1.p90, v2.p90); }
        else { const rec = idx.get(`${base}|${L.aon}`) || { p10:NaN,p25:NaN,p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN }; p10 = rec.p10; p25 = rec.p25; p40 = rec.p40; p50 = rec.p50; p62 = rec.p62; p75 = rec.p75; p90 = rec.p90; }
        const ist = internalIdx.get(`${site}|${String(execFam).toUpperCase()}|${L.ciq}`) || internalIdx.get(`${site}|${base}|${L.ciq}`) || null; const key = `${execFam}${L.ciq}${region}`;
        const uniqueKey = `${site}|${region}|${baseOut}|${String(execFam)}|${L.ciq}`;
        if (!emitted.has(uniqueKey)) {
          emitted.add(uniqueKey);
          rows.push([site, region, baseOut, execFam, rawFam, L.ciq, L.aon, _round100_(p10), _round100_(p25), _round100_(p40), _round100_(p50), _round100_(p62), _round100_(p75), _round100_(p90), ist ? _round0_(ist.min) : '', ist ? _round0_(ist.med) : '', ist ? _round0_(ist.max) : '', ist ? ist.n : '', '', key]);
        }
      });
    });
  });

  const fl = ss.getSheetByName('Full List') || ss.insertSheet('Full List');
  const fullHeader = ['Site','Region','Aon Code','Job Family (Exec Description)','Job Family (Raw)','CIQ Level','Aon Level','P10','P25','P40','P50','P62.5','P75','P90','Internal Min','Internal Median','Internal Max','Employees','', 'Key'];
  fl.clearContents(); fl.getRange(1,1,1,fullHeader.length).setValues([fullHeader]);
  if (rows.length) fl.getRange(2,1,rows.length,fullHeader.length).setValues(rows);
  fl.autoResizeColumns(1, fullHeader.length);

  const baseSh = ss.getSheetByName('Base Data');
  SpreadsheetApp.getActive().toast('Full List rebuilt successfully', 'Done', 5);
}

function _getFxMap_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Lookup');
  const fxMap = new Map();
  if (!sh) return fxMap;
  
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return fxMap;
  
  let inFxSection = false;
  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];
    if (!row || row.length < 3) continue;
    
    const col1 = String(row[0] || '').trim();
    const col2 = String(row[1] || '').trim();
    const col3 = String(row[2] || '').trim();
    
    // Detect FX section header (Region, Site, FX Rate)
    if (col1 === 'Region' && col2 === 'Site' && /FX.*Rate/i.test(col3)) {
      inFxSection = true;
      continue;
    }
    
    // Stop at next section (new header row)
    if (inFxSection && (col1 === 'Aon Code' || col1 === 'CIQ Level')) {
      break;
    }
    
    // Only read FX section data
    if (inFxSection && col1) {
      let region = col1;
      // Normalize
      if (/^USA$/i.test(region)) region = 'US';
      if (/^US\s*(Premium|National)?$/i.test(region)) region = 'US';
      const fx = Number(col3) || 0;
      if (region && fx > 0) fxMap.set(region, fx);
    }
  }
  
  return fxMap;
}

function buildFullListUsd_() {
  SpreadsheetApp.getActive().toast('Converting to USD...', 'Build Market Data', 3);
  
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName('Full List');
  if (!src) { SpreadsheetApp.getActive().toast('Full List not found','Error',5); return; }
  const values = src.getDataRange().getValues();
  if (values.length < 2) { SpreadsheetApp.getActive().toast('Full List empty','Info',3); return; }
  const head = values[0].map(h => String(h || '').trim());
  const cRegion = head.indexOf('Region');
  const cP10  = head.indexOf('P10');
  const cP25  = head.indexOf('P25');
  const cP40  = head.indexOf('P40');
  const cP50  = head.indexOf('P50');
  const cP625 = head.indexOf('P62.5');
  const cP75  = head.indexOf('P75');
  const cP90  = head.indexOf('P90');
  const cRangeStart = head.indexOf('Range Start');
  const cRangeMid = head.indexOf('Range Mid');
  const cRangeEnd = head.indexOf('Range End');
  const cIMin = head.indexOf('Internal Min');
  const cIMed = head.indexOf('Internal Median');
  const cIMax = head.indexOf('Internal Max');
  // CR columns don't need FX conversion (they're ratios)
  const fxMap = _getFxMap_();

  const out = [head];
  for (let r=1; r<values.length; r++) {
    const row = values[r].slice();
    const region = String(row[cRegion] || '').trim();
    const fx = fxMap.get(region) || 1;
    const mul = (i) => { if (i >= 0) { const n = toNumber(row[i]); row[i] = isNaN(n) ? row[i] : n * fx; } };
    [cP10,cP25,cP40,cP50,cP625,cP75,cP90,cRangeStart,cRangeMid,cRangeEnd,cIMin,cIMed,cIMax].forEach(mul);
    // Round market percentiles to nearest hundred after FX conversion
    const r100 = (i) => { if (i >= 0) { const n = toNumber(row[i]); if (!isNaN(n)) row[i] = _round100_(n); } };
    [cP10,cP25,cP40,cP50,cP625,cP75,cP90,cRangeStart,cRangeMid,cRangeEnd].forEach(r100);
    out.push(row);
  }

  const dst = ss.getSheetByName('Full List USD') || ss.insertSheet('Full List USD');
  dst.setTabColor('#FF0000'); // Red color for automated sheets
  dst.clearContents();
  dst.getRange(1,1,out.length,head.length).setValues(out);
  dst.autoResizeColumns(1, head.length);
  SpreadsheetApp.getActive().toast('✅ Full List USD built\n⚡ Optimized (v4.6.0)', 'Complete', 5);
}

function exportProposedSalaryRanges_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Export Proposed Salary Ranges', 'Enter category (X0 or Y1). Default X0:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const category = String(resp.getResponseText() || 'X0').trim().toUpperCase();
  if (!/^(X0|Y1)$/.test(category)) { ui.alert('Invalid category. Use X0 or Y1.'); return; }
  // For brevity, reuse Full List logic then map columns per category. Users can use Full List for calculators; export remains optional.
  rebuildFullListTabs_(); ui.alert('Use the Full List sheet for calculators; export file creation trimmed in merged build.');
}

function buildHelpSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('About & Help') || ss.insertSheet('About & Help');
  sh.clearContents();
  const lines = [
    ['💰 Salary Range Calculator - Help & Getting Started'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['🎯 SIMPLIFIED WORKFLOW - 3 STEPS ONLY'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['📋 FIRST TIME SETUP'],
    [''],
    ['🏗️ STEP 1: Fresh Build (Create All Sheets)'],
    ['   Menu: 💰 Salary Ranges Calculator → 🏗️ Fresh Build (Create All Sheets)'],
    ['   What it creates:'],
    ['   ✓ Aon region tabs (India, US, UK) with headers'],
    ['   ✓ Mapping sheets (Lookup, Job family Descriptions, Employee Level Mapping, etc.)'],
    ['   ✓ Calculator UIs (Salary Ranges X0 and Y1)'],
    ['   ✓ Full List placeholders'],
    ['   Time: ~10 seconds'],
    [''],
    ['   After running Fresh Build:'],
    ['   → Paste your Aon market data into:'],
    ['      • Aon India - 2025'],
    ['      • Aon US - 2025'],
    ['      • Aon UK - 2025'],
    ['   → Configure HiBob API credentials:'],
    ['      Extensions → Apps Script → Project Settings → Script Properties'],
    ['      Add: BOB_ID and BOB_KEY (from HiBob service account)'],
    [''],
    ['📥 STEP 2: Import Bob Data'],
    ['   Menu: 💰 Salary Ranges Calculator → 📥 Import Bob Data'],
    ['   What it imports:'],
    ['   ✓ Base Data (employee list with salaries)'],
    ['   ✓ Bonus History (latest bonus/commission per employee)'],
    ['   ✓ Comp History (latest compensation change per employee)'],
    ['   ✓ Employees Mapped sheet is manually maintained (legacy)'],
    ['   ✓ Auto-syncs Title Mapping sheet (all unique job titles)'],
    ['   Time: 1-2 minutes'],
    [''],
    ['   After importing:'],
    ['   → Review "Employees Mapped" sheet (if using legacy mapping)'],
    ['   → Map each employee to:'],
    ['      • Aon Code (job family like EN.SODE, FI.FINA)'],
    ['      • Level (L2 IC through L9 Mgr)'],
    ['   → Review "Title Mapping" sheet'],
    ['   → Map job titles to Aon Codes'],
    [''],
    ['📊 STEP 3: Build Market Data'],
    ['   Menu: 💰 Salary Ranges Calculator → 📊 Build Market Data (Full Lists)'],
    ['   What it builds:'],
    ['   ✓ Full List (all X0/Y1 job family/level combinations, local currency)'],
    ['   ✓ Full List USD (USD converted for multi-region analysis)'],
    ['   ✓ Includes ALL eligible job families (not just mapped employees)'],
    ['   ✓ Adds internal stats (Min/Median/Max/Count) where employees exist'],
    ['   Time: 30-90 seconds'],
    [''],
    ['   You can now use the calculators!'],
    ['   → "Salary Ranges (X0)" - Engineering & Product'],
    ['   → "Salary Ranges (Y1)" - Everyone Else'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['🔄 REGULAR REFRESH WORKFLOW'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Weekly/Monthly Data Refresh:'],
    ['1) 📥 Import Bob Data (auto-syncs all employees with smart suggestions)'],
    ['2) ✅ Review Employee Mappings (approve new/changed mappings)'],
    ['3) 📊 Build Market Data (rebuilds Full Lists with CR values)'],
    [''],
    ['💡 Feedback Loop: Approved mappings auto-update Legacy Mappings for next import'],
    [''],
    ['After Aon Data Update:'],
    ['1) Paste new Aon data into region tabs'],
    ['2) 📊 Build Market Data (rebuild Full Lists)'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['📊 CALCULATORS'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Two Calculators Available:'],
    [''],
    ['🔧 Salary Ranges (X0) - Engineering & Product'],
    ['   • Range: P25 (Start) → P50 (Mid) → P90 (End)'],
    ['   • For: Engineering, Product, AI/ML roles'],
    ['   • Category fixed to X0'],
    [''],
    ['👥 Salary Ranges (Y1) - Everyone Else'],
    ['   • Range: P10 (Start) → P40 (Mid) → P62.5 (End)'],
    ['   • For: All other job families'],
    ['   • Category fixed to Y1'],
    [''],
    ['How to use:'],
    ['1. Select Job Family from dropdown'],
    ['2. Select Region (US, UK, India)'],
    ['3. View ranges for all levels (L2 IC through L9 Mgr)'],
    ['4. Compare market ranges vs internal stats'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['🔧 CUSTOM FUNCTIONS (For Formulas)'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Salary Range Functions:'],
    ['  =SALARY_RANGE(category, region, family, level)'],
    ['  =SALARY_RANGE_MIN(category, region, family, level)'],
    ['  =SALARY_RANGE_MID(category, region, family, level)'],
    ['  =SALARY_RANGE_MAX(category, region, family, level)'],
    [''],
    ['Examples:'],
    ['  =SALARY_RANGE_MIN("X0", "US", "EN.SODE", "L5 IC")  → Returns P25 for X0'],
    ['  =SALARY_RANGE_MID("Y1", "India", "FI.FINA", "L6 IC")  → Returns P40 for Y1'],
    [''],
    ['Aon Percentile Functions:'],
    ['  =AON_P10(region, family, level)'],
    ['  =AON_P25(region, family, level)'],
    ['  =AON_P40(region, family, level)'],
    ['  =AON_P50(region, family, level)'],
    ['  =AON_P625(region, family, level)'],
    ['  =AON_P75(region, family, level)'],
    ['  =AON_P90(region, family, level)'],
    [''],
    ['Internal Stats Function:'],
    ['  =INTERNAL_STATS(region, family, level)'],
    ['  Returns: [Min, Median, Max, Employee Count]'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['🗺️ MAPPING SHEETS'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Employees Mapped - Smart employee-to-Aon code mapping with approval workflow'],
    ['   Columns: Employee ID, Name, Title, Dept, Site, Aon Code, Level, Confidence, Source, Status, Salary, Start Date'],
    ['   Purpose: Map employees to job families and levels for CR calculations'],
    ['   Updated: Auto-synced during Import Bob Data (uses Legacy + Title mappings)'],
    ['   Workflow: Review → Approve → Auto-updates Legacy Mappings'],
    ['   Sources: Legacy (100%), Title-Based (95%), Manual (50%)'],
    [''],
    ['Job family Descriptions - Maps Aon codes to friendly names'],
    ['   Columns: Aon Code, Job Family (Exec Description)'],
    ['   Purpose: Maps EN.SODE → "Software Engineer", FI.FINA → "Finance Analyst"'],
    ['   Updated: Auto-populated from Aon data'],
    [''],
    ['Title Mapping - Maps job titles to Aon codes'],
    ['   Columns: Job title (live), Job title (Mapped), Job family'],
    ['   Purpose: Helps suggest mappings for employees'],
    ['   Updated: Auto-synced when you run "Import Bob Data"'],
    [''],
    ['Legacy Mappings - Historical employee mapping data (feedback loop)'],
    ['   Columns: Employee ID, Job Family, Full Mapping'],
    ['   Purpose: Stores approved mappings for future imports'],
    ['   Updated: Auto-updated from Employees Mapped (approved entries only)'],
    ['   Feedback Loop: Approved mappings → Legacy → Next import (100% confidence)'],
    [''],
    ['Lookup - Comprehensive mapping reference (single source of truth)'],
    ['   Section 1: CIQ Level → Aon Level mapping (L5 IC → P5)'],
    ['   Section 2: Region/Site → FX rates (US=1.0, UK=1.37, India=0.0125)'],
    ['   Section 3: Aon Code → Job Family + Category (71 codes)'],
    [''],
    ['DEPRECATED SHEETS (delete if present):'],
    ['   ❌ Job family Descriptions - Use Lookup instead'],
    ['   ❌ Employee Level Mapping - Use Employees Mapped instead'],
    ['   ❌ Aon Code Remap - Handled in code'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['🛠️ TOOLS MENU'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['💱 Apply Currency Format'],
    ['   Applies region-appropriate currency formatting ($, £, ₹)'],
    [''],
    ['🗑️ Clear All Caches'],
    ['   Clears cached data (use if calculator shows stale values)'],
    [''],
    ['📖 Generate Help Sheet'],
    ['   Creates/updates this help documentation'],
    [''],
    ['ℹ️ Quick Instructions'],
    ['   Shows quick-start modal dialog'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['📝 DATA FLOW'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['Aon Region Tabs → Job family Descriptions → Full List'],
    ['HiBob API → Base Data → Employees Mapped → Full List (internal stats)'],
    ['Full List → Calculators (via SALARY_RANGE functions)'],
    ['Full List → Full List USD (via FX rates)'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['🐛 TROUBLESHOOTING'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['"Prerequisites Missing" Error'],
    ['   → Run: 🏗️ Fresh Build'],
    ['   → Check Aon data is pasted in region tabs'],
    ['   → Check Lookup and Job family Descriptions have data'],
    [''],
    ['"Sheet not found" Error'],
    ['   → Run: 🏗️ Fresh Build'],
    [''],
    ['"Missing BOB_ID or BOB_KEY" Error'],
    ['   → Configure Script Properties (see Step 1 above)'],
    [''],
    ['Calculator Shows Old Data'],
    ['   → Tools → 🗑️ Clear All Caches'],
    ['   → Run: 📊 Build Market Data'],
    [''],
    ['Missing Mappings (red highlighting)'],
    ['   → Fill in "Employees Mapped" sheet'],
    ['   → Map Aon Code and Level for each employee'],
    ['   → Run: 📊 Build Market Data'],
    [''],
    ['Formula Returns Blank'],
    ['   → Check if Full List has data for that combination'],
    ['   → Check if job family is eligible for X0 or Y1'],
    ['   → Run: 📊 Build Market Data'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['💡 TIPS'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['• Fallback logic: If a percentile is missing, system uses next higher percentile'],
    ['  Example: P10 blank → uses P25 instead'],
    [''],
    ['• Full List includes ALL combinations for X0/Y1 families, not just mapped employees'],
    ['  This ensures you can use the calculator for any role, even if no employees currently exist'],
    [''],
    ['• Internal stats (Min/Median/Max/Count) only show where actual employees exist'],
    [''],
    ['• Caches expire after 10 minutes to ensure fresh data'],
    [''],
    ['• Half-levels (L5.5, L6.5) are calculated by averaging neighboring levels'],
    [''],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    ['📞 NEED MORE HELP?'],
    ['═══════════════════════════════════════════════════════════════════════════════'],
    [''],
    ['See: MENU_FUNCTIONS_GUIDE.md for detailed function descriptions'],
    ['Version: 3.4.0 - Simplified Workflow'],
    ['Last Updated: 2025-11-27']
  ];
  sh.getRange(1,1,lines.length,1).setValues(lines.map(r => [r[0]]));
  sh.setColumnWidth(1, 800);
}

/********************************
 * Aon code remapping support (central alias updates)
 ********************************/
function getAonCodeRemapMap_() {
  const cacheKey = 'REM:CODEMAP';
  const cached = _cacheGet_(cacheKey);
  if (cached) return new Map(cached);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Aon Code Remap');
  const map = new Map();
  // Built-in default: EN.SOML → EN.AIML
  map.set('EN.SOML', 'EN.AIML');
  if (sh) {
    const vals = sh.getDataRange().getValues();
    if (vals.length > 1) {
      const head = vals[0].map(h => String(h || '').trim().toLowerCase());
      const iFrom = head.findIndex(h => /from\s*code/.test(h) || /^from$/i.test(h));
      const iTo   = head.findIndex(h => /to\s*code/.test(h)   || /^to$/i.test(h));
      for (let r=1; r<vals.length; r++) {
        const from = iFrom>=0 ? String(vals[r][iFrom] || '').trim() : '';
        const to   = iTo>=0   ? String(vals[r][iTo]   || '').trim() : '';
        if (from && to) map.set(from, to);
      }
    }
  }
  _cachePut_(cacheKey, Array.from(map.entries()), CACHE_TTL);
  return map;
}

function remapAonCode_(code) {
  const c = String(code || '').trim(); if (!c) return c;
  const m = getAonCodeRemapMap_();
  return m.get(c) || c;
}

function reverseRemapAonCode_(code) {
  const c = String(code || '').trim(); if (!c) return c;
  const m = getAonCodeRemapMap_();
  for (const [from, to] of m.entries()) { if (to === c) return from; }
  return c;
}

function _getAonValueWithCodeFallback_(sheet, fam, targetNum, prefLetter, ciqBaseLevel, headerRegex) {
  const tries = [];
  const f0 = String(fam || '').trim();
  const f1 = remapAonCode_(f0);
  const f2 = reverseRemapAonCode_(f0);
  [f0, f1, f2].forEach(x => { if (x && !tries.includes(x)) tries.push(x); });
  for (const f of tries) {
    const v = getAonValueStrictByHeader_(sheet, f, targetNum, prefLetter, ciqBaseLevel, headerRegex);
    if (v !== '') return v;
  }
  return '';
}

function createAonPlaceholderSheets_() {
  const ss = SpreadsheetApp.getActive();
  const targets = [
    'Aon India - 2025',
    'Aon US - 2025',
    'Aon UK - 2025'
  ];
  const headers = [
    'Job Code',
    'Job Family',
    'Market \n\n (43) CFY Fixed Pay: 10th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 25th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 40th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 50th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 62.5th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 75th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 90th Percentile'
  ];
  targets.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.setTabColor('#FF0000'); // Red color for automated sheets
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, headers.length);
      // Format numeric columns (percentiles)
      const rows = Math.max(1000, sh.getMaxRows() - 1);
      sh.getRange(2, 3, rows, headers.length - 2).setNumberFormat('#,##0');
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Ensured Aon placeholder tabs exist (headers ready).', 'Done', 5);
}

/********************************
 * Category picker + UI wrappers
 ********************************/
function ensureCategoryPicker_() {
  const sh = uiSheet_();
  if (!sh) return;
  const cell = sh.getRange('B3');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['X0','Y1'], true) // Only 2 categories now
    .setAllowInvalid(false)
    .build();
  const currentRule = cell.getDataValidation();
  if (!currentRule || String(currentRule) !== String(rule)) cell.setDataValidation(rule);
  const v = String(cell.getDisplayValue() || '').trim();
  if (!v || v === 'X1') cell.setValue('X0'); // Default to X0, convert old X1 to X0
}

function ensureCategoryPickerY1_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(UI_SHEET_NAME_Y1);
  if (!sh) return;
  const cell = sh.getRange('B3');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y1'], true) // Y1 only for this calculator
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
  cell.setValue('Y1'); // Always Y1
}

function ensureRegionPicker_() {
  const sh = uiSheet_();
  if (!sh) return;
  const cell = sh.getRange('B4');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['US','India','UK'], true)
    .setAllowInvalid(false)
    .build();
  const current = cell.getDataValidation();
  if (!current || String(current) !== String(rule)) cell.setDataValidation(rule);
  const v = String(cell.getDisplayValue() || '').trim();
  if (!v) cell.setValue('US');
}

/**
 * Creates Job Family dropdown from Lookup sheet
 * Reads from Section 3 (Aon Code → Job Family mapping)
 */
function ensureExecFamilyPicker_() {
  const ss = SpreadsheetApp.getActive();
  const sh = uiSheet_(); if (!sh) return;
  
  // Try reading from Lookup sheet first (new format)
  let families = [];
  const lookupSh = ss.getSheetByName('Lookup');
  if (lookupSh) {
    const vals = lookupSh.getDataRange().getValues();
    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];
      if (!row || row.length < 2) continue;
      
      const col1 = String(row[0] || '').trim();
      const col2 = String(row[1] || '').trim();
      
      // Skip header rows
      if (col1 === 'Aon Code' || col1 === 'CIQ Level' || col1 === 'Region') continue;
      
      // If column 1 is an Aon code (contains dot) and column 2 has job family
      if (col1 && col1.includes('.') && col2) {
        families.push(col2);
      }
    }
  }
  
  // Fall back to Job family Descriptions sheet if Lookup doesn't have data
  if (families.length === 0) {
    const mapSh = ss.getSheetByName('Job family Descriptions');
    if (mapSh && mapSh.getLastRow() > 1) {
      const vals = mapSh.getRange(2, 2, mapSh.getLastRow() - 1, 1).getValues();
      families = vals.map(r => String(r[0] || '').trim()).filter(Boolean);
    }
  }
  
  if (families.length === 0) return; // No data found
  
  // Create dropdown with unique sorted values
  const uniq = Array.from(new Set(families)).sort();
  const cell = sh.getRange('B2');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniq, true)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(rule);
}

/**
 * Builds calculator UI for X0 (Engineering and Product)
 * Range: P25 → P62.5 → P90
 */
function buildCalculatorUI_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(UI_SHEET_NAME_X0);
  if (!sh) {
    sh = ss.insertSheet(UI_SHEET_NAME_X0);
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  
  // Get X0 families only
  const categoryMap = _getCategoryMap_();
  const execMap = _getExecDescMap_();
  const x0Families = [];
  categoryMap.forEach((cat, code) => {
    if (cat === 'X0') {
      const desc = execMap.get(code);
      if (desc) x0Families.push(desc);
    }
  });
  
  // Job Family dropdown (X0 families only)
  if (x0Families.length > 0) {
    const uniq = Array.from(new Set(x0Families)).sort();
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniq, true)
      .setAllowInvalid(false)
      .build();
    sh.getRange('B2').setDataValidation(rule);
  }

  // Labels (keeps existing styling; only writes text)
  sh.getRange('A2').setValue('Job Family');
  sh.getRange('A3').setValue('Region');
  sh.getRange('A4').setValue('Currency');

  // Region dropdown
  const regionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['US', 'India', 'UK'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B3').setDataValidation(regionRule);
  const currentRegion = sh.getRange('B3').getValue();
  if (!currentRegion) sh.getRange('B3').setValue('US');

  // Currency dropdown (Local/USD)
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Local', 'USD'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B4').setDataValidation(currencyRule);
  const currentCurrency = sh.getRange('B4').getValue();
  if (!currentCurrency) sh.getRange('B4').setValue('Local');

  // Header row - Market Range
  sh.getRange('A7').setValue('Level');
  sh.getRange('B7').setValue('Range Start');
  sh.getRange('C7').setValue('Range Mid');
  sh.getRange('D7').setValue('Range End');
  
  // Header row - Internal Range
  sh.getRange('F7').setValue('Min');
  sh.getRange('G7').setValue('Median');
  sh.getRange('H7').setValue('Max');
  sh.getRange('I7').setValue('Emp Count');
  sh.getRange('J7').setValue('Avg CR');
  sh.getRange('K7').setValue('TT CR');
  sh.getRange('L7').setValue('New Hire CR');
  sh.getRange('M7').setValue('BT CR');

  // Level list
  const levels = ['L2 IC','L3 IC','L4 IC','L5 IC','L5.5 IC','L6 IC','L6.5 IC','L7 IC','L4 Mgr','L5 Mgr','L5.5 Mgr','L6 Mgr','L6.5 Mgr','L7 Mgr','L8 Mgr','L9 Mgr'];
  sh.getRange(8,1,levels.length,1).setValues(levels.map(s=>[s]));

  // OPTIMIZED: Batch formula generation
  const formulasRangeStart = [], formulasRangeMid = [], formulasRangeEnd = [];
  const formulasIntMin = [], formulasIntMed = [], formulasIntMax = [], formulasIntCount = [];
  const formulasAvgCR = [], formulasTTCR = [], formulasNewHireCR = [], formulasBTCR = [];
  
  levels.forEach((level, i) => {
    const aRow = 8 + i;
    
    // Market Range: Currency-aware XLOOKUP (Column N=Range Start, O=Range Mid, P=Range End)
    // FIX: KEY is in Column Y, not U!
    formulasRangeStart.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$N:$N,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$N:$N,""))`]);
    formulasRangeMid.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$O:$O,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$O:$O,""))`]);
    formulasRangeEnd.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$P:$P,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$P:$P,""))`]);
    
    // Internal stats (Column Q=Internal Min, R=Median, S=Max, T=Emp Count)
    // Currency-aware: Switch between Full List (local) and Full List USD
    formulasIntMin.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$Q:$Q,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$Q:$Q,""))`]);
    formulasIntMed.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$R:$R,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$R:$R,""))`]);
    formulasIntMax.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$S:$S,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$S:$S,""))`]);
    formulasIntCount.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$T:$T,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$T:$T,""))`]);
    
    // Compa Ratio columns - XLOOKUP from Full List (pre-calculated)
    // Column Y = Key, Column U = Avg CR, Column V = TT CR, Column W = New Hire CR, Column X = BT CR
    formulasAvgCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$U:$U,"")`]);
    formulasTTCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$V:$V,"")`]);
    formulasNewHireCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$W:$W,"")`]);
    formulasBTCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$X:$X,"")`]);
  });
  
  // Batch set all formulas at once (single API call per column)
  sh.getRange(8, 2, levels.length, 1).setFormulas(formulasRangeStart);   // Column B: Range Start
  sh.getRange(8, 3, levels.length, 1).setFormulas(formulasRangeMid);     // Column C: Range Mid
  sh.getRange(8, 4, levels.length, 1).setFormulas(formulasRangeEnd);     // Column D: Range End
  sh.getRange(8, 6, levels.length, 1).setFormulas(formulasIntMin);       // Column F: Min
  sh.getRange(8, 7, levels.length, 1).setFormulas(formulasIntMed);       // Column G: Median
  sh.getRange(8, 8, levels.length, 1).setFormulas(formulasIntMax);       // Column H: Max
  sh.getRange(8, 9, levels.length, 1).setFormulas(formulasIntCount);     // Column I: Emp Count
  sh.getRange(8,10, levels.length, 1).setFormulas(formulasAvgCR);        // Column J: Avg CR
  sh.getRange(8,11, levels.length, 1).setFormulas(formulasTTCR);         // Column K: TT CR
  sh.getRange(8,12, levels.length, 1).setFormulas(formulasNewHireCR);    // Column L: New Hire CR
  sh.getRange(8,13, levels.length, 1).setFormulas(formulasBTCR);         // Column M: BT CR

  applyCurrency_();
  SpreadsheetApp.getActive().toast('Calculator UI built. Choose Region, Category, and Job Family to calculate.', 'Done', 5);
}

function UI_SALARY_RANGE(region, family, ciqLevel) {
  const sh = uiSheet_();
  const category = sh ? String(sh.getRange('B3').getDisplayValue() || 'X0').trim().toUpperCase() : 'X0';
  return SALARY_RANGE(category, region, family, ciqLevel);
}

function UI_SALARY_RANGE_MIN(region, family, ciqLevel) {
  const sh = uiSheet_();
  const category = sh ? String(sh.getRange('B3').getDisplayValue() || 'X0').trim().toUpperCase() : 'X0';
  return SALARY_RANGE_MIN(category, region, family, ciqLevel);
}

function UI_SALARY_RANGE_MID(region, family, ciqLevel) {
  const sh = uiSheet_();
  const category = sh ? String(sh.getRange('B3').getDisplayValue() || 'X0').trim().toUpperCase() : 'X0';
  return SALARY_RANGE_MID(category, region, family, ciqLevel);
}

function UI_SALARY_RANGE_MAX(region, family, ciqLevel) {
  const sh = uiSheet_();
  const category = sh ? String(sh.getRange('B3').getDisplayValue() || 'X0').trim().toUpperCase() : 'X0';
  return SALARY_RANGE_MAX(category, region, family, ciqLevel);
}

/********************************
 * Full List index (fast lookups for SALARY_RANGE)
 ********************************/
function _getFullListIndex_() {
  const cacheKey = 'FL:INDEX';
  const cached = _cacheGet_(cacheKey);
  if (cached) return cached; // already a plain object
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Full List');
  const index = {}; // key -> {p40,p50,p625,p75,p90}
  if (!sh) { _cachePut_(cacheKey, index, CACHE_TTL); return index; }
  const values = sh.getDataRange().getValues();
  if (!values.length) { _cachePut_(cacheKey, index, CACHE_TTL); return index; }
  const head = values[0].map(h => String(h || '').trim());
  const cExec = head.findIndex(h => /Job Family.*Exec/i.test(h));
  const cCIQ  = head.indexOf('CIQ Level');
  const cRegion = head.indexOf('Region');
  const cP10  = head.indexOf('P10');
  const cP25  = head.indexOf('P25');
  const cP40  = head.indexOf('P40');
  const cP50  = head.indexOf('P50');
  const cP625 = head.indexOf('P62.5');
  const cP75  = head.indexOf('P75');
  const cP90  = head.indexOf('P90');
  if ([cExec,cCIQ,cRegion,cP50,cP625,cP75].some(i => i < 0)) { _cachePut_(cacheKey, index, CACHE_TTL); return index; }
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const exec = String(row[cExec] || '').trim();
    const ciq  = String(row[cCIQ] || '').trim();
    const region = String(row[cRegion] || '').trim();
    if (!exec || !ciq || !region) continue;
    const key = `${exec}${ciq}${region}`;
    index[key] = {
      p10:  cP10  >= 0 ? toNumber_(row[cP10])  : NaN,
      p25:  cP25  >= 0 ? toNumber_(row[cP25])  : NaN,
      p40:  cP40  >= 0 ? toNumber_(row[cP40])  : NaN,
      p50:  cP50  >= 0 ? toNumber_(row[cP50])  : NaN,
      p625: cP625 >= 0 ? toNumber_(row[cP625]) : NaN,
      p75:  cP75  >= 0 ? toNumber_(row[cP75])  : NaN,
      p90:  cP90  >= 0 ? toNumber_(row[cP90])  : NaN
    };
  }
  _cachePut_(cacheKey, index, CACHE_TTL);
  return index;
}

function _familyToExecDesc_(familyOrCode) {
  const fam = String(familyOrCode || '').trim();
  if (!fam) return fam;
  const map = _getExecDescMap_();
  return map.get(fam) || fam;
}

function _codesForExecDesc_(desc) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Job family Descriptions');
  const out = [];
  if (!sh) return out;
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return out;
  const head = vals[0].map(h=>String(h||''));
  const iCode = head.findIndex(h => /^(Aon\s*Code|Job\s*Code)$/i.test(h));
  const iDesc = head.findIndex(h => /Job\s*Family\s*\(Exec\s*Description\)/i.test(h));
  const want = String(desc||'').trim().toUpperCase();
  for (let r=1;r<vals.length;r++) {
    const code = iCode>=0 ? String(vals[r][iCode]||'').trim().toUpperCase() : '';
    const d    = iDesc>=0 ? String(vals[r][iDesc]||'').trim().toUpperCase() : '';
    if (code && d === want) out.push(code);
  }
  return out;
}

function _isEngineeringOrAllowedTE_(familyOrCode) {
  const fam = String(familyOrCode || '').trim().toUpperCase();
  if (!fam) return false;
  if (/^EN\./.test(fam) || /^TE\.DADS/.test(fam) || /^TE\.DABD/.test(fam)) return true;
  // If it's an Exec Description, check mapped codes
  const codes = _codesForExecDesc_(fam);
  for (const c of codes) {
    if (/^EN\./.test(c) || /^TE\.DADS/.test(c) || /^TE\.DABD/.test(c)) return true;
  }
  return false;
}

/**
 * Determines category (X0 or Y1) for a given Aon code or family
 * Reads from Lookup sheet Category mapping
 * Falls back to legacy logic if not found
 */
function _effectiveCategoryForFamily_(familyOrCode) {
  const code = String(familyOrCode || '').trim();
  if (!code) return 'Y1';
  
  // Try reading from Lookup sheet category mapping
  const categoryMap = _getCategoryMap_();
  if (categoryMap.has(code)) {
    return categoryMap.get(code);
  }
  
  // Fall back to legacy logic: X0 for Engineering/Product, Y1 for others
  if (_isEngineeringOrAllowedTE_(code)) {
    return 'X0';
  }
  return 'Y1';
}

function _getRangeFromFullList_(category, region, family, ciqLevel) {
  const exec = _familyToExecDesc_(family);
  const ciq = String(ciqLevel || '').trim();
  const reg = String(region || '').trim();
  if (!exec || !ciq || !reg) return { min:'', mid:'', max:'' };
  const idx = _getFullListIndex_();
  const rec = idx[`${exec}${ciq}${reg}`];
  if (!rec) return { min:'', mid:'', max:'' };
  const pick = (cat) => {
    // Updated range definitions with fallback logic:
    // X0 (Engineering/Product): P25 → P62.5 → P90 (with fallbacks)
    // Y1 (Everyone Else): P10 → P40 → P62.5 (with fallbacks)
    if (cat === 'X0') {
      const min = rec.p25 || rec.p40 || rec.p50 || '';
      const mid = rec.p625 || rec.p75 || rec.p90 || '';
      const max = rec.p90 || '';
      return { min, mid, max };
    }
    if (cat === 'Y1') {
      const min = rec.p10 || rec.p25 || rec.p40 || '';
      const mid = rec.p40 || rec.p50 || rec.p625 || '';
      const max = rec.p625 || rec.p75 || rec.p90 || '';
      return { min, mid, max };
    }
    return { min:'', mid:'', max:'' };
  };
  const out = pick(String(category || '').trim().toUpperCase());
  const n = (v) => (v == null || isNaN(v)) ? '' : Number(v);
  return { min: n(out.min), mid: n(out.mid), max: n(out.max) };
}

/**
 * DEPRECATED: No longer creates deprecated sheets
 * - Job family Descriptions → Use Lookup sheet instead
 * - Employee Level Mapping → Use Employees Mapped instead
 * - Aon Code Remap → Not needed (handled in code)
 * - Title Mapping → Auto-populated during Import Bob Data
 */
function createMappingPlaceholderSheets_() {
  // No longer creates any sheets - all handled by other functions
  // This function kept for backward compatibility only
}

/********************************
 * DEPRECATED FUNCTIONS (kept for backward compatibility only)
 * These functions managed the old "Job family Descriptions" sheet
 * Now use Lookup sheet instead (auto-created with 71 codes)
 ********************************/

function listExecMappings_() {
  // DEPRECATED: Use Lookup sheet instead
  return [];
}

function upsertExecMapping_(code, desc) {
  // DEPRECATED: Lookup sheet is auto-managed
  SpreadsheetApp.getActive().toast('Use Lookup sheet instead (auto-managed)', 'Deprecated', 3);
}

function deleteExecMapping_(code) {
  // DEPRECATED: Lookup sheet is auto-managed
  SpreadsheetApp.getActive().toast('Use Lookup sheet instead (auto-managed)', 'Deprecated', 3);
}

function openExecMappingManager_() {
  // DEPRECATED: HTML UI no longer needed
  SpreadsheetApp.getUi().alert(
    '⚠️ Deprecated Feature',
    'This feature is deprecated.\n\n' +
    'Aon Code mappings are now managed in the Lookup sheet,\n' +
    'which is auto-created with all 71 job family codes.\n\n' +
    'No manual management needed!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function seedExecMappingsFromAon_() {
  // DEPRECATED: Aon sheets now include Job Family column
  // All mappings pre-loaded in Lookup sheet
  SpreadsheetApp.getActive().toast('Not needed - Lookup sheet has all 71 codes', 'Deprecated', 3);
}

function enhanceMappingSheets_() {
  const ss = SpreadsheetApp.getActive();
  // Employee Level Mapping
  (function(){
    const sh = ss.getSheetByName('Employee Level Mapping'); if (!sh) return;
    const head = sh.getRange(1,1,1,Math.max(3, sh.getLastColumn())).getValues()[0].map(h=>String(h||''));
    const colEmp = head.findIndex(h => /^Emp\s*ID/i.test(h)) + 1;
    const colMap = (head.indexOf('Mapping') >= 0 ? head.indexOf('Mapping') : head.findIndex(h => /Is\s*Mapped\?/i.test(h))) + 1;
    const colStatus = head.indexOf('Status') >= 0 ? head.indexOf('Status')+1 : (Math.max(colEmp,colMap)+1);
    // Headers
    sh.getRange(1,colStatus).setValue('Status');
    const empA = _colToLetter_(colEmp), mapA = _colToLetter_(colMap), statA = _colToLetter_(colStatus);
    // Status ARRAYFORMULA: blank when Emp ID blank; Missing when Mapping blank
    sh.getRange(2,colStatus).setFormula(`=ARRAYFORMULA(IF(LEN(${empA}2:${empA})=0,"",IF(LEN(${mapA}2:${mapA})=0,"Missing","")))`);
    // Conditional format only when Emp ID present and Mapping blank
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(`${mapA}2:${mapA}`);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(LEN($${empA}2)>0,LEN(${mapA}2)=0)`).setBackground('#FDE7E9').setFontColor('#D32F2F').setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
    // Missing count: only rows with Emp ID present
    sh.getRange(1, colStatus+1).setValue('Missing Count');
    const missA = _colToLetter_(colStatus+1);
    sh.getRange(2, colStatus+1).setFormula(`=COUNTIFS(${empA}2:${empA},"<>",${mapA}2:${mapA},"=")`);
  })();

  // Title Mapping
  (function(){
    const sh = ss.getSheetByName('Title Mapping'); if (!sh) return;
    const vals = sh.getDataRange().getValues(); if (!vals.length) return;
    const head = vals[0].map(h => String(h||''));
    let iLive = head.findIndex(h=>/Job\s*title\s*\(live\)/i.test(h)); if (iLive < 0) { iLive = head.findIndex(h=>/Job\s*title/i.test(h)); }
    let iFam = head.findIndex(h=>/Job\s*family/i.test(h)); if (iFam < 0) { iFam = 2; sh.getRange(1,iFam+1).setValue('Job family'); }
    let iStatus = head.indexOf('Status'); if (iStatus < 0) { iStatus = head.length; sh.getRange(1,iStatus+1).setValue('Status'); }
    const liveA = _colToLetter_(iLive+1), famA = _colToLetter_(iFam+1);
    sh.getRange(2,iStatus+1).setFormula(`=ARRAYFORMULA(IF(LEN(${liveA}2:${liveA})=0,"",IF(LEN(${famA}2:${famA})=0,"Missing","")))`);
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(`${famA}2:${famA}`);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(LEN($${liveA}2)>0,LEN(${famA}2)=0)`).setBackground('#FDE7E9').setFontColor('#D32F2F').setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
    sh.getRange(1, iStatus+2).setValue('Missing Count');
    sh.getRange(2, iStatus+2).setFormula(`=COUNTIFS(${liveA}2:${liveA},"<>",${famA}2:${famA},"=")`);
  })();

  SpreadsheetApp.getActive().toast('Mapping sheets enhanced', 'Done', 5);
}

function fillRegionFamilies_() {
  // DEPRECATED: Aon data now includes Job Family column
  // No need to populate from Job Code since source data has it
  SpreadsheetApp.getActive().toast('Not needed - Aon data includes Job Family column', 'Deprecated', 3);
}

function syncEmployeeLevelMappingFromBob_() {
  const ss = SpreadsheetApp.getActive();
  const base = ss.getSheetByName('Base Data');
  if (!base || base.getLastRow() <= 1) { SpreadsheetApp.getActive().toast('Base Data not found or empty','Info',4); return; }
  const vals = _getSheetDataCached_(base); // OPTIMIZED: Use cached data
  const head = vals[0].map(h => String(h||'').replace(/\s+/g,' ').trim());
  const cEmp   = head.findIndex(h => /^Emp\s*ID$/i.test(h) || /Employee\s*ID/i.test(h));
  const cAct   = head.findIndex(h => /^Active$/i.test(h) || /Active\s*\/\s*Inactive/i.test(h));
  const cTitle = head.findIndex(h => /^(Job\s*title|Job\s*Title|Title|Job\s*name)$/i.test(h));
  if (cEmp < 0 || cAct < 0) { SpreadsheetApp.getActive().toast('Base Data missing Emp ID/Active columns','Error',6); return; }

  const existingSh = ss.getSheetByName('Employee Level Mapping') || ss.insertSheet('Employee Level Mapping');
  const existingVals = existingSh.getDataRange().getValues();
  const existingMap = new Map();
  if (existingVals.length > 1) {
    const h = existingVals[0].map(x=>String(x||''));
    const eIdx = h.findIndex(x=>/^Emp\s*ID/i.test(x));
    let mIdx = h.findIndex(x=>/^Mapping$/i.test(x)); if (mIdx < 0) mIdx = h.findIndex(x=>/Is\s*Mapped\?/i.test(x));
    for (let r=1;r<existingVals.length;r++) {
      const id = eIdx>=0 ? String(existingVals[r][eIdx]||'').trim() : '';
      const map = mIdx>=0 ? String(existingVals[r][mIdx]||'').trim() : '';
      if (id) existingMap.set(id, map);
    }
  }

  const unique = new Map();
  for (let r=1;r<vals.length;r++) {
    const isActive = String(vals[r][cAct]||'').toLowerCase() === 'active';
    const id = String(vals[r][cEmp]||'').trim();
    if (!isActive || !id) continue;
    if (!unique.has(id)) unique.set(id, { title: cTitle>=0 ? String(vals[r][cTitle]||'').trim() : '' });
  }

  // Build suggestion via Title Mapping
  let suggByTitle = new Map();
  try { suggByTitle = buildTitleToFamilyMap_(ss); } catch (_) {}

  const outHead = ['Emp ID','Mapping','Status','Missing Count','Suggested'];
  const out = [outHead];
  const ids = Array.from(unique.keys()).sort();
  ids.forEach(id => {
    const mapping = existingMap.get(id) || '';
    const title = unique.get(id).title || '';
    const norm = (s) => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    const suggestion = title ? (suggByTitle.get(title) || suggByTitle.get(norm(title)) || '') : '';
    out.push([id, mapping, '', '', suggestion]);
  });

  existingSh.clearContents();
  existingSh.getRange(1,1,out.length,out[0].length).setValues(out);
  existingSh.setFrozenRows(1);
  // Status formula only for data rows
  if (ids.length) {
    existingSh.getRange(2,3,ids.length,1).setFormulaR1C1('=IF(LEN(RC[-1])=0,"Missing","")');
    existingSh.getRange(2,4,1,1).setFormula(`=COUNTIF(C2:C${ids.length+1},"Missing")`);
    // Conditional format for Mapping blank only within data rows
    const rules = existingSh.getConditionalFormatRules();
    const rng = existingSh.getRange(2,2,ids.length,1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=LEN(${rng.getA1Notation().split(':')[0]})=0`).setBackground('#FDE7E9').setFontColor('#D32F2F').setRanges([rng]).build());
    existingSh.setConditionalFormatRules(rules);
    existingSh.autoResizeColumns(1, outHead.length);
  }
  SpreadsheetApp.getActive().toast(`Employee Level Mapping synced: ${ids.length} active employees`, 'Done', 5);
}

function syncTitleMappingFromBob_() {
  const ss = SpreadsheetApp.getActive();
  const base = ss.getSheetByName('Base Data');
  if (!base || base.getLastRow() <= 1) { SpreadsheetApp.getActive().toast('Base Data not found or empty','Info',4); return; }
  const vals = _getSheetDataCached_(base); // OPTIMIZED: Use cached data
  const head = vals[0].map(h => String(h||'').replace(/\s+/g,' ').trim());
  const cTitle = head.findIndex(h => /^(Job\s*title|Job\s*Title|Title|Job\s*name)$/i.test(h));
  const cAct   = head.findIndex(h => /^Active$/i.test(h) || /Active\s*\/\s*Inactive/i.test(h));
  if (cTitle < 0 || cAct < 0) { SpreadsheetApp.getActive().toast('Base Data missing Title/Active columns','Error',6); return; }

  const liveSet = new Set();
  for (let r=1;r<vals.length;r++) {
    if (String(vals[r][cAct]||'').toLowerCase() !== 'active') continue;
    const t = String(vals[r][cTitle]||'').trim(); if (t) liveSet.add(t);
  }

  const sh = ss.getSheetByName('Title Mapping') || ss.insertSheet('Title Mapping');
  const existing = sh.getDataRange().getValues();
  const head2 = existing.length ? existing[0].map(h=>String(h||'')) : [];
  let iLive = head2.findIndex(h=>/Job\s*title\s*\(live\)/i.test(h)); if (iLive < 0) iLive = 0;
  let iFam  = head2.findIndex(h=>/Job\s*family/i.test(h)); if (iFam < 0) iFam = 2;
  // Build existing map
  const have = new Set();
  for (let r=1;r<existing.length;r++) { const v = String(existing[r][iLive]||'').trim(); if (v) have.add(v); }
  const toAppend = Array.from(liveSet).filter(t => !have.has(t)).map(t => [t, '', '']);
  if (existing.length === 0) sh.getRange(1,1,1,3).setValues([[ 'Job title (live)','Job title (Mapped)','Job family' ]]);
  if (toAppend.length) sh.getRange(sh.getLastRow()+1, 1, toAppend.length, 3).setValues(toAppend);
  // Enhance missing formatting/counts
  enhanceMappingSheets_();
  SpreadsheetApp.getActive().toast(`Title Mapping synced: +${toAppend.length} titles`, 'Done', 5);
}

// ============================================================================
// SIMPLIFIED COMBINED FUNCTIONS
// ============================================================================

/**
 * Syncs ALL Bob-based mappings (Employee Level + Title Mapping)
 * Combines syncEmployeeLevelMappingFromBob_ + syncTitleMappingFromBob_
 */
function syncAllBobMappings_() {
  SpreadsheetApp.getActive().toast('Syncing all Bob mappings...', 'In Progress', 3);
  syncEmployeeLevelMappingFromBob_();
  syncTitleMappingFromBob_();
  SpreadsheetApp.getActive().toast('All Bob mappings synced!', 'Complete', 5);
}

/**
 * Seeds ALL job family mappings (Exec Mappings + Job Family Fill)
 * Combines seedExecMappingsFromAon_ + fillRegionFamilies_
 */
function seedAllJobFamilyMappings_() {
  // DEPRECATED: Use Fresh Build → Import Bob Data workflow instead
  SpreadsheetApp.getActive().toast('Deprecated - Use Fresh Build instead', 'Deprecated', 3);
}

/**
 * QUICK SETUP - Initializes entire system in correct order
 * Run this ONCE after pasting Aon data into region tabs
 * 
 * Steps performed:
 * 1. Create all necessary tabs (Aon, Mapping, Calculator)
 * 2. Seed exec mappings from Aon data
 * 3. Fill job families in region tabs
 * 4. Build calculator UI with dropdowns
 * 5. Generate help documentation
 * 6. Enhance mapping sheets with formatting
 */
function quickSetup_() {
  // DEPRECATED: Use the new 3-step workflow instead
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '⚠️ Deprecated Feature',
    'Quick Setup has been replaced with a streamlined 3-step workflow:\n\n' +
    '1️⃣ Fresh Build - Creates all sheets and structure\n' +
    '2️⃣ Import Bob Data - Loads employee data with smart mapping\n' +
    '3️⃣ Build Market Data - Generates Full Lists and calculators\n\n' +
    'Use Menu → Fresh Build to start!',
    ui.ButtonSet.OK
  );
}

/**
 * Validates prerequisites before building Full List
 * Returns {valid: boolean, errors: string[]}
 */
function validatePrerequisites_() {
  const ss = SpreadsheetApp.getActive();
  const errors = [];
  
  // Check Aon region tabs
  const regions = ['Aon US - 2025', 'Aon UK - 2025', 'Aon India - 2025'];
  regions.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() <= 1) {
      errors.push(`❌ ${name} tab missing or empty`);
    }
  });
  
  // Check mapping tabs
  const mappings = ['Lookup', 'Job family Descriptions'];
  mappings.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() <= 1) {
      errors.push(`❌ ${name} tab missing or empty`);
    }
  });
  
  // Check Bob credentials
  const bobId = PropertiesService.getScriptProperties().getProperty('BOB_ID');
  const bobKey = PropertiesService.getScriptProperties().getProperty('BOB_KEY');
  if (!bobId || !bobKey) {
    errors.push('⚠️ HiBob API credentials not configured (Script Properties)');
  }
  
  return { valid: errors.length === 0, errors: errors };
}

/**
 * Validates and then rebuilds Full List with prerequisite checks
 */
function rebuildFullListTabsWithValidation_() {
  const validation = validatePrerequisites_();
  
  if (!validation.valid) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '⚠️ Prerequisites Missing',
      'Cannot rebuild Full List:\n\n' + validation.errors.join('\n') + '\n\n' +
      'Run Quick Setup first if this is initial setup.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  rebuildFullListTabs_();
}

/********************************
 * ========================================
 * HELPER FUNCTIONS FOR SIMPLIFIED WORKFLOW
 * ========================================
 ********************************/

/**
 * Creates Legacy Mappings sheet with historical employee mappings
 * Data is stored in Script Properties (persistent storage)
 * Sheet is a "view" of the persistent data
 */
function createLegacyMappingsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAMES.LEGACY_MAPPINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.LEGACY_MAPPINGS);
  }
  sh.setTabColor('#808080'); // Gray for reference data
  
  // Clear and set headers
  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([['Employee ID', 'Job Family (Base)', 'Full Mapping']]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#757575').setFontColor('#FFFFFF');
  
  // Try to load from Script Properties first (persistent storage)
  let legacyData = _loadLegacyMappingsFromStorage_();
  
  // If no data in storage, use embedded data as initial seed
  if (!legacyData || legacyData.length === 0) {
    legacyData = _getLegacyMappingData_();
    // Save to storage for future use
    if (legacyData.length > 0) {
      _saveLegacyMappingsToStorage_(legacyData);
    }
  }
  
  // Populate sheet
  if (legacyData.length > 0) {
    sh.getRange(2,1,legacyData.length,3).setValues(legacyData);
    sh.autoResizeColumns(1,3);
    SpreadsheetApp.getActive().toast(`Loaded ${legacyData.length} legacy mappings from storage`, 'Legacy Mappings', 5);
  } else {
    sh.autoResizeColumns(1,3);
  }
}

/**
 * Saves legacy mappings to Script Properties (persistent storage)
 * Survives sheet deletion and Fresh Build
 */
function _saveLegacyMappingsToStorage_(legacyData) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    // Convert array to JSON string
    const jsonData = JSON.stringify(legacyData);
    
    // Script Properties has a 9KB limit per key, so we might need to chunk large data
    const maxChunkSize = 8000; // 8KB chunks to be safe
    const chunks = [];
    
    for (let i = 0; i < jsonData.length; i += maxChunkSize) {
      chunks.push(jsonData.substring(i, i + maxChunkSize));
    }
    
    // Save chunk count
    scriptProperties.setProperty('LEGACY_MAPPINGS_CHUNKS', chunks.length.toString());
    
    // Save each chunk
    chunks.forEach((chunk, idx) => {
      scriptProperties.setProperty(`LEGACY_MAPPINGS_${idx}`, chunk);
    });
    
    // Save timestamp
    scriptProperties.setProperty('LEGACY_MAPPINGS_UPDATED', new Date().toISOString());
    
    Logger.log(`Saved ${legacyData.length} legacy mappings to storage (${chunks.length} chunks)`);
  } catch (e) {
    Logger.log(`Error saving legacy mappings to storage: ${e.message}`);
    SpreadsheetApp.getActive().toast('Warning: Could not save to persistent storage', 'Storage Error', 5);
  }
}

/**
 * Loads legacy mappings from Script Properties (persistent storage)
 * Returns array of [EmpID, JobFamily, FullMapping] or null if not found
 */
function _loadLegacyMappingsFromStorage_() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const chunkCount = parseInt(scriptProperties.getProperty('LEGACY_MAPPINGS_CHUNKS') || '0');
    
    if (chunkCount === 0) {
      Logger.log('No legacy mappings found in storage');
      return null;
    }
    
    // Reconstruct JSON from chunks
    let jsonData = '';
    for (let i = 0; i < chunkCount; i++) {
      const chunk = scriptProperties.getProperty(`LEGACY_MAPPINGS_${i}`);
      if (!chunk) {
        Logger.log(`Missing chunk ${i}, storage may be corrupted`);
        return null;
      }
      jsonData += chunk;
    }
    
    // Parse JSON
    const legacyData = JSON.parse(jsonData);
    const lastUpdated = scriptProperties.getProperty('LEGACY_MAPPINGS_UPDATED');
    
    Logger.log(`Loaded ${legacyData.length} legacy mappings from storage (last updated: ${lastUpdated})`);
    return legacyData;
  } catch (e) {
    Logger.log(`Error loading legacy mappings from storage: ${e.message}`);
    return null;
  }
}

/**
/**
 * Returns legacy mapping data (Employee ID | Job Family | Full Mapping)
 * This data was provided by the user for historical employee mappings
 * Updated 2025-11-27: Comprehensive dataset with EN.SOML → EN.AIML replacement
 */
function _getLegacyMappingData_() {
  // Legacy data: EmpID → Job Family → Full Mapping
  const mappings = {
    '20033': ['HR.ARIS', 'HR.ARIS.P5'],
    '20037': ['CS.RSTS', 'CS.RSTS.M5'],
    '20039': ['HR.GLMF', 'HR.GLMF.E1'],
    '20052': ['EN.SODE', 'EN.SODE.M4'],
    '20072': ['EN.SODE', 'EN.SODE.M6'],
    '20077': ['EN.SODE', 'EN.SODE.P6'],
    '20079': ['EN.0000', 'EN.0000.E3'],
    '20126': ['EN.SODE', 'EN.SODE.P6'],
    '20139': ['EN.SODE', 'EN.SODE.M4'],
    '20146': ['EN.SODE', 'EN.SODE.P5'],
    '20148': ['EN.SODE', 'EN.SODE.P5'],
    '20150': ['EN.SODE', 'EN.SODE.P5'],
    '20151': ['EN.SODE', 'EN.SODE.M4'],
    '20153': ['EN.PMPD', 'EN.PMPD.P5'],
    '20157': ['EN.PGPG', 'EN.PGPG.P5'],
    '20158': ['EN.SODE', 'EN.SODE.P4'],
    '20160': ['EN.UUUD', 'EN.UUUD.P5'],
    '20163': ['EN.SODE', 'EN.SODE.P4'],
    '20167': ['EN.SODE', 'EN.SODE.P4'],
    '20171': ['EN.SODE', 'EN.SODE.P5'],
    '20173': ['EN.SODE', 'EN.SODE.P4'],
    '20174': ['EN.SODE', 'EN.SODE.P5'],
    '20175': ['EN.SODE', 'EN.SODE.M5'],
    '20178': ['TE.DADA', 'TE.DADA.M4'],
    '20181': ['TE.DADS', 'TE.DADS.M5'],
    '20185': ['EN.SODE', 'EN.SODE.P5'],
    '20188': ['EN.SODE', 'EN.SODE.M4'],
    '20189': ['SA.CRCS', 'SA.CRCS.M5'],
    '20190': ['EN.PGPG', 'EN.PGPG.M5'],
    '20193': ['TE.DADA', 'TE.DADA.P4'],
    '20194': ['EN.SODE', 'EN.SODE.P4'],
    '20195': ['TE.DADS', 'TE.DADS.R6'],
    '20199': ['EN.PGPG', 'EN.PGPG.P5'],
    '20201': ['EN.SODE', 'EN.SODE.P4'],
    '20202': ['EN.SODE', 'EN.SODE.P5'],
    '20204': ['EN.SODE', 'EN.SODE.M5'],
    '20206': ['TE.DADA', 'TE.DADA.M4'],
    '20209': ['EN.SODE', 'EN.SODE.P4'],
    '20210': ['TE.DADA', 'TE.DADA.P4'],
    '20211': ['EN.PGPG', 'EN.PGPG.P4'],
    '20213': ['EN.SODE', 'EN.SODE.P4'],
    '20214': ['EN.SODE', 'EN.SODE.P5'],
    '20215': ['EN.SODE', 'EN.SODE.P5'],
    '20216': ['EN.SODE', 'EN.SODE.P4'],
    '20221': ['TE.DADS', 'TE.DADS.P4'],
    '20222': ['CS.RSTS', 'CS.RSTS.P3'],
    '20223': ['CS.RSTS', 'CS.RSTS.M4'],
    '20224': ['EN.PGPG', 'EN.PGPG.M5'],
    '20225': ['EN.SODE', 'EN.SODE.P4'],
    '20226': ['TE.DADS', 'TE.DADS.P4'],
    '20229': ['EN.SODE', 'EN.SODE.M4'],
    '20230': ['EN.PGPG', 'EN.PGPG.P5'],
    '20233': ['EN.SODE', 'EN.SODE.P5'],
    '20234': ['EN.SODE', 'EN.SODE.P4'],
    '20239': ['TE.DADA', 'TE.DADA.P3'],
    '20242': ['EN.SODE', 'EN.SODE.P3'],
    '20243': ['EN.SODE', 'EN.SODE.P5'],
    '20244': ['TE.DADA', 'TE.DADA.P3'],
    '20245': ['HR.TMTA', 'HR.TMTA.P4'],
    '20246': ['HR.GLGL', 'HR.GLGL.P5'],
    '20248': ['TE.DADA', 'TE.DADA.P3'],
    '20250': ['EN.PGPG', 'EN.PGPG.P5'],
    '20251': ['TE.DADS', 'TE.DADS.P3'],
    '20252': ['EN.SODE', 'EN.SODE.M5'],
    '20253': ['HR.TMTA', 'HR.TMTA.P6'],
    '20254': ['TE.DADS', 'TE.DADS.P4'],
    '20255': ['EN.UUUD', 'EN.UUUD.P4'],
    '20256': ['EN.SODE', 'EN.SODE.M5'],
    '20257': ['EN.SODE', 'EN.SODE.P4'],
    '20258': ['TE.DADS', 'TE.DADS.P5'],
    '20259': ['TE.DADS', 'TE.DADS.P6'],
    '20260': ['EN.PGPG', 'EN.PGPG.P5'],
    '20263': ['EN.SODE', 'EN.SODE.P4'],
    '20264': ['CS.RSTS', 'CS.RSTS.P4'],
    '20267': ['EN.SODE', 'EN.SODE.P3'],
    '20269': ['EN.SODE', 'EN.SODE.P4'],
    '20270': ['EN.SODE', 'EN.SODE.P3'],
    '20272': ['EN.PMPD', 'EN.PMPD.M5'],
    '20273': ['EN.SODE', 'EN.SODE.P4'],
    '20276': ['EN.SODE', 'EN.SODE.P3'],
    '20277': ['EN.UUUD', 'EN.UUUD.P4'],
    '20280': ['EN.SODE', 'EN.SODE.P4'],
    '20281': ['EN.SODE', 'EN.SODE.P3'],
    '20282': ['EN.SODE', 'EN.SODE.P3'],
    '20284': ['EN.SODE', 'EN.SODE.P4'],
    '20286': ['EN.SODE', 'EN.SODE.P4'],
    '20287': ['EN.SODE', 'EN.SODE.M4'],
    '20288': ['EN.SODE', 'EN.SODE.P3'],
    '20289': ['EN.SODE', 'EN.SODE.P4'],
    '20290': ['EN.SODE', 'EN.SODE.P4'],
    '20291': ['EN.PGPG', 'EN.PGPG.P4'],
    '20292': ['EN.SODE', 'EN.SODE.P3'],
    '20293': ['TE.DADA', 'TE.DADA.P3'],
    '20294': ['TE.DADA', 'TE.DADA.P3'],
    '20295': ['EN.SODE', 'EN.SODE.P4'],
    '20296': ['EN.SODE', 'EN.SODE.P3'],
    '20297': ['EN.SODE', 'EN.SODE.P4'],
    '20298': ['EN.SODE', 'EN.SODE.P4'],
    '20299': ['EN.SODE', 'EN.SODE.P5'],
    '20300': ['EN.SODE', 'EN.SODE.P4'],
    '20302': ['EN.SODE', 'EN.SODE.P3'],
    '20303': ['EN.SODE', 'EN.SODE.P3'],
    '20305': ['TE.DADS', 'TE.DADS.M5'],
    '20306': ['EN.SODE', 'EN.SODE.P4'],
    '20308': ['EN.SODE', 'EN.SODE.P5'],
    '20309': ['EN.SODE', 'EN.SODE.P3'],
    '20310': ['EN.SODE', 'EN.SODE.P4'],
    '20311': ['EN.PGPG', 'EN.PGPG.P5'],
    '20313': ['EN.PMPD', 'EN.PMPD.P5'],
    '20315': ['EN.SODE', 'EN.SODE.P4'],
    '20316': ['EN.SODE', 'EN.SODE.P4'],
    '20318': ['EN.SODE', 'EN.SODE.P5'],
    '20319': ['SA.CRCS', 'SA.CRCS.P4'],
    '20320': ['EN.SODE', 'EN.SODE.P3'],
    '20321': ['CS.RSTS', 'CS.RSTS.P4'],
    '20322': ['EN.SODE', 'EN.SODE.P4'],
    '20323': ['EN.SODE', 'EN.SODE.P4'],
    '20326': ['EN.SODE', 'EN.SODE.P4'],
    '20327': ['EN.SODE', 'EN.SODE.M6'],
    '20328': ['SA.CRCS', 'SA.CRCS.P4'],
    '20330': ['EN.SODE', 'EN.SODE.P4'],
    '20332': ['SA.CRCS', 'SA.CRCS.P4'],
    '20333': ['EN.SODE', 'EN.SODE.P4'],
    '20334': ['MK.PIPM', 'MK.PIPM.M5'],
    '20335': ['SA.CRCS', 'SA.CRCS.P5'],
    '20336': ['SA.CRCS', 'SA.CRCS.P4'],
    '20337': ['SA.CRCS', 'SA.CRCS.P5'],
    '20338': ['EN.PMPD', 'EN.PMPD.P5'],
    '20340': ['CS.RSTS', 'CS.RSTS.P4'],
    '20341': ['CS.RSTS', 'CS.RSTS.M5'],
    '20342': ['SA.CRCS', 'SA.CRCS.P5'],
    '20343': ['EN.SODE', 'EN.SODE.P4'],
    '20344': ['EN.SODE', 'EN.SODE.P4'],
    '20345': ['EN.SODE', 'EN.SODE.P4'],
    '20346': ['EN.SODE', 'EN.SODE.M4'],
    '20347': ['EN.SODE', 'EN.SODE.M5'],
    '20350': ['MK.PIPM', 'MK.PIPM.P3'],
    '20351': ['SP.BDBD', 'SP.BDBD.P5'],
    '20352': ['SA.CRCS', 'SA.CRCS.P5'],
    '20353': ['EN.SODE', 'EN.SODE.M6'],
    '20354': ['HR.GLGL', 'HR.GLGL.M5'],
    '20355': ['EN.SODE', 'EN.SODE.P5'],
    '20356': ['SP.BDBD', 'SP.BDBD.P4'],
    '20357': ['EN.DODO', 'EN.DODO.M5'],
    '20358': ['TE.DABD', 'TE.DABD.P4'],
    '20359': ['TE.DABD', 'TE.DABD.P4'],
    '20360': ['EN.SODE', 'EN.SODE.P3'],
    '20361': ['EN.SODE', 'EN.SODE.P4'],
    '20362': ['EN.SODE', 'EN.SODE.P5'],
    '20363': ['EN.SODE', 'EN.SODE.P5'],
    '20364': ['EN.SODE', 'EN.SODE.P4'],
    '20367': ['MK.CIDB', 'MK.CIDB.P4'],
    '20368': ['MK.PIPM', 'MK.PIPM.M5'],
    '20370': ['SA.CRCS', 'SA.CRCS.P4'],
    '20371': ['SA.CRCS', 'SA.CRCS.M5'],
    '20372': ['FI.OPMF', 'FI.OPMF.M5'],
    '20375': ['EN.SODE', 'EN.SODE.P4'],
    '20377': ['SA.CRCS', 'SA.CRCS.P5'],
    '20378': ['FI.ACGA', 'FI.ACGA.P4'],
    '20379': ['MK.PIPM', 'MK.PIPM.P5'],
    '20380': ['EN.SODE', 'EN.SODE.P4'],
    '20381': ['CS.RSTS', 'CS.RSTS.P3'],
    '20382': ['EN.SODE', 'EN.SODE.P3'],
    '20383': ['CS.RSTS', 'CS.RSTS.M3'],
    '20384': ['EN.DODO', 'EN.DODO.P3'],
    '20385': ['EN.SODE', 'EN.SODE.P4'],
    '20386': ['EN.SODE', 'EN.SODE.P4'],
    '20387': ['EN.0000', 'EN.0000.E3'],
    '20388': ['SA.CRCS', 'SA.CRCS.P4'],
    '20389': ['CS.RSTS', 'CS.RSTS.P3'],
    '20390': ['EN.SODE', 'EN.SODE.P3'],
    '20391': ['EN.PGPG', 'EN.PGPG.P4'],
    '20392': ['EN.SODE', 'EN.SODE.P3'],
    '20393': ['EN.SODE', 'EN.SODE.P3'],
    '20394': ['EN.SODE', 'EN.SODE.P3'],
    '20395': ['EN.SODE', 'EN.SODE.P4'],
    '20397': ['EN.SODE', 'EN.SODE.P3'],
    '20398': ['SA.CRCS', 'SA.CRCS.P5'],
    '20399': ['MK.PIPM', 'MK.PIPM.M4'],
    '20400': ['CS.RSTS', 'CS.RSTS.P3'],
    '20401': ['FI.ACGA', 'FI.ACGA.P4'],
    '20402': ['EN.SODE', 'EN.SODE.P3'],
    '20403': ['EN.DODO', 'EN.DODO.P5'],
    '20404': ['EN.SODE', 'EN.SODE.P4'],
    '20405': ['SA.CRCS', 'SA.CRCS.P4'],
    '20406': ['EN.UUUD', 'EN.UUUD.P5'],
    '20407': ['EN.SODE', 'EN.SODE.P4'],
    '20408': ['EN.SODE', 'EN.SODE.P3'],
    '20409': ['EN.SODE', 'EN.SODE.P3'],
    '20410': ['EN.SODE', 'EN.SODE.P3'],
    '20411': ['EN.SODE', 'EN.SODE.P4'],
    '20412': ['EN.SODE', 'EN.SODE.P5'],
    '20413': ['EN.SODE', 'EN.SODE.P4'],
    '20414': ['EN.SODE', 'EN.SODE.P4'],
    '20415': ['EN.SODE', 'EN.SODE.P5'],
    '20416': ['SA.CRCS', 'SA.CRCS.P4'],
    '20417': ['EN.SODE', 'EN.SODE.P4'],
    '20418': ['CS.RSTS', 'CS.RSTS.M4'],
    '20419': ['EN.SODE', 'EN.SODE.P3'],
    '20420': ['EN.SODE', 'EN.SODE.P4'],
    '20421': ['EN.SODE', 'EN.SODE.P3'],
    '20422': ['SA.CRCS', 'SA.CRCS.P4'],
    '20423': ['EN.DVEX', 'EN.DVEX.E1'],
    '20424': ['EN.PGPG', 'EN.PGPG.P5'],
    '20425': ['SA.CRCS', 'SA.CRCS.P3'],
    '20426': ['EN.SODE', 'EN.SODE.M4'],
    '20427': ['EN.SODE', 'EN.SODE.P3'],
    '20428': ['EN.0000', 'EN.0000.E1'],
    '20429': ['TE.DABD', 'TE.DABD.P4'],
    '20430': ['SA.CRCS', 'SA.CRCS.P4'],
    '20431': ['SA.CRCS', 'SA.CRCS.P4'],
    '20432': ['SA.CRCS', 'SA.CRCS.M4'],
    '20433': ['EN.SODE', 'EN.SODE.M4'],
    '20434': ['EN.SODE', 'EN.SODE.P4'],
    '20435': ['EN.DODO', 'EN.DODO.P4'],
    '20436': ['SA.CRCS', 'SA.CRCS.P4'],
    '20437': ['SA.CRCS', 'SA.CRCS.P4'],
    '20438': ['EN.SODE', 'EN.SODE.M6'],
    '20439': ['EN.DODO', 'EN.DODO.P4'],
    '20440': ['EN.PGPG', 'EN.PGPG.P4'],
    '20441': ['TE.DADA', 'TE.DADA.P4'],
    '20442': ['TE.DADS', 'TE.DADS.P3'],
    '20443': ['SA.CRCS', 'SA.CRCS.P4'],
    '20444': ['SA.CRCS', 'SA.CRCS.P4'],
    '20445': ['SA.CRCS', 'SA.CRCS.P5'],
    '20446': ['SP.BDBD', 'SP.BDBD.P3'],
    '20447': ['EN.SODE', 'EN.SODE.P4'],
    '20448': ['EN.SODE', 'EN.SODE.P5'],
    '20449': ['EN.SODE', 'EN.SODE.P3'],
    '20451': ['CS.RSTS', 'CS.RSTS.M5'],
    '20452': ['EN.SODE', 'EN.SODE.P3'],
    '20453': ['EN.SODE', 'EN.SODE.P3'],
    '20454': ['EN.SODE', 'EN.SODE.P3'],
    '20455': ['TE.DADA', 'TE.DADA.P3'],
    '20456': ['SA.CRCS', 'SA.CRCS.P4'],
    '20457': ['TE.DADS', 'TE.DADS.P3'],
    '20458': ['EN.SODE', 'EN.SODE.P5'],
    '20459': ['EN.SODE', 'EN.SODE.P4'],
    '20460': ['EN.SODE', 'EN.SODE.P4'],
    '20461': ['SA.CRCS', 'SA.CRCS.P4'],
    '20462': ['EN.SODE', 'EN.SODE.P3'],
    '20463': ['EN.SODE', 'EN.SODE.P3'],
    '20464': ['SP.BDBD', 'SP.BDBD.P3'],
    '20465': ['FI.ACGA', 'FI.ACGA.P4'],
    '20466': ['CS.RSTS', 'CS.RSTS.P4'],
    '20467': ['TE.DADS', 'TE.DADS.P4'],
    '20468': ['EN.SODE', 'EN.SODE.P3'],
    '20469': ['SA.CRCS', 'SA.CRCS.P4'],
    '20470': ['EN.SODE', 'EN.SODE.P3'],
    '20471': ['TE.DADA', 'TE.DADA.P4'],
    '20472': ['TE.DADA', 'TE.DADA.P3'],
    '20473': ['SA.CRCS', 'SA.CRCS.P4'],
    '20474': ['SA.CRCS', 'SA.CRCS.P5'],
    '20475': ['CS.RSTS', 'CS.RSTS.P3'],
    '20476': ['CS.RSTS', 'CS.RSTS.P3'],
    '20477': ['SA.CRCS', 'SA.CRCS.P4'],
    '20478': ['TE.DADA', 'TE.DADA.P3'],
    '20479': ['SA.CRCS', 'SA.CRCS.P5'],
    '20480': ['EN.SODE', 'EN.SODE.P3'],
    '20481': ['EN.SODE', 'EN.SODE.P3'],
    '20482': ['EN.SODE', 'EN.SODE.P3'],
    '20483': ['TE.DADS', 'TE.DADS.P3'],
    '20484': ['SA.CRCS', 'SA.CRCS.P4'],
    '20485': ['EN.SODE', 'EN.SODE.P3'],
    '20486': ['TE.DADS', 'TE.DADS.P3'],
    '20487': ['EN.SODE', 'EN.SODE.P4'],
    '20488': ['EN.SODE', 'EN.SODE.P2'],
    '20489': ['CS.RSTS', 'CS.RSTS.M5'],
    '20490': ['EN.SODE', 'EN.SODE.P3'],
    '20491': ['CS.RSTS', 'CS.RSTS.P3'],
    '20492': ['CS.RSTS', 'CS.RSTS.P3'],
    '20493': ['SA.CRCS', 'SA.CRCS.P5'],
    '20494': ['EN.SODE', 'EN.SODE.P4'],
    '20495': ['EN.SODE', 'EN.SODE.P3'],
    '20496': ['EN.SODE', 'EN.SODE.P3'],
    '20497': ['CS.RSTS', 'CS.RSTS.P4'],
    '20498': ['EN.SODE', 'EN.SODE.P3'],
    '20499': ['CS.RSTS', 'CS.RSTS.P3'],
    '20500': ['CS.RSTS', 'CS.RSTS.M5'],
    '20501': ['SA.CRCS', 'SA.CRCS.P5'],
    '20502': ['EN.SODE', 'EN.SODE.P3'],
    '20503': ['SA.CRCS', 'SA.CRCS.P4'],
    '20504': ['EN.SODE', 'EN.SODE.P3'],
    '20505': ['EN.SODE', 'EN.SODE.P3'],
    '20506': ['CS.RSTS', 'CS.RSTS.P2'],
    '20507': ['EN.AIML', 'EN.AIML.M5'],
    '20508': ['SA.CRCS', 'SA.CRCS.P5'],
    '20509': ['EN.SODE', 'EN.SODE.P2'],
    '20510': ['TE.DADA', 'TE.DADA.P4'],
    '20511': ['EN.SODE', 'EN.SODE.P4'],
    '20512': ['EN.SODE', 'EN.SODE.P4'],
    '20513': ['EN.0000', 'EN.0000.E3'],
    '20514': ['CS.RSTS', 'CS.RSTS.P2'],
    '20515': ['SA.CRCS', 'SA.CRCS.P5'],
    '20517': ['EN.UUUD', 'EN.UUUD.P4'],
    '20518': ['EN.SODE', 'EN.SODE.P3'],
    '20519': ['EN.SODE', 'EN.SODE.P3'],
    '20520': ['EN.SODE', 'EN.SODE.P3'],
    '20521': ['EN.SODE', 'EN.SODE.P4'],
    '20522': ['EN.SODE', 'EN.SODE.P3'],
    '20523': ['EN.AIML', 'EN.AIML.P4'],
    '20524': ['EN.SODE', 'EN.SODE.P4'],
    '20525': ['EN.SODE', 'EN.SODE.P3'],
    '20526': ['EN.SODE', 'EN.SODE.P3'],
    '20527': ['EN.SODE', 'EN.SODE.P3'],
    '20528': ['EN.PGPG', 'EN.PGPG.P5'],
    '20529': ['SA.CRCS', 'SA.CRCS.P4'],
    '20530': ['EN.PMPD', 'EN.PMPD.P5'],
    '20531': ['SA.OPSR', 'SA.OPSR.P5'],
    '20532': ['EN.SODE', 'EN.SODE.P3'],
    '20533': ['SA.CRCS', 'SA.CRCS.P4'],
    '20534': ['EN.DODO', 'EN.DODO.P4'],
    '20535': ['EN.SODE', 'EN.SODE.P3'],
    '20536': ['EN.SODE', 'EN.SODE.P4'],
    '20537': ['SA.CRCS', 'SA.CRCS.P4'],
    '20538': ['CS.RSTS', 'CS.RSTS.P4'],
    '20539': ['EN.SODE', 'EN.SODE.P5'],
    '20540': ['TE.DADS', 'TE.DADS.P3'],
    '20541': ['EN.AIML', 'EN.AIML.P3'],
    '20542': ['TE.INMF', 'TE.INMF.E1'],
    '20543': ['EN.SODE', 'EN.SODE.P4'],
    '20544': ['EN.SODE', 'EN.SODE.P3'],
    '20545': ['EN.AIML', 'EN.AIML.P4'],
    '20546': ['EN.SODE', 'EN.SODE.P5'],
    '20547': ['EN.SODE', 'EN.SODE.P4'],
    '20548': ['EN.SODE', 'EN.SODE.P4'],
    '20549': ['HR.TMTA', 'HR.TMTA.P5'],
    '20550': ['EN.AIML', 'EN.AIML.P4'],
    '20551': ['EN.SODE', 'EN.SODE.M6'],
    '20552': ['EN.SODE', 'EN.SODE.P4'],
    '20553': ['TE.DADS', 'TE.DADS.P3'],
    '20554': ['EN.SODE', 'EN.SODE.P4'],
    '20555': ['CS.RSTS', 'CS.RSTS.P4'],
    '20556': ['EN.SODE', 'EN.SODE.P3'],
    '20557': ['CS.RSTS', 'CS.RSTS.P4'],
    '20558': ['EN.PGPG', 'EN.PGPG.P6'],
    '20559': ['EN.SODE', 'EN.SODE.P3'],
    '20560': ['EN.UUUD', 'EN.UUUD.P4'],
    '20561': ['EN.SODE', 'EN.SODE.P5'],
    '20562': ['CS.RSTS', 'CS.RSTS.P4'],
    '20563': ['CS.RSTS', 'CS.RSTS.P4'],
    '20564': ['EN.SODE', 'EN.SODE.P3'],
    '20565': ['EN.SODE', 'EN.SODE.P3'],
    '20566': ['TE.DADA', 'TE.DADA.P3'],
    '20567': ['EN.SODE', 'EN.SODE.P3'],
    '20568': ['SA.ASSN', 'SA.ASSN.P6'],
    '20569': ['CS.RSTS', 'CS.RSTS.P2'],
    '20570': ['TE.DABD', 'TE.DABD.P4'],
    '20571': ['TE.DADA', 'TE.DADA.P3'],
    '20572': ['EN.DODO', 'EN.DODO.P3'],
    '20573': ['TE.DADS', 'TE.DADS.P3'],
    '20574': ['EN.SODE', 'EN.SODE.P4'],
    '20575': ['EN.SODE', 'EN.SODE.P3'],
    '20576': ['EN.SODE', 'EN.SODE.P3'],
    '20577': ['EN.SODE', 'EN.SODE.P5'],
    '20578': ['TE.DABD', 'TE.DABD.P4'],
    '20579': ['EN.SODE', 'EN.SODE.P4'],
    '20580': ['EN.SODE', 'EN.SODE.P3'],
    '20581': ['EN.SODE', 'EN.SODE.P3'],
    '20582': ['CS.RSTS', 'CS.RSTS.P2'],
    '20583': ['CS.RSTS', 'CS.RSTS.P2'],
    '20584': ['EN.SODE', 'EN.SODE.P5'],
    '20585': ['TE.DADS', 'TE.DADS.M5'],
    '20586': ['EN.SODE', 'EN.SODE.P6'],
    '20587': ['EN.DODO', 'EN.DODO.P3'],
    '20588': ['EN.DODO', 'EN.DODO.P5'],
    '20589': ['EN.SODE', 'EN.SODE.P5'],
    '20590': ['HR.GL00', 'HR.GL00.P6'],
    '20591': ['TE.DADS', 'TE.DADS.P4'],
    '20592': ['EN.PGPG', 'EN.PGPG.P4'],
    '20593': ['EN.SODE', 'EN.SODE.P4'],
    '20594': ['TE.DADS', 'TE.DADS.P3'],
    '20595': ['EN.SODE', 'EN.SODE.P3'],
    '20596': ['EN.SODE', 'EN.SODE.P3'],
    '20597': ['EN.SODE', 'EN.SODE.P3'],
    '20598': ['EN.SODE', 'EN.SODE.M4'],
    '20599': ['HR.ERER', 'HR.ERER.P3'],
    '20600': ['SP.BDBD', 'SP.BDBD.P4'],
    '20601': ['CS.RSTS', 'CS.RSTS.P3'],
    '20602': ['CS.RSTS', 'CS.RSTS.P2'],
    '20603': ['SA.CRCS', 'SA.CRCS.P5'],
    '20604': ['EN.SODE', 'EN.SODE.P3'],
    '20605': ['EN.SODE', 'EN.SODE.M6'],
    '20606': ['EN.SODE', 'EN.SODE.P4'],
    '20607': ['SA.CRCS', 'SA.CRCS.P5'],
    '20608': ['CS.RSTS', 'CS.RSTS.P2'],
    '20609': ['', ''],
    '20610': ['EN.SODE', 'EN.SODE.P3'],
    '20611': ['EN.UUUD', 'EN.UUUD.P4'],
    '20612': ['EN.SODE', 'EN.SODE.P3'],
    '20613': ['SA.CRCS', 'SA.CRCS.P4'],
    '20614': ['', ''],
    '100597': ['SA.FAF1', 'SA.FAF1.P6'],
    '102529': ['SA.CRCE', 'SA.CRCE.E3'],
    '102535': ['LE.GLEC', 'LE.GLEC.E6'],
    '102539': ['EN.PGHC', 'EN.PGHC.E3'],
    '104133': ['SA.ASRS', 'SA.ASRS.P6'],
    '105779': ['FI.GLFI', 'FI.GLFI.E5'],
    '110002': ['EN.SODE', 'EN.SODE.M5'],
    '110004': ['SA.CRCS', 'SA.CRCS.M5'],
    '110005': ['CS.RSTS', 'CS.RSTS.P3'],
    '110006': ['EN.SODE', 'EN.SODE.P5'],
    '110008': ['SA.0000', 'SA.0000.E1'],
    '110009': ['CS.RSTS', 'CS.RSTS.M3'],
    '110012': ['EN.SODE', 'EN.SODE.P5'],
    '110023': ['EN.SODE', 'EN.SODE.P4'],
    '110026': ['SA.CRCE', 'SA.CRCE.E1'],
    '110028': ['CS.GLTC', 'CS.GLTC.E3'],
    '110030': ['MK.PIMC', 'MK.PIMC.P5'],
    '110032': ['EN.PGPG', 'EN.PGPG.P5'],
    '110036': ['EN.SODE', 'EN.SODE.P4'],
    '110043': ['EN.SODE', 'EN.SODE.P6'],
    '110046': ['FI.ACCO', 'FI.ACCO.F5'],
    '110047': ['CS.RSTS', 'CS.RSTS.P3'],
    '110048': ['CS.RSTS', 'CS.RSTS.P3'],
    '110050': ['EN.SODE', 'EN.SODE.M5'],
    '110051': ['SA.CRCS', 'SA.CRCS.P5'],
    '110052': ['EN.SODE', 'EN.SODE.P5'],
    '110054': ['SA.CRCS', 'SA.CRCS.M6'],
    '110055': ['MK.PIPM', 'MK.PIPM.P5'],
    '110057': ['TE.DADA', 'TE.DADA.P4'],
    '110059': ['CS.RSTS', 'CS.RSTS.P3'],
    '110063': ['CS.RSTS', 'CS.RSTS.P3'],
    '110064': ['MK.PIPM', 'MK.PIPM.P6'],
    '110065': ['EN.SODE', 'EN.SODE.P3'],
    '110069': ['SA.CRCS', 'SA.CRCS.P5'],
    '110080': ['TE.DADA', 'TE.DADA.M5'],
    '110081': ['TE.DABD', 'TE.DABD.P5'],
    '110083': ['SA.CRCS', 'SA.CRCS.M5'],
    '110084': ['SA.CRCS', 'SA.CRCS.P2'],
    '110085': ['EN.SODE', 'EN.SODE.P5'],
    '110086': ['SA.CRCS', 'SA.CRCS.P5'],
    '110087': ['SA.CRCS', 'SA.CRCS.P5'],
    '110092': ['SA.FAF1', 'SA.FAF1.P6'],
    '110093': ['SA.FAF1', 'SA.FAF1.P6'],
    '110095': ['SA.0000', 'SA.0000.P6'],
    '110096': ['SA.FAF1', 'SA.FAF1.P6'],
    '110097': ['SA.CRCS', 'SA.CRCS.P5'],
    '110098': ['SA.CRCS', 'SA.CRCS.P5'],
    '110100': ['TE.DADA', 'TE.DADA.P5'],
    '110101': ['CS.CSCX', 'CS.CSCX.M6'],
    '110102': ['SA.CRCS', 'SA.CRCS.P5'],
    '110103': ['SA.CRCS', 'SA.CRCS.M5'],
    '110104': ['SA.CRCS', 'SA.CRCS.P5'],
    '110105': ['SA.FAF1', 'SA.FAF1.P6'],
    '110106': ['SA.CRCS', 'SA.CRCS.P5'],
    '110107': ['TE.DADA', 'TE.DADA.P5'],
    '110108': ['SA.CRCS', 'SA.CRCS.P4'],
    '110109': ['SA.CRCS', 'SA.CRCS.M5'],
    '110111': ['TE.DADA', 'TE.DADA.P5'],
    '110112': ['SA.0000', 'SA.0000.E1'],
    '110113': ['SP.BDBD', 'SP.BDBD.P4'],
    '110114': ['SA.ASRS', 'SA.ASRS.P6'],
    '110115': ['SA.FAF1', 'SA.FAF1.P6'],
    '117646': ['SA.FAF1', 'SA.FAF1.P6'],
    '121026': ['EN.PGHC', 'EN.PGHC.E1'],
    '128742': ['SA.FAF1', 'SA.FAF1.P6'],
    '132300': ['EN.PGPG', 'EN.PGPG.P6'],
    '133850': ['EN.UUUD', 'EN.UUUD.M5'],
    '134964': ['MK.PIPM', 'MK.PIPM.M5'],
    '142296': ['SA.CRCS', 'SA.CRCS.P6'],
    '150368': ['SA.FAF1', 'SA.FAF1.P6'],
    '154763': ['SA.0000', 'SA.0000.E1'],
    '155950': ['SA.CRCS', 'SA.CRCS.M5'],
    '156503': ['SA.CRCS', 'SA.CRCS.P5'],
    '159691': ['SA.FAF1', 'SA.FAF1.P6'],
    '160894': ['SA.CRCS', 'SA.CRCS.P5'],
    '160900': ['SA.0000', 'SA.0000.M6'],
    '162147': ['MK.PIPM', 'MK.PIPM.M6'],
    '162639': ['SA.ASRS', 'SA.ASRS.P6'],
    '162822': ['SA.FAF1', 'SA.FAF1.P6'],
    '163290': ['SA.CRCS', 'SA.CRCS.M5'],
    '167571': ['SA.CRCS', 'SA.CRCS.M5'],
    '168279': ['EN.PGPG', 'EN.PGPG.M6'],
    '168894': ['TE.DADA', 'TE.DADA.M6'],
    '179376': ['SA.CRCS', 'SA.CRCS.M5'],
    '180004': ['SA.CRCS', 'SA.CRCS.P5'],
    '180179': ['SA.ASRS', 'SA.ASRS.P6'],
    '181302': ['SA.FAF1', 'SA.FAF1.P6'],
    '181305': ['SA.FAF1', 'SA.FAF1.P6'],
    '182920': ['SA.FAF1', 'SA.FAF1.P6'],
    '186873': ['SA.CRCE', 'SA.CRCE.E1'],
    '187547': ['SA.CRCS', 'SA.CRCS.P5'],
    '187770': ['SA.FAF1', 'SA.FAF1.P6'],
    '187921': ['SA.CRCS', 'SA.CRCS.M5'],
    '188499': ['CB.0000', 'CB.0000.P5'],
    '190331': ['SA.CRCS', 'SA.CRCS.P5'],
    '190698': ['SA.CRCS', 'SA.CRCS.P5'],
    '191024': ['SA.CRCS', 'SA.CRCS.P5'],
    '191961': ['SA.CRCS', 'SA.CRCS.P5'],
    '192526': ['SA.CRCS', 'SA.CRCS.M5'],
    '192990': ['SA.CRCS', 'SA.CRCS.M5'],
    '193050': ['EN.PGPG', 'EN.PGPG.P5'],
    '193191': ['SA.FAF1', 'SA.FAF1.P6'],
    '193480': ['SA.CRCS', 'SA.CRCS.M6'],
    '193836': ['SA.ASRS', 'SA.ASRS.P6'],
    '194529': ['EN.UUUD', 'EN.UUUD.P5'],
    '194948': ['SA.ASRS', 'SA.ASRS.P6'],
    '195304': ['CS.RSTS', 'CS.RSTS.P4'],
    '195462': ['SA.CRCS', 'SA.CRCS.P6'],
    '196193': ['SA.CRCS', 'SA.CRCS.M5'],
    '196295': ['SP.SPMF', 'SP.SPMF.E1'],
    '196621': ['HR.TMTA', 'HR.TMTA.M5'],
    '196968': ['SA.FAF1', 'SA.FAF1.P5'],
    '197271': ['FI.CNCE', 'FI.CNCE.E1'],
    '197388': ['HR.ARIS', 'HR.ARIS.P5'],
    '197695': ['SA.ASRS', 'SA.ASRS.P6'],
    '197696': ['SA.CRCS', 'SA.CRCS.P5'],
    '197774': ['SA.CRCS', 'SA.CRCS.P5'],
    '198331': ['SA.CRCS', 'SA.CRCS.P5'],
    '198354': ['SA.CRCS', 'SA.CRCS.P5'],
    '198475': ['MK.GLHD', 'MK.GLHD.E1'],
    '198674': ['SA.CRCS', 'SA.CRCS.M5'],
    '199026': ['SA.ASRS', 'SA.ASRS.M6'],
    '199347': ['TE.DADA', 'TE.DADA.P5'],
    '199351': ['SA.CRCS', 'SA.CRCS.P5'],
    '199352': ['SA.CRCS', 'SA.CRCS.P5'],
    '199353': ['SA.CRCS', 'SA.CRCS.P5'],
    '199354': ['SA.CRCS', 'SA.CRCS.P6'],
    '199357': ['SA.CRCS', 'SA.CRCS.P6'],
    '199358': ['SA.CRCS', 'SA.CRCS.P5'],
    '199360': ['SA.CRCS', 'SA.CRCS.P5'],
    '199364': ['HR.TMTA', 'HR.TMTA.P4'],
    '199369': ['SA.CRCS', 'SA.CRCS.P5'],
    '199370': ['SA.ASRS', 'SA.ASRS.M6'],
    '199376': ['SA.CRCS', 'SA.CRCS.M6'],
    '199380': ['FI.ACCO', 'FI.ACCO.M5'],
    '199383': ['SA.OPSO', 'SA.OPSO.P6'],
    '199384': ['HR.TMTA', 'HR.TMTA.P5'],
    '199386': ['SA.FAF1', 'SA.FAF1.P6'],
    '199387': ['MK.PIPM', 'MK.PIPM.P6'],
    '199389': ['SA.CRCS', 'SA.CRCS.P5'],
    '199390': ['SA.CRCS', 'SA.CRCS.P5'],
    '199391': ['HR.ARIS', 'HR.ARIS.M5'],
    '199392': ['SA.CRCS', 'SA.CRCS.P5'],
    '199393': ['SA.CRCE', 'SA.CRCE.E1'],
    '199394': ['FI.ACFP', 'FI.ACFP.P5'],
    '199395': ['SA.CRCS', 'SA.CRCS.P5'],
    '199396': ['FI.ACGA', 'FI.ACGA.M4'],
    '199397': ['SA.FAF1', 'SA.FAF1.P6'],
    '199398': ['SA.CRCS', 'SA.CRCS.P5'],
    '199399': ['HR.GLBP', 'HR.GLBP.P6'],
    '199400': ['SA.CRCS', 'SA.CRCS.P5'],
    '199403': ['EN.PGPG', 'EN.PGPG.M6'],
    '199404': ['SA.CRCS', 'SA.CRCS.P5'],
    '199405': ['SA.CRCS', 'SA.CRCS.P5'],
    '199406': ['SA.CRCS', 'SA.CRCS.M5'],
    '199407': ['SA.CRCS', 'SA.CRCS.P5'],
    '199408': ['SA.CRCS', 'SA.CRCS.M5'],
    '199409': ['SA.OPDD', 'SA.OPDD.P6'],
    '199410': ['SA.0000', 'SA.0000.E1'],
    '199411': ['SA.CRCS', 'SA.CRCS.M6'],
    '199412': ['FI.GLFE', 'FI.GLFE.E1'],
    '199413': ['SA.GL00', 'SA.GL00.EA'],
    '199414': ['HR.GLMF', 'HR.GLMF.E3'],
    '199415': ['FI.ACRR', 'FI.ACRR.M5'],
    '199416': ['SA.CRCS', 'SA.CRCS.P5'],
    '199417': ['SA.OPSO', 'SA.OPSO.M6'],
    '199419': ['SA.CRCS', 'SA.CRCS.P5'],
    '199420': ['SA.APMF', 'SA.APMF.E1'],
    '199421': ['MK.PMME', 'MK.PMME.E1'],
    '199422': ['FI.ACGA', 'FI.ACGA.P5'],
    '199424': ['SA.CRCS', 'SA.CRCS.P5'],
    '199425': ['SA.CRCS', 'SA.CRCS.P5'],
    '199426': ['LG.GLMF', 'LG.GLMF.E1'],
    '199427': ['TE.DADA', 'TE.DADA.P4'],
    '199428': ['SA.CRCS', 'SA.CRCS.P5'],
    '199429': ['SA.CRCS', 'SA.CRCS.P5'],
    '199430': ['TE.DADA', 'TE.DADA.P4'],
    '199431': ['SP.BDBD', 'SP.BDBD.P4'],
    '199432': ['SA.FAF1', 'SA.FAF1.P6'],
    '199433': ['SP.BDBD', 'SP.BDBD.P4'],
    '199434': ['SA.OPSO', 'SA.OPSO.M6'],
    '199435': ['FI.ACFP', 'FI.ACFP.P6'],
    '199436': ['SA.ASME', 'SA.ASME.E1'],
    '199437': ['SA.FAF1', 'SA.FAF1.P6'],
    '199438': ['MK.PIDG', 'MK.PIDG.M6'],
    '199439': ['SA.ASRS', 'SA.ASRS.P6'],
    '199440': ['MK.PIPM', 'MK.PIPM.P6'],
    '199441': ['SP.BDBD', 'SP.BDBD.M5'],
    '199442': ['FI.ACRR', 'FI.ACRR.M5'],
    '199443': ['MK.APES', 'MK.APES.P5'],
    '199444': ['HR.SSHR', 'HR.SSHR.P4'],
    '199446': ['EN.GLCC', 'EN.GLCC.E5'],
    '199447': ['EN.PGPG', 'EN.PGPG.M6'],
    '199449': ['MK.PIMC', 'MK.PIMC.P6'],
    '199450': ['SA.OPSO', 'SA.OPSO.P5'],
    '199451': ['SA.CRCS', 'SA.CRCS.P5'],
    '199453': ['SA.CRCS', 'SA.CRCS.M5'],
    '199454': ['SA.CRCS', 'SA.CRCS.P5'],
    '199455': ['SA.CRCS', 'SA.CRCS.P5'],
    '199456': ['HR.TMTA', 'HR.TMTA.P4'],
    '199457': ['SA.CRCS', 'SA.CRCS.M5'],
    '199458': ['SA.CRCS', 'SA.CRCS.P5'],
    '199459': ['HR.TMTA', 'HR.TMTA.M5'],
    '199460': ['FI.GLFI', 'FI.GLFI.E5'],
    '199461': ['HR.TMTA', 'HR.TMTA.P6'],
    '199462': ['SA.FAF1', 'SA.FAF1.P6'],
    '199463': ['SP.BOBI', 'SP.BOBI.P5'],
    '199464': ['FI.ACGA', 'FI.ACGA.P5'],
    '199465': ['FI.ACGA', 'FI.ACGA.P4'],
    '199466': ['SA.CRCS', 'SA.CRCS.P6'],
    '199467': ['HR.TMTA', 'HR.TMTA.P5'],
    '199468': ['MK.GLHD', 'MK.GLHD.E5'],
    '199469': ['SA.FAF1', 'SA.FAF1.P6'],
    '199470': ['LG.GLMF', 'LG.GLMF.E1'],
    '199471': ['SA.CRCS', 'SA.CRCS.P6'],
    '199472': ['MK.PIPM', 'MK.PIPM.P6'],
    '199473': ['SA.CRCS', 'SA.CRCS.P5'],
    '199474': ['SA.FAF1', 'SA.FAF1.P6'],
    '199475': ['SA.OPSV', 'SA.OPSV.E1'],
    '199476': ['TE.DADA', 'TE.DADA.P6'],
    '199477': ['SA.FAF1', 'SA.FAF1.P6'],
    '199478': ['EN.PMPD', 'EN.PMPD.M6'],
    '199479': ['HR.TMTA', 'HR.TMTA.M6'],
    '199480': ['SA.CRCS', 'SA.CRCS.P6'],
    '199481': ['SA.CRCE', 'SA.CRCE.E1'],
    '199482': ['SA.CRCS', 'SA.CRCS.M5'],
    '199483': ['SA.FAF1', 'SA.FAF1.P6'],
    '199484': ['CS.GLTC', 'CS.GLTC.E5'],
    '199485': ['SA.OPDD', 'SA.OPDD.P4'],
    '199486': ['MK.PIPM', 'MK.PIPM.P5'],
    '199487': ['EN.PMPD', 'EN.PMPD.M6'],
    '199488': ['SA.CRCS', 'SA.CRCS.M5'],
    '199489': ['EN.PGHC', 'EN.PGHC.E1'],
    '199490': ['SA.CRCS', 'SA.CRCS.P5'],
    '199491': ['EN.SODE', 'EN.SODE.P5'],
    '199492': ['', ''],
    '199493': ['SA.CRCS', 'SA.CRCS.P5'],
    '199494': ['', '']
  };
  
  // Convert to array format
  const rows = [];
  for (const empID in mappings) {
    const [jobFamily, fullMapping] = mappings[empID];
    rows.push([empID, jobFamily, fullMapping]);
  }
  
  return rows;
}

/**
 * Loads all legacy mappings at once (more efficient than per-employee lookup)
 * Returns Map: empID → {aonCode, ciqLevel}
 * @returns {Map<string, {aonCode: string, ciqLevel: string}>}
 */
function _loadAllLegacyMappings_() {
  const legacyMap = new Map();
  
  // Try Script Properties first (persistent storage)
  const storedData = _loadLegacyMappingsFromStorage_();
  if (storedData && storedData.length > 0) {
    storedData.forEach(row => {
      const empID = String(row[0] || '').trim();
      const fullMapping = String(row[2] || '').trim();
      if (!empID || !fullMapping) return;
      
      // Parse full mapping (e.g., "EN.SODE.P5")
      const parts = fullMapping.split('.');
      if (parts.length < 3) return;
      
      const aonCode = `${parts[0]}.${parts[1]}`;
      const levelToken = parts[2];
      const ciqLevel = _parseLevelToken_(levelToken);
      
      if (aonCode && ciqLevel) {
        legacyMap.set(empID, {aonCode, ciqLevel, source: 'Legacy'});
      }
    });
    return legacyMap;
  }
  
  // Fallback to sheet
  const ss = SpreadsheetApp.getActive();
  const legacySh = ss.getSheetByName(SHEET_NAMES.LEGACY_MAPPINGS);
  if (legacySh && legacySh.getLastRow() > 1) {
    const legacyVals = legacySh.getRange(2,1,legacySh.getLastRow()-1,3).getValues();
    legacyVals.forEach(row => {
      const empID = String(row[0] || '').trim();
      const fullMapping = String(row[2] || '').trim();
      if (!empID || !fullMapping) return;
      
      const parts = fullMapping.split('.');
      if (parts.length < 3) return;
      
      const aonCode = `${parts[0]}.${parts[1]}`;
      const levelToken = parts[2];
      const ciqLevel = _parseLevelToken_(levelToken);
      
      if (aonCode && ciqLevel) {
        legacyMap.set(empID, {aonCode, ciqLevel, source: 'Legacy'});
      }
    });
  }
  
  // Fallback to embedded data if both storage and sheet are empty
  if (legacyMap.size === 0) {
    const embeddedData = _getLegacyMappingData_();
    embeddedData.forEach(row => {
      const empID = String(row[0] || '').trim();
      const fullMapping = String(row[2] || '').trim();
      if (!empID || !fullMapping) return;
      
      const parts = fullMapping.split('.');
      if (parts.length < 3) return;
      
      const aonCode = `${parts[0]}.${parts[1]}`;
      const levelToken = parts[2];
      const ciqLevel = _parseLevelToken_(levelToken);
      
      if (aonCode && ciqLevel) {
        legacyMap.set(empID, {aonCode, ciqLevel, source: 'Legacy'});
      }
    });
  }
  
  return legacyMap;
}

/**
 * Gets legacy mapping for an employee
 * Reads from Script Properties (persistent storage) first, then falls back to sheet
 */
function _getLegacyMapping_(empID) {
  // Try to load all legacy data (cached for performance)
  let legacyData = _loadLegacyMappingsFromStorage_();
  
  // Fallback to sheet if storage is empty
  if (!legacyData || legacyData.length === 0) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAMES.LEGACY_MAPPINGS);
    if (!sh || sh.getLastRow() <= 1) return null;
    legacyData = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
  }
  
  // Find employee in legacy data
  for (let r=0; r<legacyData.length; r++) {
    if (String(legacyData[r][0]).trim() === String(empID).trim()) {
      const fullMapping = String(legacyData[r][2] || '').trim();
      if (!fullMapping) return null;
      
      // Parse full mapping (e.g., "EN.SODE.P5" → aonCode="EN.SODE", level="L5 IC")
      const parts = fullMapping.split('.');
      if (parts.length < 3) return null;
      
      const aonCode = `${parts[0]}.${parts[1]}`;
      const levelToken = parts[2]; // e.g., "P5", "M4", "E3"
      const ciqLevel = _parseLevelToken_(levelToken);
      
      return { aonCode, ciqLevel, source: 'Legacy' };
    }
  }
  return null;
}

/**
 * Updates Legacy Mappings sheet from approved Employees Mapped entries
 * This creates a feedback loop: approved mappings become the new legacy data
 */
function updateLegacyMappingsFromApproved_() {
  const ss = SpreadsheetApp.getActive();
  const empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
  const legacySh = ss.getSheetByName(SHEET_NAMES.LEGACY_MAPPINGS);
  
  if (!empSh || empSh.getLastRow() <= 1) {
    SpreadsheetApp.getActive().toast('Employees Mapped sheet not found', 'Skipped', 3);
    return;
  }
  
  if (!legacySh) {
    SpreadsheetApp.getActive().toast('Legacy Mappings sheet not found', 'Skipped', 3);
    return;
  }
  
  // Get all approved mappings from Employees Mapped
  const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,15).getValues();
  const approvedMappings = new Map(); // empID → {jobFamily, fullMapping}
  
  let approvedCount = 0;
  let skippedCount = 0;
  
  empVals.forEach(row => {
    const empID = String(row[0] || '').trim();
    const aonCode = String(row[5] || '').trim(); // Column F (index 5)
    const ciqLevel = String(row[7] || '').trim(); // Column H (index 7)
    const status = String(row[10] || '').trim(); // Column K (index 10)
    
    // Debug logging for first few rows
    if (approvedCount + skippedCount < 3) {
      Logger.log(`Row ${approvedCount + skippedCount + 1}: EmpID=${empID}, Status="${status}", AonCode=${aonCode}, Level=${ciqLevel}`);
    }
    
    // Only sync approved mappings
    if (status === 'Approved' && empID && aonCode && ciqLevel) {
      const jobFamily = aonCode; // e.g., "EN.SODE"
      const levelToken = _ciqLevelToToken_(ciqLevel); // e.g., "L5 IC" → "P5"
      const fullMapping = levelToken ? `${aonCode}.${levelToken}` : '';
      
      if (fullMapping) {
        approvedMappings.set(empID, {jobFamily, fullMapping});
        approvedCount++;
      } else {
        Logger.log(`⚠️ Could not convert level "${ciqLevel}" to token for ${empID}`);
        skippedCount++;
      }
    } else if (empID) {
      skippedCount++;
    }
  });
  
  Logger.log(`Found ${approvedCount} approved mappings, skipped ${skippedCount} rows`);
  
  if (approvedMappings.size === 0) {
    const msg = `No approved mappings found.\n\n` +
      `📋 Checked ${empVals.length} employees\n` +
      `✓ To approve: Set Status = "Approved" in column K\n` +
      `✓ Ensure Aon Code (F) and Level (H) are filled`;
    SpreadsheetApp.getActive().toast(msg, 'Legacy Mappings', 8);
    return;
  }
  
  // Get existing legacy data
  const existingMap = new Map(); // empID → row index
  if (legacySh.getLastRow() > 1) {
    const legacyVals = legacySh.getRange(2,1,legacySh.getLastRow()-1,3).getValues();
    legacyVals.forEach((row, idx) => {
      const empID = String(row[0] || '').trim();
      if (empID) {
        existingMap.set(empID, idx + 2); // +2 for header and 0-index
      }
    });
  }
  
  // Prepare update/insert rows
  const updates = []; // [rowNum, [empID, jobFamily, fullMapping]]
  const inserts = []; // [empID, jobFamily, fullMapping]
  
  approvedMappings.forEach((mapping, empID) => {
    if (existingMap.has(empID)) {
      // Update existing row
      const rowNum = existingMap.get(empID);
      updates.push([rowNum, [empID, mapping.jobFamily, mapping.fullMapping]]);
    } else {
      // Insert new row
      inserts.push([empID, mapping.jobFamily, mapping.fullMapping]);
    }
  });
  
  // Apply updates
  updates.forEach(([rowNum, data]) => {
    legacySh.getRange(rowNum, 1, 1, 3).setValues([data]);
  });
  
  // Apply inserts
  if (inserts.length > 0) {
    legacySh.getRange(legacySh.getLastRow() + 1, 1, inserts.length, 3).setValues(inserts);
  }
  
  // Save all mappings to Script Properties (persistent storage)
  const allLegacyData = legacySh.getRange(2, 1, legacySh.getLastRow() - 1, 3).getValues();
  _saveLegacyMappingsToStorage_(allLegacyData);
  
  const msg = `✅ Updated Legacy Mappings:\n\n` +
    `📝 ${updates.length} updated\n` +
    `➕ ${inserts.length} new\n` +
    `💾 ${allLegacyData.length} total in storage\n\n` +
    `Saved to persistent storage ✓`;
  SpreadsheetApp.getActive().toast(msg, 'Legacy Mappings Synced', 8);
  
  Logger.log(`Successfully synced ${approvedMappings.size} approved mappings to persistent storage`);
}

/**
 * Converts CIQ level to Aon level token
 * E.g., "L5 IC" → "P5", "L4 Mgr" → "M4", "L7 Mgr" → "E3"
 */
function _ciqLevelToToken_(ciqLevel) {
  const s = String(ciqLevel || '').trim();
  const match = s.match(/^L([\d.]+)\s+(IC|Mgr)$/i);
  if (!match) return '';
  
  const levelNum = parseFloat(match[1]);
  const role = match[2].toLowerCase();
  
  if (role === 'ic') {
    // IC levels
    if (levelNum <= 6.5) {
      return `P${Math.floor(levelNum)}`;
    } else {
      return 'E1'; // L7 IC = E1
    }
  } else {
    // Manager levels
    if (levelNum <= 6.5) {
      return `M${Math.floor(levelNum)}`;
    } else if (levelNum === 7) {
      return 'E1';
    } else if (levelNum === 8) {
      return 'E3';
    } else if (levelNum === 9) {
      return 'E5';
    } else if (levelNum === 10) {
      return 'E6';
    }
  }
  
  return '';
}

/**
 * Parses Aon level token (P5, M4, E3) to CIQ level (L5 IC, L4 Mgr, L3 IC)
 */
function _parseLevelToken_(token) {
  if (!token) return '';
  const match = token.match(/^([PME])(\d+)$/);
  if (!match) return '';
  
  const letter = match[1];
  const num = parseInt(match[2]);
  
  if (letter === 'P') return `L${num} IC`;
  if (letter === 'M') return `L${num} Mgr`;
  if (letter === 'E') {
    // Executive mapping (per user specification)
    // E1 = VP = L7 Mgr
    // E3 = SVP = L8 Mgr
    // E5 = C-Suite = L9 Mgr
    // E6 = CEO = L10+ Mgr
    if (num === 1) return 'L7 Mgr';
    if (num === 3) return 'L8 Mgr';
    if (num === 5) return 'L9 Mgr';
    if (num === 6) return 'L10+ Mgr';
  }
  return '';
}

/**
 * Creates comprehensive Lookup sheet with all mappings
 * Single source of truth for: Level mapping, Category assignment, FX rates
 */
function createLookupSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Lookup');
  if (!sh) {
    sh = ss.insertSheet('Lookup');
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  
  // Clear existing content
  sh.clearContents();
  
  // === SECTION 1: CIQ LEVEL → AON LEVEL MAPPING ===
  let currentRow = 1;
  sh.getRange(currentRow, 1, 1, 2).setValues([['CIQ Level', 'Aon Level']]);
  sh.getRange(currentRow, 1, 1, 2).setFontWeight('bold').setBackground('#4A148C').setFontColor('#FFFFFF');
  currentRow++;
  
  const levelData = [
    ['L2 IC', 'P2'],
    ['L3 IC', 'P3'],
    ['L4 IC', 'P4'],
    ['L5 IC', 'P5'],
    ['L5.5 IC', 'Avg of P5 and P6'],
    ['L6 IC', 'P6'],
    ['L6.5 IC', 'Avg of P6 and E1'],
    ['L7 IC', 'E1'],
    ['L4 Mgr', 'M3'],
    ['L5 Mgr', 'M4'],
    ['L5.5 Mgr', 'Avg of M4 and M5'],
    ['L6 Mgr', 'M5'],
    ['L6.5 Mgr', 'M6'],
    ['L7 Mgr', 'E1'],
    ['L8 Mgr', 'E3'],
    ['L9 Mgr', 'E5'],
    ['L10 Mgr', 'E6']
  ];
  sh.getRange(currentRow, 1, levelData.length, 2).setValues(levelData);
  currentRow += levelData.length + 2;
  
  // === SECTION 2: REGION/SITE → FX MAPPING ===
  sh.getRange(currentRow, 1, 1, 3).setValues([['Region', 'Site', 'FX Rate']]);
  sh.getRange(currentRow, 1, 1, 3).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFFFFF');
  currentRow++;
  
  const fxData = [
    ['India', 'India', 0.0125],
    ['USA', 'US', 1],
    ['UK', 'UK', 1.37]
  ];
  sh.getRange(currentRow, 1, fxData.length, 3).setValues(fxData);
  currentRow += fxData.length + 2;
  
  // === SECTION 3: AON CODE → JOB FAMILY + CATEGORY MAPPING ===
  sh.getRange(currentRow, 1, 1, 3).setValues([['Aon Code', 'Job Family (Exec Description)', 'Category']]);
  sh.getRange(currentRow, 1, 1, 3).setFontWeight('bold').setBackground('#2E7D32').setFontColor('#FFFFFF');
  currentRow++;
  
  const categoryData = [
    // X0 CATEGORIES (Engineering & Product)
    ['EN.AIML', 'Engineering - AI/ML', 'X0'],
    ['EN.PGPG', 'Engineering - Product Management/ TPM', 'X0'],
    ['EN.SODE', 'Engineering - Software Development', 'X0'],
    ['EN.UUUD', 'Engineering - Product Design', 'X0'],
    ['EN.0000', 'Engineering - CTO', 'X0'],
    ['EN.GLCC', 'Engineering - CTO', 'X0'],
    ['EN.PGHC', 'Engineering - CPO (Product Leadership)', 'X0'],
    ['EN.SDCD', 'Engineering - System Design & Cloud Architecture', 'X0'],
    ['TE.DADS', 'Data - Data Science', 'X0'],
    ['TE.DABD', 'Data - Big Data Engineering', 'X0'],
    ['EN.DVEX', 'Engineering - Architect / Distinguished Engineer', 'X0'],
    ['EN.DVDE', 'Engineering - Architect', 'X0'],
    
    // Y1 CATEGORIES (Everyone Else)
    ['LE.GLEC', 'CEO', 'Y1'],
    ['CB.0000', 'Corporate - Executive Assistant', 'Y1'],
    ['CB.ADEA', 'Corporate - Executive Assistant', 'Y1'],
    ['CB.ADCE', 'Leadership - Executive Assistant', 'Y1'],
    ['SP.SPMF', 'Corporate - Strategic Planning (Sr. Leadership)', 'Y1'],
    ['SP.BOBI', 'Corporate : Business Intelligence', 'Y1'],
    ['CS.CSAS', 'Customer Support - Account Services', 'Y1'],
    ['CS.GLTC', 'Customer Support - CCO', 'Y1'],
    ['CS.RSTS', 'Customer Support - Tech Support', 'Y1'],
    ['CS.CSCX', 'Customer Support - Tech Support (Leadership)', 'Y1'],
    ['TE.DADA', 'Data - Analysis & Insights', 'Y1'],
    ['EN.DODO', 'Engineering - DevOps', 'Y1'],
    ['TE.INMF', 'Engineering - DevOps & Infrastructure (Leadership)', 'Y1'],
    ['EN.PMPD', 'Engineering - Agile/Project Management', 'Y1'],
    ['FI.ACRR', 'Finance - Accounting - Revenue', 'Y1'],
    ['FI.GLFI', 'Finance - CFO', 'Y1'],
    ['FI.ACCO', 'Finance - Controller', 'Y1'],
    ['FI.CNCE', 'Finance - Controller (Leadership)', 'Y1'],
    ['FI.ACFP', 'Finance - FP&A', 'Y1'],
    ['FI.ACGA', 'Finance - General Accounting', 'Y1'],
    ['FI.OPMF', 'Finance - Multi Focus (Leadership)', 'Y1'],
    ['FI.GLFE', 'Finance - Multi Focus (Senior Leadership)', 'Y1'],
    ['HR.GLBP', 'HR - Business Partner', 'Y1'],
    ['HR.GLGL', 'HR - Generalist (BP + TA)', 'Y1'],
    ['HR.ARIS', 'HR - HRIS Administrator/People Operations', 'Y1'],
    ['HR.GLMF', 'HR - Leadership/CHRO', 'Y1'],
    ['HR.SSHR', 'HR - Specialist/Shared Services', 'Y1'],
    ['HR.GL00', 'HR - Strategy (Multi Focus)', 'Y1'],
    ['HR.TMTA', 'HR - Talent Acquisition', 'Y1'],
    ['HR.TATA', 'HR - Talent Acquisition', 'Y1'],
    ['CB.ADAA', 'HR - Workplace Services', 'Y1'],
    ['CB.ASAS', 'HR - Workplace Services', 'Y1'],
    ['LG.GLMF', 'Legal - General Counsel', 'Y1'],
    ['SP.BDBD', 'Marketing - Business Development', 'Y1'],
    ['MK.GLHD', 'Marketing - CMO', 'Y1'],
    ['MK.PIMC', 'Marketing - Communications', 'Y1'],
    ['MK.PIDG', 'Marketing - Demand Generation', 'Y1'],
    ['MK.APES', 'Marketing - Events', 'Y1'],
    ['MK.CIDB', 'Marketing - Graphic/Web Design', 'Y1'],
    ['MK.PIPM', 'Marketing - Product Marketing', 'Y1'],
    ['MK.PMME', 'Marketing - Product Marketing Leadership', 'Y1'],
    ['SA.GL00', 'Sales - C- Level Leadership', 'Y1'],
    ['SA.CRCS', 'Sales - Customer Success', 'Y1'],
    ['SA.CRCE', 'Sales - Customer Success - Sr. Leadership', 'Y1'],
    ['SA.OPDD', 'Sales - Deal Desk', 'Y1'],
    ['SA.FSDS', 'Sales - Direct Sales', 'Y1'],
    ['SA.OPSE', 'Sales - Enablement', 'Y1'],
    ['SA.GLMF', 'Sales - Leadership', 'Y1'],
    ['SA.OPSO', 'Sales - Operations & Enablement', 'Y1'],
    ['SA.OPSV', 'Sales - Operations Leadership', 'Y1'],
    ['SA.APMF', 'Sales - Partnerships (Leadership)', 'Y1'],
    ['SA.GLSX', 'Sales - Regional Leadership', 'Y1'],
    ['SA.OPSR', 'Sales - Salesforce Administrator', 'Y1'],
    ['SA.FAF1', 'Sales - Senior & Strategic Accounts Executives', 'Y1'],
    ['SA.ASSN', 'Sales - Solutions Consulting (Bonus)', 'Y1'],
    ['SA.ASRS', 'Sales - Solutions Consulting (Commissions)', 'Y1'],
    ['SA.ASME', 'Sales - Solutions Consulting (Leadership)', 'Y1'],
    ['SA.0000', 'Sales - Sr. Leadership', 'Y1']
  ];
  sh.getRange(currentRow, 1, categoryData.length, 3).setValues(categoryData);
  
  // Format and freeze
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
  
  SpreadsheetApp.getActive().toast('Lookup sheet created with comprehensive mappings', 'Done', 5);
}

/**
 * Creates Y1 calculator UI (Everyone Else)
 * Range: P10 → P40 → P62.5
 */
function buildCalculatorUIForY1_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(UI_SHEET_NAME_Y1);
  if (!sh) {
    sh = ss.insertSheet(UI_SHEET_NAME_Y1);
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  
  // Get Y1 families only
  const categoryMap = _getCategoryMap_();
  const execMap = _getExecDescMap_();
  const y1Families = [];
  categoryMap.forEach((cat, code) => {
    if (cat === 'Y1') {
      const desc = execMap.get(code);
      if (desc) y1Families.push(desc);
    }
  });
  
  // Job Family dropdown (Y1 families only)
  if (y1Families.length > 0) {
    const uniq = Array.from(new Set(y1Families)).sort();
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniq, true)
      .setAllowInvalid(false)
      .build();
    sh.getRange('B2').setDataValidation(rule);
  }
  
  // Labels
  sh.getRange('A2').setValue('Job Family');
  sh.getRange('A3').setValue('Region');
  sh.getRange('A4').setValue('Currency');

  // Region dropdown
  const regionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['US', 'India', 'UK'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B3').setDataValidation(regionRule);
  if (!sh.getRange('B3').getValue()) sh.getRange('B3').setValue('US');

  // Currency dropdown (Local/USD)
  const currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Local', 'USD'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('B4').setDataValidation(currencyRule);
  if (!sh.getRange('B4').getValue()) sh.getRange('B4').setValue('Local');

  // Header row - Market Range
  sh.getRange('A7').setValue('Level');
  sh.getRange('B7').setValue('Range Start');
  sh.getRange('C7').setValue('Range Mid');
  sh.getRange('D7').setValue('Range End');
  
  // Header row - Internal Range
  sh.getRange('F7').setValue('Min');
  sh.getRange('G7').setValue('Median');
  sh.getRange('H7').setValue('Max');
  sh.getRange('I7').setValue('Emp Count');
  sh.getRange('J7').setValue('Avg CR');
  sh.getRange('K7').setValue('TT CR');
  sh.getRange('L7').setValue('New Hire CR');
  sh.getRange('M7').setValue('BT CR');
  
  // Level list
  const levels = ['L2 IC','L3 IC','L4 IC','L5 IC','L5.5 IC','L6 IC','L6.5 IC','L7 IC','L4 Mgr','L5 Mgr','L5.5 Mgr','L6 Mgr','L6.5 Mgr','L7 Mgr','L8 Mgr','L9 Mgr'];
  sh.getRange(8,1,levels.length,1).setValues(levels.map(s=>[s]));
  
  // Formulas (same as X0)
  const formulasRangeStart = [], formulasRangeMid = [], formulasRangeEnd = [];
  const formulasIntMin = [], formulasIntMed = [], formulasIntMax = [], formulasIntCount = [];
  const formulasAvgCR = [], formulasTTCR = [], formulasNewHireCR = [], formulasBTCR = [];
  
  levels.forEach((level, i) => {
    const aRow = 8 + i;
    
    // Market Range: Currency-aware XLOOKUP (Column N=Range Start, O=Range Mid, P=Range End)
    // FIX: KEY is in Column Y, not U!
    formulasRangeStart.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$N:$N,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$N:$N,""))`]);
    formulasRangeMid.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$O:$O,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$O:$O,""))`]);
    formulasRangeEnd.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$P:$P,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$P:$P,""))`]);
    
    // Internal stats (Column Q=Internal Min, R=Median, S=Max, T=Emp Count)
    // Currency-aware: Switch between Full List (local) and Full List USD
    formulasIntMin.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$Q:$Q,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$Q:$Q,""))`]);
    formulasIntMed.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$R:$R,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$R:$R,""))`]);
    formulasIntMax.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$S:$S,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$S:$S,""))`]);
    formulasIntCount.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$T:$T,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$Y:$Y,'Full List USD'!$T:$T,""))`]);
    
    // CR columns - XLOOKUP from Full List (pre-calculated)
    formulasAvgCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$U:$U,"")`]);
    formulasTTCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$V:$V,"")`]);
    formulasNewHireCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$W:$W,"")`]);
    formulasBTCR.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$Y:$Y,'Full List'!$X:$X,"")`]);
  });
  
  // Set formulas
  sh.getRange(8, 2, levels.length, 1).setFormulas(formulasRangeStart);
  sh.getRange(8, 3, levels.length, 1).setFormulas(formulasRangeMid);
  sh.getRange(8, 4, levels.length, 1).setFormulas(formulasRangeEnd);
  sh.getRange(8, 6, levels.length, 1).setFormulas(formulasIntMin);
  sh.getRange(8, 7, levels.length, 1).setFormulas(formulasIntMed);
  sh.getRange(8, 8, levels.length, 1).setFormulas(formulasIntMax);
  sh.getRange(8, 9, levels.length, 1).setFormulas(formulasIntCount);
  sh.getRange(8,10, levels.length, 1).setFormulas(formulasAvgCR);
  sh.getRange(8,11, levels.length, 1).setFormulas(formulasTTCR);
  sh.getRange(8,12, levels.length, 1).setFormulas(formulasNewHireCR);
  sh.getRange(8,13, levels.length, 1).setFormulas(formulasBTCR);
  
  // Apply currency formatting (includes hide-zero format for Emp Count)
  applyCurrency_();
}

/**
 * Creates Full List placeholder sheets
 */
function createFullListPlaceholders_() {
  const ss = SpreadsheetApp.getActive();
  
  // Full List
  let sh = ss.getSheetByName('Full List');
  if (!sh) {
    sh = ss.insertSheet('Full List');
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,25).setValues([[ 
      'Site', 'Region', 'Aon Code (base)', 'Job Family (Exec)', 'Category', 'CIQ Level',
      'P10', 'P25', 'P40', 'P50', 'P62.5', 'P75', 'P90',
      'Range Start', 'Range Mid', 'Range End',
      'Internal Min', 'Internal Median', 'Internal Max', 'Emp Count',
      'Avg CR', 'TT CR', 'New Hire CR', 'BT CR',
      'Key'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,25).setFontWeight('bold');
    sh.autoResizeColumns(1,25);
  }
  
  // Full List USD
  sh = ss.getSheetByName('Full List USD');
  if (!sh) {
    sh = ss.insertSheet('Full List USD');
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,25).setValues([[ 
      'Site', 'Region', 'Aon Code (base)', 'Job Family (Exec)', 'Category', 'CIQ Level',
      'P10', 'P25', 'P40', 'P50', 'P62.5', 'P75', 'P90',
      'Range Start', 'Range Mid', 'Range End',
      'Internal Min', 'Internal Median', 'Internal Max', 'Emp Count',
      'Avg CR', 'TT CR', 'New Hire CR', 'BT CR',
      'Key'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,25).setFontWeight('bold');
    sh.autoResizeColumns(1,25);
  }
}

/**
 * Syncs Employees Mapped sheet with Base Data using smart mapping logic
 * Priority: Legacy > Existing Approved > Title-Based Suggestion
 */
function syncEmployeesMappedSheet_() {
  const ss = SpreadsheetApp.getActive();
  const baseSh = ss.getSheetByName(SHEET_NAMES.BASE_DATA);
  if (!baseSh || baseSh.getLastRow() <= 1) {
    SpreadsheetApp.getActive().toast('Base Data not found or empty', 'Skipped', 3);
    return;
  }
  
  let empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
  if (!empSh) {
    empSh = ss.insertSheet(SHEET_NAMES.EMPLOYEES_MAPPED);
    empSh.setTabColor('#FF0000');
  }
  
  // Create headers if needed (20 columns total)
  if (empSh.getLastRow() === 0) {
    empSh.getRange(1,1,1,20).setValues([[ 
      'Employee ID', 'Employee Name', 'Job Title', 'Department', 'Site',
      'Aon Code', 'Job Family (Exec Description)', 'Level', 'Full Aon Code', 'Mapping Override',
      'Confidence', 'Source', 'Status', 'Base Salary', 'Start Date',
      'Recent Promotion', 'Level Anomaly', 'Title Anomaly', 'Market Data Missing', 'Active/Inactive'
    ]]);
    empSh.setFrozenRows(1);
    empSh.getRange(1,1,1,20).setFontWeight('bold');
    
    // Highlight editable columns (F: Aon Code, I: Full Aon Code)
    empSh.getRange('F1').setBackground('#FFFFCC');
    empSh.getRange('I1').setBackground('#FFFFCC');
  }
  
  // Get existing mappings (preserve approved ones and user edits to Column I)
  const existing = new Map();
  if (empSh.getLastRow() > 1) {
    const maxCols = Math.min(20, empSh.getLastColumn()); // Handle both 19 and 20 column versions
    const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,maxCols).getValues();
    empVals.forEach(row => {
      if (row[0]) {
        existing.set(String(row[0]).trim(), {
          aonCode: row[5] || '',           // Column F
          jobFamilyDesc: row[6] || '',     // Column G
          level: row[7] || '',             // Column H
          fullAonCode: row[8] || '',       // Column I (preserve user edits)
          confidence: row[10] || '',       // Column K
          source: row[11] || '',           // Column L
          status: row[12] || ''            // Column M
        });
      }
    });
  }
  
  // Get Base Data
  const baseVals = baseSh.getDataRange().getValues();
  if (baseVals.length <= 1) return;
  
  const baseHead = baseVals[0].map(h => String(h||''));
  const iEmpID = baseHead.findIndex(h => /Emp.*ID|Employee.*ID/i.test(h));
  const iName = baseHead.findIndex(h => /(Display.*name|^Name$|Emp.*Name)/i.test(h));
  const iTitle = baseHead.findIndex(h => /Job.*title/i.test(h));
  const iDept = baseHead.findIndex(h => /Department/i.test(h));
  const iSite = baseHead.findIndex(h => /^Site$/i.test(h));
  const iJobLevel = baseHead.findIndex(h => /Job.*level/i.test(h));
  const iSalary = baseHead.findIndex(h => /Base.*salary/i.test(h));
  const iStart = baseHead.findIndex(h => /Start.*date/i.test(h));
  const iActive = baseHead.findIndex(h => /Active.*Inactive|Status/i.test(h));
  const iTerm = baseHead.findIndex(h => /Termination.*date|Term.*date|End.*date|Leave.*date/i.test(h));
  
  // DEBUG: Log column detection
  Logger.log(`\n=== BASE DATA COLUMN DETECTION ===`);
  Logger.log(`Total columns in Base Data: ${baseHead.length}`);
  Logger.log(`Column C (index 2): "${baseHead[2]}"`);
  Logger.log(`Start Date column detected at index: ${iStart} (Column ${String.fromCharCode(65 + iStart)})`);
  if (iStart >= 0) {
    Logger.log(`Start Date header: "${baseHead[iStart]}"`);
  } else {
    Logger.log(`⚠️ Start Date column NOT detected! Headers: ${baseHead.slice(0, 10).join(', ')}`);
  }
  Logger.log(`===\n`);
  
  if (iEmpID < 0) {
    SpreadsheetApp.getActive().toast('Employee ID column not found in Base Data', 'Error', 5);
    return;
  }
  
  // Progress indicator
  SpreadsheetApp.getActive().toast('Loading employee data...', 'Employee Mapping', 3);
  
  // Cutoff date: Jan 1, 2024 for filtering exits
  const exitCutoffDate = new Date('2024-01-01');
  
  // OPTIMIZATION: Build employee ID → title index ONCE (eliminates O(n²) nested loop)
  const empToTitle = new Map(); // empID → title
  for (let i = 1; i < baseVals.length; i++) {
    const empID = String(baseVals[i][iEmpID] || '').trim();
    const title = iTitle >= 0 ? String(baseVals[i][iTitle] || '').trim() : '';
    if (empID && title) {
      empToTitle.set(empID, title);
    }
  }
  
  // OPTIMIZATION: Load ALL legacy mappings ONCE (eliminates 600+ individual lookups)
  SpreadsheetApp.getActive().toast('Loading legacy mappings...', 'Employee Mapping', 3);
  const allLegacyMappings = _loadAllLegacyMappings_();
  
  // OPTIMIZATION: Pre-load Aon data ONCE to check market data availability
  SpreadsheetApp.getActive().toast('Loading market data...', 'Employee Mapping', 3);
  const aonCache = _preloadAonData_();
  
  // Load Comp History for recent promotions (last 90 days)
  SpreadsheetApp.getActive().toast('Checking promotions...', 'Employee Mapping', 3);
  const compHistSh = ss.getSheetByName('Comp History');
  const promotionMap = new Map(); // empID → {date, reason}
  const promotionCutoffDate = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000); // 90 days ago
  
  if (compHistSh && compHistSh.getLastRow() > 1) {
    const compHistVals = compHistSh.getDataRange().getValues();
    const compHistHead = compHistVals[0].map(h => String(h||''));
    const iCompEmpID = compHistHead.findIndex(h => /Emp.*ID|Employee.*ID/i.test(h));
    const iHistReason = compHistHead.findIndex(h => /History.*reason|Reason/i.test(h));
    const iEffDate = compHistHead.findIndex(h => /Effective.*date|Eff.*date/i.test(h));
    
    if (iCompEmpID >= 0 && iHistReason >= 0 && iEffDate >= 0) {
      for (let i = 1; i < compHistVals.length; i++) {
        const row = compHistVals[i];
        const empID = String(row[iCompEmpID] || '').trim();
        const reason = String(row[iHistReason] || '').toLowerCase();
        const effDate = row[iEffDate];
        
        // Check if reason indicates promotion
        if (reason && (reason.includes('promotion') || reason.includes('promoted') || reason.includes('promo'))) {
          const effDateObj = effDate instanceof Date ? effDate : new Date(effDate);
          if (effDateObj && !isNaN(effDateObj.getTime()) && effDateObj >= promotionCutoffDate) {
            // Store most recent promotion for this employee
            if (!promotionMap.has(empID) || effDateObj > promotionMap.get(empID).date) {
              promotionMap.set(empID, {date: effDateObj, reason: row[iHistReason]});
            }
          }
        }
      }
      Logger.log(`Found ${promotionMap.size} employees with promotions in last 90 days`);
    }
  }
  
  // Build title-to-mapping suggestions inline (no separate Title Mapping sheet needed)
  SpreadsheetApp.getActive().toast('Building smart suggestions (1/3)...', 'Employee Mapping', 3);
  const titleToMappings = new Map(); // title → {aonCode|level → count}
  
  // Collect from Legacy Mappings using optimized index lookup
  allLegacyMappings.forEach((mapping, legacyEmpID) => {
    const jobTitle = empToTitle.get(legacyEmpID);
    if (jobTitle) {
      if (!titleToMappings.has(jobTitle)) {
        titleToMappings.set(jobTitle, new Map());
      }
      const key = `${mapping.aonCode}|${mapping.ciqLevel}`;
      const mappingMap = titleToMappings.get(jobTitle);
      mappingMap.set(key, (mappingMap.get(key) || 0) + 1);
    }
  });
  
  // Build most common mapping per title
  const titleMap = new Map();
  titleToMappings.forEach((mappings, title) => {
    let maxCount = 0, bestMapping = null;
    mappings.forEach((count, key) => {
      if (count > maxCount) {
        maxCount = count;
        const [aonCode, level] = key.split('|');
        bestMapping = {aonCode, level, count};
      }
    });
    if (bestMapping) {
      titleMap.set(title, bestMapping);
    }
  });
  
  // Get Job Family descriptions from Lookup
  const execDescMap = _getExecDescMap_();
  
  // Build new rows
  SpreadsheetApp.getActive().toast('Processing employees (2/3)...', 'Employee Mapping', 3);
  const rows = [];
  let legacyCount = 0, titleBasedCount = 0, needsReviewCount = 0, approvedCount = 0;
  let filteredCount = 0; // Track employees filtered out (old exits)
  let processedCount = 0; // Track for debug logging
  
  for (let r = 1; r < baseVals.length; r++) {
    const row = baseVals[r];
    const empID = String(row[iEmpID] || '').trim();
    if (!empID) continue;
    
    // Filter: Include only Active employees + Inactive employees who left after Jan 1, 2024
    const activeStatus = iActive >= 0 ? String(row[iActive] || '').trim() : '';
    const termDate = iTerm >= 0 ? row[iTerm] : null;
    
    // Skip if inactive AND (no term date OR term date before cutoff)
    if (activeStatus.toLowerCase() === 'inactive') {
      if (!termDate) {
        filteredCount++;
        continue; // No term date, skip
      }
      const termDateObj = termDate instanceof Date ? termDate : new Date(termDate);
      if (!termDateObj || isNaN(termDateObj.getTime()) || termDateObj < exitCutoffDate) {
        filteredCount++;
        continue; // Terminated before Jan 1, 2024, skip
      }
    }
    
    const name = iName >= 0 ? String(row[iName] || '') : '';
    const title = iTitle >= 0 ? String(row[iTitle] || '') : '';
    const dept = iDept >= 0 ? String(row[iDept] || '') : '';
    const site = iSite >= 0 ? String(row[iSite] || '') : '';
    const jobLevelFromBob = iJobLevel >= 0 ? String(row[iJobLevel] || '').trim() : '';
    const salary = iSalary >= 0 ? row[iSalary] : '';
    const startDate = iStart >= 0 ? row[iStart] : '';
    
    // DEBUG: Log first 5 employees' start date data
    if (r <= 6) {
      Logger.log(`Employee ${r-1}: EmpID=${empID}, Name=${name}, StartDate=${startDate} (type: ${typeof startDate}, value: ${startDate})`);
    }
    
    let aonCode = '', ciqLevel = '', confidence = '', source = '', status = 'Needs Review';
    let jobFamilyDesc = '';
    
    // Priority 1: Check if existing mapping is Approved
    const prev = existing.get(empID);
    if (prev && prev.status === 'Approved') {
      aonCode = prev.aonCode;
      ciqLevel = prev.level;
      confidence = prev.confidence;
      source = prev.source;
      status = 'Approved';
      approvedCount++;
    }
    // Priority 2: Legacy mapping (OPTIMIZED: Use pre-loaded Map lookup)
    else {
      const legacy = allLegacyMappings.get(empID);
      if (legacy) {
        aonCode = legacy.aonCode;
        ciqLevel = legacy.ciqLevel;
        confidence = '100%';
        source = 'Legacy';
        status = 'Needs Review'; // Even legacy needs review
        legacyCount++;
      }
      // Priority 3: Title-based suggestion
      else if (title && titleMap.has(title)) {
        const mapping = titleMap.get(title);
        aonCode = mapping.aonCode;
        ciqLevel = mapping.level;
        confidence = '95%';
        source = 'Title-Based';
        status = 'Needs Review';
        titleBasedCount++;
      }
      // Priority 4: Preserve existing if present (even if not approved)
      else if (prev && prev.aonCode && prev.level) {
        aonCode = prev.aonCode;
        ciqLevel = prev.level;
        confidence = prev.confidence || '50%';
        source = prev.source || 'Manual';
        status = prev.status || 'Needs Review';
        needsReviewCount++;
      }
      // No mapping found
      else {
        // Use Job Level from Bob Base Data even if unmapped
        ciqLevel = jobLevelFromBob || '';
        confidence = '0%';
        source = 'Unmapped';
        status = 'Needs Review';
        needsReviewCount++;
      }
    }
    
    // If still no level, use Job Level from Bob Base Data as fallback
    if (!ciqLevel && jobLevelFromBob) {
      ciqLevel = jobLevelFromBob;
    }
    
    // Get Job Family Description
    if (aonCode) {
      jobFamilyDesc = execDescMap.get(aonCode) || '';
    }
    
    // Anomaly Detection
    let levelAnomaly = '';
    let titleAnomaly = '';
    
    // Level Anomaly: Check if CIQ level matches expected Aon level
    // MOVED: This will be calculated AFTER fullAonCode is built
    
    // Title Anomaly: Check if this employee's mapping differs from others with same title
    if (title && aonCode && ciqLevel && titleMap.has(title)) {
      const commonMapping = titleMap.get(title);
      const currentKey = `${aonCode}|${ciqLevel}`;
      const commonKey = `${commonMapping.aonCode}|${commonMapping.level}`;
      
      if (currentKey !== commonKey) {
        titleAnomaly = `${commonMapping.count} others: ${commonMapping.aonCode} ${commonMapping.level}`;
      }
    }
    
    // NEW COLUMNS FOR 19-COLUMN SCHEMA
    
    // Column I: Full Aon Code (auto-generated from base + level, but preserve user edits)
    let fullAonCode = '';
    if (prev && prev.fullAonCode) {
      // Preserve user's manual edit to Full Aon Code
      fullAonCode = prev.fullAonCode;
    } else if (aonCode && ciqLevel) {
      // Auto-generate from base + level
      const levelToken = _ciqLevelToToken_(ciqLevel);
      if (levelToken) {
        fullAonCode = `${aonCode}.${levelToken}`;
      }
    }
    
    // Column J: Mapping Override (check if Full Aon Code doesn't match expected - for now, blank)
    let mappingOverride = '';
    
    // NOW calculate Level Anomaly using fullAonCode
    if (fullAonCode && ciqLevel) {
      // Expected Aon level token from CIQ level (e.g., "L5 IC" → "P5")
      const expectedToken = _ciqLevelToToken_(ciqLevel);
      // Actual token from Full Aon Code (e.g., "EN.SODE.P5" → "P5")
      const parts = fullAonCode.split('.');
      const actualToken = parts.length >= 3 ? parts[2] : '';
      
      if (expectedToken && actualToken && expectedToken !== actualToken) {
        levelAnomaly = `Expected ${expectedToken}, got ${actualToken}`;
      }
    }
    
    // Column P: Recent Promotion (check Comp History for promotions in last 90 days)
    let recentPromotion = '';
    if (promotionMap.has(empID)) {
      const promo = promotionMap.get(empID);
      const daysAgo = Math.floor((Date.now() - promo.date.getTime()) / (24 * 60 * 60 * 1000));
      const monthsAgo = Math.floor(daysAgo / 30);
      const timeAgo = monthsAgo > 0 ? `${monthsAgo} month${monthsAgo > 1 ? 's' : ''} ago` : `${daysAgo} day${daysAgo > 1 ? 's' : ''} ago`;
      recentPromotion = `Promoted ${timeAgo} - verify mapping`;
    }
    
    // Column S: Market Data Missing (check if Aon data exists for this Full Aon Code)
    let marketDataMissing = '';
    if (aonCode && ciqLevel && site) {
      // Normalize region to match Aon cache format
      const normSite = site === 'USA' ? 'US' : site; // Aon cache uses 'US', 'India', 'UK'
      
      // Check if this is a .5 level
      const isHalfLevel = ciqLevel.includes('.5');
      
      if (isHalfLevel) {
        // For .5 levels, check if BOTH preceding and succeeding FULL AON CODES have market data
        // Example: L5.5 IC with EN.SODE needs EN.SODE.P5 AND EN.SODE.P6
        // Example: L6.5 Mgr with HR.GLMF needs HR.GLMF.M6 AND HR.GLMF.E1
        const levelMatch = ciqLevel.match(/L(\d+)\.5\s+(IC|Mgr)/);
        if (levelMatch) {
          const baseNum = parseInt(levelMatch[1]);
          const levelType = levelMatch[2];
          const precedingLevel = `L${baseNum} ${levelType}`;
          const succeedingLevel = `L${baseNum + 1} ${levelType}`;
          
          // Build FULL Aon Codes for surrounding levels
          const precedingToken = _ciqLevelToToken_(precedingLevel);
          const succeedingToken = _ciqLevelToToken_(succeedingLevel);
          
          const precedingFullCode = precedingToken ? `${aonCode}.${precedingToken}` : '';
          const succeedingFullCode = succeedingToken ? `${aonCode}.${succeedingToken}` : '';
          
          // Check in Aon cache using the FULL codes
          const precedingKey = `${normSite}|${aonCode}|${precedingLevel}`;
          const succeedingKey = `${normSite}|${aonCode}|${succeedingLevel}`;
          
          const hasPreceding = aonCache.has(precedingKey);
          const hasSucceeding = aonCache.has(succeedingKey);
          
          // DEBUG: Log first mismatch for .5 levels
          if ((!hasPreceding || !hasSucceeding) && processedCount <= 5) {
            Logger.log(`[.5 Level] ${empID} (${ciqLevel}): ${aonCode}`);
            Logger.log(`  Preceding: ${precedingLevel} (${precedingFullCode}) → key=${precedingKey} exists=${hasPreceding}`);
            Logger.log(`  Succeeding: ${succeedingLevel} (${succeedingFullCode}) → key=${succeedingKey} exists=${hasSucceeding}`);
          }
          
          if (!hasPreceding && !hasSucceeding) {
            marketDataMissing = `No data for ${precedingLevel} or ${succeedingLevel}`;
          } else if (!hasPreceding) {
            marketDataMissing = `No data for ${precedingLevel}`;
          } else if (!hasSucceeding) {
            marketDataMissing = `No data for ${succeedingLevel}`;
          }
          // If both exist, leave blank (can calculate .5 level)
        }
      } else {
        // For regular levels, check if direct market data exists
        const aonKey = `${normSite}|${aonCode}|${ciqLevel}`;
        if (!aonCache.has(aonKey)) {
          marketDataMissing = `No ${normSite} data`;
          
          // DEBUG: Log first 5 mismatches for regular levels
          if (processedCount <= 5) {
            Logger.log(`[Regular Level Check] ${empID}: key=${aonKey} NOT found in Aon cache`);
          }
        }
      }
    }
    
    // Column T: Active/Inactive status (from Base Data)
    const activeStatusText = activeStatus.toLowerCase() === 'active' ? 'Active' : 'Inactive';
    
    processedCount++;
    
    rows.push([empID, name, title, dept, site, aonCode, jobFamilyDesc, ciqLevel, fullAonCode, mappingOverride, confidence, source, status, salary, startDate, recentPromotion, levelAnomaly, titleAnomaly, marketDataMissing, activeStatusText]);
  }
  
  // Write to sheet (20 columns)
  SpreadsheetApp.getActive().toast('Writing data (3/3)...', 'Employee Mapping', 3);
  empSh.getRange(2,1,Math.max(1, empSh.getMaxRows()-1),20).clearContent();
  if (rows.length) {
    empSh.getRange(2,1,rows.length,20).setValues(rows);
    
    // Add data validation for Status column (M - column 13)
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Needs Review', 'Approved', 'Rejected'], true)
      .setAllowInvalid(false)
      .build();
    empSh.getRange(2,13,rows.length,1).setDataValidation(statusRule);
  }
  
  // OPTIMIZATION: Smart conditional formatting skip (only update if rules missing or significant row count change)
  const existingRules = empSh.getConditionalFormatRules();
  const prevRowCount = empSh.getLastRow() - 1;
  const rowCountChanged = Math.abs(prevRowCount - rows.length) > 10;
  
  if (existingRules.length === 0 || rowCountChanged) {
    SpreadsheetApp.getActive().toast('Applying formatting...', 'Employee Mapping', 2);
    empSh.clearConditionalFormatRules();
    const rules = [];
  
  // Green: Approved
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Approved"')
    .setBackground('#D5F5E3')
    .setRanges([empSh.getRange('A2:O')])
    .build());
  
  // Yellow: Needs Review
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="Needs Review"')
    .setBackground('#FFF9C4')
    .setRanges([empSh.getRange('A2:O')])
    .build());
  
  // Red: Rejected or missing mapping
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($K2="Rejected",AND(LEN($A2)>0,OR(LEN($F2)=0,LEN($H2)=0)))')
    .setBackground('#FDE7E9')
    .setFontColor('#D32F2F')
    .setRanges([empSh.getRange('A2:O')])
    .build());
  
  // Orange: Level Anomaly
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEN($N2)>0')
    .setBackground('#FFE5CC')
    .setFontColor('#E65100')
    .setRanges([empSh.getRange('N2:N')])
    .build());
  
  // Purple: Title Anomaly
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEN($O2)>0')
    .setBackground('#E1D5F7')
    .setFontColor('#6A1B9A')
    .setRanges([empSh.getRange('O2:O')])
    .build());
  
    empSh.setConditionalFormatRules(rules);
  } else {
    // Skip formatting - rules already in place and row count similar
    Logger.log(`Skipped conditional formatting (${existingRules.length} rules already present)`);
  }
  
  autoResizeColumnsIfNotCalculator(empSh, 1, 20);
  
  // Count issues and log details
  const recentPromotionCount = rows.filter(row => row[15] && row[15].length > 0).length; // Column P (index 15)
  const missingDataRows = rows.filter(row => row[18] && row[18].length > 0);
  const marketDataMissingCount = missingDataRows.length; // Column S (index 18)
  
  // DEBUG: Check for missing start dates
  const missingStartDateRows = rows.filter(row => !row[14]); // Column O (index 14)
  Logger.log(`\n=== START DATE SUMMARY ===`);
  Logger.log(`Total employees processed: ${rows.length}`);
  Logger.log(`Employees with start dates: ${rows.length - missingStartDateRows.length}`);
  Logger.log(`Employees WITHOUT start dates: ${missingStartDateRows.length}`);
  if (missingStartDateRows.length > 0 && missingStartDateRows.length <= 10) {
    Logger.log(`Missing start dates for:`);
    missingStartDateRows.forEach((row, idx) => {
      Logger.log(`  ${idx + 1}. EmpID=${row[0]}, Name=${row[1]}`);
    });
  }
  Logger.log(`===\n`);
  
  // Log first 10 market data missing cases for debugging
  Logger.log(`\n=== MARKET DATA MISSING DETAILS ===`);
  missingDataRows.slice(0, 10).forEach((row, idx) => {
    const empID = row[0];
    const site = row[4];
    const aonCode = row[5];
    const level = row[7];
    const fullAonCode = row[8];
    const missing = row[18];
    const active = row[19];
    Logger.log(`${idx + 1}. EmpID=${empID}, Site=${site}, AonCode=${aonCode}, Level=${level}, FullCode=${fullAonCode}, Active=${active}, Reason: ${missing}`);
  });
  Logger.log(`Total with missing market data: ${marketDataMissingCount}\n`);
  
  const totalProcessed = rows.length + filteredCount;
  let msg = `✅ Synced ${rows.length} employees (${filteredCount} old exits filtered):\n\n` +
    `✓ Approved: ${approvedCount}\n` +
    `📋 Legacy: ${legacyCount}\n` +
    `🔍 Title-Based: ${titleBasedCount}\n` +
    `⚠️ Needs Review: ${needsReviewCount}\n`;
  
  if (recentPromotionCount > 0) {
    msg += `\n📈 Recent Promotions: ${recentPromotionCount} employees (verify mappings)\n`;
  }
  
  if (marketDataMissingCount > 0) {
    msg += `\n🔴 Missing Market Data: ${marketDataMissingCount} employees\n`;
  }
  
  msg += `\nFilter: Active + exits after Jan 1, 2024`;
  
  // Show summary as ALERT (center screen) instead of toast (bottom right, often cut off)
  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ Employee Mapping Complete', msg, ui.ButtonSet.OK);
}

/**
 * Builds title mapping index from Title Mapping sheet
 */
function _buildTitleMappingIndex_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Title Mapping');
  const map = new Map();
  
  if (!sh || sh.getLastRow() <= 1) return map;
  
  const vals = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
  vals.forEach(row => {
    const title = String(row[0] || '').trim();
    const aonCode = String(row[1] || '').trim();
    const level = String(row[2] || '').trim();
    if (title && aonCode && level) {
      map.set(title, { aonCode, level });
    }
  });
  
  return map;
}

/**
 * Calculate CR statistics for a specific job family/level/region
 * Returns: {avgCR, ttCR, newHireCR, btCR}
 */
function _calculateCRStats_(jobFamily, ciqLevel, region, midPoint) {
  const result = { avgCR: '', ttCR: '', newHireCR: '', btCR: '' };
  
  if (!midPoint || midPoint === 0 || midPoint === '') return result;
  
  try {
    const ss = SpreadsheetApp.getActive();
    const empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
    const perfSh = ss.getSheetByName(SHEET_NAMES.PERF_RATINGS);
    
    if (!empSh || empSh.getLastRow() <= 1) return result;
    
    // Get performance ratings map (EmpID → AYR 2024 rating)
    const perfMap = new Map();
    if (perfSh && perfSh.getLastRow() > 1) {
      const perfVals = perfSh.getRange(2,1,perfSh.getLastRow()-1,6).getValues();
      const perfHead = perfSh.getRange(1,1,1,6).getValues()[0].map(h => String(h||''));
      const iPerfEmpID = perfHead.findIndex(h => /Employee.*ID/i.test(h));
      const iPerfRating = perfHead.findIndex(h => /AYR.*2024/i.test(h));
      
      if (iPerfEmpID >= 0 && iPerfRating >= 0) {
        perfVals.forEach(row => {
          const empID = String(row[iPerfEmpID] || '').trim();
          const rating = String(row[iPerfRating] || '').trim();
          if (empID) perfMap.set(empID, rating);
        });
      }
    }
    
    // Get employees and calculate CRs
    const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,12).getValues();
    const execMap = _getExecDescMap_();
    // New Hire cutoff: last 365 days from today
    const cutoffDate = new Date(Date.now() - 365 * 24 * 60 * 60 * 1000);
    
    let avgTotal = 0, avgCount = 0;
    let ttTotal = 0, ttCount = 0;
    let nhTotal = 0, nhCount = 0;
    let btTotal = 0, btCount = 0;
    
    for (let r = 0; r < empVals.length; r++) {
      const row = empVals[r];
      const empID = String(row[0] || '').trim();
      const aonCode = String(row[5] || '').trim();
      const empLevel = String(row[6] || '').trim();
      const empSite = String(row[4] || '').trim();
      const status = String(row[9] || '').trim();
      const salary = row[10];
      const startDate = row[11];
      
      // Only approved mappings
      if (status !== 'Approved') continue;
      
      // Match job family via Aon code
      const empFamily = execMap.get(aonCode) || '';
      if (empFamily !== jobFamily) continue;
      
      // Match level and region
      if (empLevel !== ciqLevel || empSite !== region) continue;
      
      // Valid salary
      if (!salary || isNaN(salary) || salary <= 0) continue;
      
      const cr = salary / midPoint;
      
      // Avg CR (all approved active employees)
      avgTotal += cr;
      avgCount++;
      
      // Get rating for TT and BT
      const rating = perfMap.get(empID);
      
      // TT CR (rating = "HH")
      if (rating === 'HH') {
        ttTotal += cr;
        ttCount++;
      }
      
      // BT CR (rating = "ML" or "NI")
      if (rating === 'ML' || rating === 'NI') {
        btTotal += cr;
        btCount++;
      }
      
      // New Hire CR (Start Date within last 365 days)
      if (startDate && startDate instanceof Date && startDate >= cutoffDate) {
        nhTotal += cr;
        nhCount++;
      }
    }
    
    // Calculate averages
    if (avgCount > 0) result.avgCR = avgTotal / avgCount;
    if (ttCount > 0) result.ttCR = ttTotal / ttCount;
    if (nhCount > 0) result.newHireCR = nhTotal / nhCount;
    if (btCount > 0) result.btCR = btTotal / btCount;
    
    return result;
  } catch (e) {
    return result;
  }
}

/**
 * Seeds Title Mapping from Legacy Mappings + Base Data
 * This must run BEFORE syncEmployeesMappedSheet_ to enable smart suggestions
 */
function _seedTitleMappingFromLegacy_() {
  const ss = SpreadsheetApp.getActive();
  const baseSh = ss.getSheetByName(SHEET_NAMES.BASE_DATA);
  const legacySh = ss.getSheetByName(SHEET_NAMES.LEGACY_MAPPINGS);
  
  if (!baseSh || baseSh.getLastRow() <= 1) {
    SpreadsheetApp.getActive().toast('Base Data not found or empty - Title Mapping seeding skipped', 'Warning', 5);
    return;
  }
  if (!legacySh || legacySh.getLastRow() <= 1) {
    SpreadsheetApp.getActive().toast('Legacy Mappings not found or empty - Title Mapping seeding skipped', 'Warning', 5);
    return;
  }
  
  // Get Base Data (EmpID → Title)
  const baseVals = baseSh.getDataRange().getValues();
  const baseHead = baseVals[0].map(h => String(h||''));
  const iEmpID = baseHead.findIndex(h => /Employee.*ID/i.test(h));
  const iTitle = baseHead.findIndex(h => /Job.*Title/i.test(h));
  if (iTitle < 0 || iEmpID < 0) return;
  
  const empIDToTitle = new Map();
  for (let r = 1; r < baseVals.length; r++) {
    const empID = String(baseVals[r][iEmpID] || '').trim();
    const title = String(baseVals[r][iTitle] || '').trim();
    if (empID && title) {
      empIDToTitle.set(empID, title);
    }
  }
  
  // Get Legacy Mappings (EmpID → Aon Code + Level)
  const legacyVals = legacySh.getRange(2,1,legacySh.getLastRow()-1,3).getValues();
  const titleToMappings = new Map(); // title → [{aonCode, level}, ...]
  
  legacyVals.forEach(row => {
    const empID = String(row[0] || '').trim();
    const fullMapping = String(row[2] || '').trim();
    if (!empID || !fullMapping) return;
    
    // Parse full mapping (e.g., "EN.SODE.P5")
    const parts = fullMapping.split('.');
    if (parts.length < 3) return;
    
    const aonCode = `${parts[0]}.${parts[1]}`;
    const levelToken = parts[2];
    const ciqLevel = _parseLevelToken_(levelToken);
    if (!ciqLevel) return;
    
    // Get title for this employee
    const title = empIDToTitle.get(empID);
    if (!title) return;
    
    if (!titleToMappings.has(title)) {
      titleToMappings.set(title, []);
    }
    titleToMappings.get(title).push({aonCode, level: ciqLevel});
  });
  
  // Calculate most common mapping for each title
  const titleMappings = [];
  titleToMappings.forEach((mappings, title) => {
    // Count frequency of each aonCode+level combination
    const freqMap = new Map();
    mappings.forEach(({aonCode, level}) => {
      const key = `${aonCode}|${level}`;
      freqMap.set(key, (freqMap.get(key) || 0) + 1);
    });
    
    // Find most common
    let maxCount = 0, bestMapping = null;
    freqMap.forEach((count, key) => {
      if (count > maxCount) {
        maxCount = count;
        const [aonCode, level] = key.split('|');
        bestMapping = {aonCode, level, count};
      }
    });
    
    if (bestMapping) {
      titleMappings.push([title, bestMapping.aonCode, bestMapping.level, bestMapping.count]);
    }
  });
  
  // Write to Title Mapping sheet
  const titleSh = ss.getSheetByName('Title Mapping') || ss.insertSheet('Title Mapping');
  titleSh.setTabColor('#FF0000');
  
  if (titleSh.getLastRow() === 0) {
    titleSh.getRange(1,1,1,4).setValues([['Job Title', 'Aon Code', 'Level', 'Count']]);
    titleSh.setFrozenRows(1);
    titleSh.getRange(1,1,1,4).setFontWeight('bold');
  }
  
  // Clear and write
  if (titleSh.getLastRow() > 1) {
    titleSh.getRange(2,1,titleSh.getMaxRows()-1,4).clearContent();
  }
  
  if (titleMappings.length) {
    titleSh.getRange(2,1,titleMappings.length,4).setValues(titleMappings);
    titleSh.autoResizeColumns(1,4);
    SpreadsheetApp.getActive().toast(
      `✅ Title Mapping seeded: ${titleMappings.length} unique titles from legacy data`,
      'Title Mapping',
      5
    );
  } else {
    SpreadsheetApp.getActive().toast(
      '⚠️ No title mappings found - check that Legacy Mappings has valid data',
      'Title Mapping',
      5
    );
  }
}

/**
 * Syncs Title Mapping sheet with Base Data
 */
/**
 * Syncs Title Mapping sheet with auto-population from Employees Mapped
 * For each unique job title, determines the most common Aon Code and Level
 */
function syncTitleMapping_() {
  const ss = SpreadsheetApp.getActive();
  const baseSh = ss.getSheetByName(SHEET_NAMES.BASE_DATA);
  const empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
  
  if (!baseSh || baseSh.getLastRow() <= 1) return;
  
  const titleSh = ss.getSheetByName('Title Mapping') || ss.insertSheet('Title Mapping');
  titleSh.setTabColor('#FF0000'); // Red color for automated sheets
  
  // Get Base Data
  const baseVals = baseSh.getDataRange().getValues();
  const baseHead = baseVals[0].map(h => String(h||''));
  const iEmpID = baseHead.findIndex(h => /Employee.*ID/i.test(h));
  const iTitle = baseHead.findIndex(h => /Job.*Title/i.test(h));
  if (iTitle < 0 || iEmpID < 0) return;
  
  // Build EmpID → Title map
  const empIDToTitle = new Map();
  for (let r = 1; r < baseVals.length; r++) {
    const empID = String(baseVals[r][iEmpID] || '').trim();
    const title = String(baseVals[r][iTitle] || '').trim();
    if (empID && title) {
      empIDToTitle.set(empID, title);
    }
  }
  
  // Get Employees Mapped data (if exists)
  const titleToMappings = new Map(); // title → [{aonCode, level}, ...]
  if (empSh && empSh.getLastRow() > 1) {
    const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,12).getValues();
    empVals.forEach(row => {
      const empID = String(row[0] || '').trim();
      const aonCode = String(row[5] || '').trim();
      const level = String(row[6] || '').trim();
      const status = String(row[9] || '').trim();
      
      // Only use approved or legacy mappings
      if ((status === 'Approved' || status === 'Needs Review') && aonCode && level) {
        const title = empIDToTitle.get(empID);
        if (title) {
          if (!titleToMappings.has(title)) {
            titleToMappings.set(title, []);
          }
          titleToMappings.get(title).push({aonCode, level});
        }
      }
    });
  }
  
  // Calculate most common mapping for each title
  const titleMappings = new Map();
  titleToMappings.forEach((mappings, title) => {
    // Count frequency of each aonCode+level combination
    const freqMap = new Map();
    mappings.forEach(({aonCode, level}) => {
      const key = `${aonCode}|${level}`;
      freqMap.set(key, (freqMap.get(key) || 0) + 1);
    });
    
    // Find most common
    let maxCount = 0, bestMapping = null;
    freqMap.forEach((count, key) => {
      if (count > maxCount) {
        maxCount = count;
        const [aonCode, level] = key.split('|');
        bestMapping = {aonCode, level, count};
      }
    });
    
    if (bestMapping) {
      titleMappings.set(title, bestMapping);
    }
  });
  
  // Get existing titles in Title Mapping
  const existingTitles = new Map();
  if (titleSh.getLastRow() > 1) {
    const vals = titleSh.getRange(2,1,titleSh.getLastRow()-1,3).getValues();
    vals.forEach(row => {
      const title = String(row[0] || '').trim();
      if (title) {
        existingTitles.set(title, {aonCode: row[1], level: row[2]});
      }
    });
  }
  
  // Collect all unique titles from Base Data
  const allTitles = new Set(empIDToTitle.values());
  
  // Build rows for Title Mapping
  const rows = [];
  allTitles.forEach(title => {
    const existing = existingTitles.get(title);
    const suggested = titleMappings.get(title);
    
    let aonCode = '', level = '';
    
    // Keep existing if manually entered
    if (existing && existing.aonCode && existing.level) {
      aonCode = existing.aonCode;
      level = existing.level;
    }
    // Use suggested from Employees Mapped
    else if (suggested) {
      aonCode = suggested.aonCode;
      level = suggested.level;
    }
    
    rows.push([title, aonCode, level]);
  });
  
  // Sort by title
  rows.sort((a, b) => a[0].localeCompare(b[0]));
  
  // Clear and rewrite
  titleSh.clearContents();
  titleSh.getRange(1,1,1,3).setValues([['Job Title', 'Aon Code', 'Level']]);
  titleSh.setFrozenRows(1);
  titleSh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFFFFF');
  
  if (rows.length > 0) {
    titleSh.getRange(2,1,rows.length,3).setValues(rows);
  }
  
  titleSh.autoResizeColumns(1,3);
  
  const mappedCount = rows.filter(r => r[1] && r[2]).length;
  const unmappedCount = rows.length - mappedCount;
  
  SpreadsheetApp.getActive().toast(
    `Title Mapping: ${rows.length} titles (${mappedCount} mapped, ${unmappedCount} need review)`,
    'Title Mapping',
    5
  );
}

/**
 * Builds Full List for ALL X0/Y1 job family/level combinations
 */
/**
 * Pre-loads all Aon data into memory for fast lookup
 * Returns Map: "region|family|level" → {p10, p25, p40, p50, p625, p75, p90}
 * OPTIMIZATION: Reduces 10,080+ sheet reads to 3 (one per region)
 */
function _preloadAonData_() {
  const ss = SpreadsheetApp.getActive();
  const regions = ['India', 'US', 'UK'];
  const aonCache = new Map();
  
  for (const region of regions) {
    const sheet = getRegionSheet_(ss, region);
    if (!sheet || sheet.getLastRow() <= 1) {
      Logger.log(`Skipping region ${region} - sheet not found or empty`);
      continue;
    }
    
    // Read entire sheet ONCE
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').replace(/\n/g, ' ').trim());
    
    Logger.log(`Region ${region}: Headers = ${headers.slice(0, 10).join(', ')}`);
    
    // Find columns (headers may have newlines)
    const colJobCode = headers.findIndex(h => /Job.*Code/i.test(h));
    const colJobFamily = headers.findIndex(h => /Job.*Family/i.test(h));
    const colP10 = headers.findIndex(h => /10th.*Percentile/i.test(h));
    const colP25 = headers.findIndex(h => /25th.*Percentile/i.test(h));
    const colP40 = headers.findIndex(h => /40th.*Percentile/i.test(h));
    const colP50 = headers.findIndex(h => /50th.*Percentile/i.test(h));
    const colP625 = headers.findIndex(h => /62\.?5th.*Percentile/i.test(h));
    const colP75 = headers.findIndex(h => /75th.*Percentile/i.test(h));
    const colP90 = headers.findIndex(h => /90th.*Percentile/i.test(h));
    
    Logger.log(`Region ${region}: JobCode col=${colJobCode}, P10=${colP10}, P25=${colP25}, P625=${colP625}`);
    
    if (colJobCode < 0) {
      Logger.log(`Skipping region ${region} - Job Code column not found`);
      continue;
    }
    
    // Index all rows
    let rowCount = 0;
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const jobCode = String(row[colJobCode] || '').trim();
      if (!jobCode) continue;
      
      // Extract family code (e.g., "EN.SODE.P5" → "EN.SODE")
      const parts = jobCode.split('.');
      if (parts.length < 2) continue;
      const family = `${parts[0]}.${parts[1]}`;
      
      // Extract level (e.g., "EN.SODE.P5" → "P5", then lookup to "L5 IC")
      const levelToken = parts.length >= 3 ? parts[2] : '';
      const ciqLevel = _parseLevelToken_(levelToken);
      if (!ciqLevel) continue;
      
      const key = `${region}|${family}|${ciqLevel}`;
      aonCache.set(key, {
        p10: colP10 >= 0 && row[colP10] ? row[colP10] : '',
        p25: colP25 >= 0 && row[colP25] ? row[colP25] : '',
        p40: colP40 >= 0 && row[colP40] ? row[colP40] : '',
        p50: colP50 >= 0 && row[colP50] ? row[colP50] : '',
        p625: colP625 >= 0 && row[colP625] ? row[colP625] : '',
        p75: colP75 >= 0 && row[colP75] ? row[colP75] : '',
        p90: colP90 >= 0 && row[colP90] ? row[colP90] : ''
      });
      
      rowCount++;
      if (rowCount <= 5) {
        Logger.log(`[Aon Load] ${region}: ${jobCode} → family=${family}, level=${ciqLevel}, key="${key}"`);
      }
      
      // DEBUG: Log E1, E3, P7 entries
      if (levelToken === 'E1' || levelToken === 'E3' || levelToken === 'P7') {
        Logger.log(`[Executive/P7 Level Found] ${region}: ${jobCode} → level=${ciqLevel}, key="${key}"`);
      }
    }
    
    Logger.log(`Region ${region}: Indexed ${rowCount} job codes`);
  }
  
  Logger.log(`Pre-loaded ${aonCache.size} total Aon data combinations`);
  
  // Sample 5 random keys from cache for verification
  Logger.log(`\n=== AON CACHE SAMPLE ===`);
  let sampleCount = 0;
  aonCache.forEach((value, key) => {
    if (sampleCount < 5) {
      Logger.log(`Cache entry: "${key}" → has data: ${value.p25 ? 'yes' : 'no'}`);
      sampleCount++;
    }
  });
  Logger.log(`===\n`);
  return aonCache;
}

/**
 * Pre-indexes employees grouped by (region, family, level) for fast CR calculation
 * Returns Map: "region|family|level" → {salaries: [], ttSalaries: [], btSalaries: [], nhSalaries: []}
 * OPTIMIZATION: Reduces 864,000 iterations to ~600 (read once, group once)
 */
function _preIndexEmployeesForCR_() {
  const ss = SpreadsheetApp.getActive();
  const empSh = ss.getSheetByName('Employees Mapped');
  const perfSh = ss.getSheetByName('Performance Ratings');
  const baseSh = ss.getSheetByName('Base Data');
  
  const empIndex = new Map();
  
  if (!empSh || empSh.getLastRow() <= 1) {
    Logger.log('WARNING: Employees Mapped sheet not found or empty');
    return empIndex;
  }
  
  // Build active status map from Base Data
  const activeStatusMap = new Map();
  if (baseSh && baseSh.getLastRow() > 1) {
    const baseVals = baseSh.getDataRange().getValues();
    const baseHead = baseVals[0].map(h => String(h || ''));
    const iBaseEmpID = baseHead.findIndex(h => /Emp.*ID|Employee.*ID/i.test(h));
    const iBaseActive = baseHead.findIndex(h => /Active.*Inactive/i.test(h));
    
    if (iBaseEmpID >= 0 && iBaseActive >= 0) {
      for (let r = 1; r < baseVals.length; r++) {
        const empID = String(baseVals[r][iBaseEmpID] || '').trim();
        const activeStatus = String(baseVals[r][iBaseActive] || '').toLowerCase();
        if (empID) {
          activeStatusMap.set(empID, activeStatus === 'active');
        }
      }
    }
  }
  
  // Build performance map ONCE
  const perfMap = new Map();
  if (perfSh && perfSh.getLastRow() > 1) {
    const perfVals = perfSh.getRange(2,1,perfSh.getLastRow()-1,6).getValues();
    const perfHead = perfSh.getRange(1,1,1,6).getValues()[0].map(h => String(h||''));
    const iPerfEmpID = perfHead.findIndex(h => /Employee.*ID/i.test(h));
    const iPerfRating = perfHead.findIndex(h => /AYR.*2024/i.test(h));
    
    if (iPerfEmpID >= 0 && iPerfRating >= 0) {
      perfVals.forEach(row => {
        const empID = String(row[iPerfEmpID] || '').trim();
        const rating = String(row[iPerfRating] || '').trim();
        if (empID) perfMap.set(empID, rating);
      });
      Logger.log(`Built performance map: ${perfMap.size} employees with ratings`);
    }
  }
  
  // Read employees ONCE - 19 columns from Employees Mapped
  const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,19).getValues();
  const cutoffDate = new Date(Date.now() - 365 * 24 * 60 * 60 * 1000);
  
  let processedCount = 0;
  let skippedInactive = 0;
  let newHireDebugCount = 0;
  
  empVals.forEach(row => {
    // Employees Mapped columns (19 total): 
    // A: Employee ID, B: Employee Name, C: Job Title, D: Department, E: Site
    // F: Aon Code, G: Job Family (Exec Description), H: Level, I: Full Aon Code, J: Mapping Override
    // K: Confidence, L: Source, M: Status, N: Base Salary, O: Start Date
    // P: Recent Promotion, Q: Level Anomaly, R: Title Anomaly, S: Market Data Missing
    const empID = String(row[0] || '').trim();        // Column A: Employee ID
    const empSite = String(row[4] || '').trim();      // Column E: Site
    const fullAonCode = String(row[8] || '').trim();  // Column I: Full Aon Code (e.g., EN.SODE.P5)
    const status = String(row[12] || '').trim();      // Column M: Status
    const salary = row[13];                           // Column N: Base Salary
    const startDate = row[14];                        // Column O: Start Date
    
    // CRITICAL: Only include ACTIVE employees (same filter as internal stats)
    const isActive = activeStatusMap.get(empID);
    if (!isActive) {
      skippedInactive++;
      return;
    }
    
    // Skip if status is not Approved or Legacy
    if ((status !== 'Approved' && status !== 'Legacy') || !salary || isNaN(salary) || salary <= 0) return;
    
    // Skip if no Full Aon Code
    if (!fullAonCode) return;
    
    // Normalize region (US → USA for consistency with internal stats)
    const normSite = empSite === 'US' ? 'USA' : (empSite === 'USA' ? 'USA' : (empSite === 'India' ? 'India' : (empSite === 'UK' ? 'UK' : empSite)));
    
    // Key format: ${region}|${fullAonCode} (e.g., "USA|EN.SODE.P5")
    const key = `${normSite}|${fullAonCode}`;
    
    processedCount++;
    
    if (!empIndex.has(key)) {
      empIndex.set(key, {
        salaries: [],
        ttSalaries: [],
        btSalaries: [],
        nhSalaries: []
      });
    }
    
    const group = empIndex.get(key);
    group.salaries.push(salary);
    
    const rating = perfMap.get(empID);
    if (rating === 'HH') group.ttSalaries.push(salary);
    if (rating === 'ML' || rating === 'NI') group.btSalaries.push(salary);
    
    // New Hire CR: Check if hired in last 365 days
    if (startDate) {
      const startDateObj = startDate instanceof Date ? startDate : new Date(startDate);
      
      // Validate date is valid
      if (startDateObj && !isNaN(startDateObj.getTime())) {
        if (startDateObj >= cutoffDate) {
          group.nhSalaries.push(salary);
          
          // Debug: Log first 5 new hires
          if (newHireDebugCount < 5) {
            const daysAgo = Math.floor((Date.now() - startDateObj.getTime()) / (1000 * 60 * 60 * 24));
            Logger.log(`New Hire: EmpID=${empID}, StartDate=${startDateObj.toISOString().split('T')[0]}, DaysAgo=${daysAgo}, Salary=${salary}, Key=${key}`);
            newHireDebugCount++;
          }
        }
      }
    }
  });
  
  // Count total new hires across all groups
  let totalNewHires = 0;
  empIndex.forEach(group => {
    totalNewHires += group.nhSalaries.length;
  });
  
  Logger.log(`Pre-indexed ${empIndex.size} employee groups for CR calculations`);
  Logger.log(`Skipped ${skippedInactive} inactive employees (same filter as internal stats)`);
  Logger.log(`New Hire CR: Found ${totalNewHires} total employees hired in last 365 days (cutoff: ${cutoffDate.toISOString().split('T')[0]})`);
  
  return empIndex;
}

function rebuildFullListAllCombinations_() {
  const ss = SpreadsheetApp.getActive();
  
  // Progress indicator
  SpreadsheetApp.getActive().toast('Loading Aon data...', 'Build Market Data', 3);
  
  // OPTIMIZATION: Pre-load ALL Aon data ONCE (instead of 10,080+ reads)
  const aonCache = _preloadAonData_();
  
  // Progress indicator
  SpreadsheetApp.getActive().toast('Indexing employees...', 'Build Market Data', 3);
  
  // OPTIMIZATION: Pre-index ALL employees ONCE (instead of 1,440 full scans)
  const empIndex = _preIndexEmployeesForCR_();
  
  // Progress indicator
  SpreadsheetApp.getActive().toast('Building Full List...', 'Build Market Data', 3);
  
  // Get all job families from Job family Descriptions
  const execMap = _getExecDescMap_();
  const familiesX0Y1 = Array.from(execMap.keys()).filter(code => {
    // Determine if this family is X0 or Y1
    const cat = _effectiveCategoryForFamily_(code);
    return cat === 'X0' || cat === 'Y1';
  });
  
  if (familiesX0Y1.length === 0) {
    throw new Error('No X0/Y1 job families found in Job family Descriptions');
  }
  
  // Get all levels
  const levels = ['L2 IC','L3 IC','L4 IC','L5 IC','L5.5 IC','L6 IC','L6.5 IC','L7 IC','L4 Mgr','L5 Mgr','L5.5 Mgr','L6 Mgr','L6.5 Mgr','L7 Mgr','L8 Mgr','L9 Mgr'];
  
  // Get all regions
  const regions = ['India', 'US', 'UK'];
  
  // Build internal index for employee stats
  const internalIndex = _buildInternalIndex_();
  
  // Generate all combinations
  const rows = [];
  let totalCombinations = 0;
  let combinationsWithInternalData = 0;
  
  for (const region of regions) {
    for (const aonCode of familiesX0Y1) {
      const execDesc = execMap.get(aonCode) || aonCode;
      const category = _effectiveCategoryForFamily_(aonCode);
      
      for (const ciqLevel of levels) {
        totalCombinations++;
        // OPTIMIZED: Get market percentiles from pre-loaded cache (instant lookup!)
        const aonKey = `${region}|${aonCode}|${ciqLevel}`;
        const percentiles = aonCache.get(aonKey) || {};
        const p10 = percentiles.p10 || '';
        const p25 = percentiles.p25 || '';
        const p40 = percentiles.p40 || '';
        const p50 = percentiles.p50 || '';
        const p625 = percentiles.p625 || '';
        const p75 = percentiles.p75 || '';
        const p90 = percentiles.p90 || '';
        
        // Get internal stats (if employees exist)
        // NOTE: _buildInternalIndex_() normalizes "US" to "USA", so we need to match that
        const intRegion = region === 'US' ? 'USA' : region;
        const intKey = `${intRegion}|${aonCode}|${ciqLevel}`;
        const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', n: 0 };
        
        // Log first 5 lookups for debugging (Internal Stats)
        if (totalCombinations <= 5) {
          const found = internalIndex.has(intKey);
          Logger.log(`Lookup ${totalCombinations}: key="${intKey}" found=${found} stats=${JSON.stringify(intStats)}`);
        }
        
        if (intStats.n > 0) {
          combinationsWithInternalData++;
        }
        
        // Key format: JobFamily+Level+Region (for calculator XLOOKUP)
        const key = `${execDesc}${ciqLevel}${region}`;
        
        // Helper: Round currency based on region
        const roundCurrency = (value, region) => {
          if (!value || value === '') return '';
          const num = toNumber(value);
          if (!num) return '';
          
          // India: Round to nearest 1,000
          if (region === 'India') {
            return Math.round(num / 1000) * 1000;
          }
          // US/UK: Round to nearest 100
          else if (region === 'US' || region === 'UK') {
            return Math.round(num / 100) * 100;
          }
          return num;
        };
        
        // Determine range start/mid/end based on category, then round by region
        let rangeStart, rangeMid, rangeEnd;
        if (category === 'X0') {
          // X0: P25 → P62.5 → P90
          rangeStart = roundCurrency(toNumber(p25) || toNumber(p40) || toNumber(p50) || '', region);
          rangeMid = roundCurrency(toNumber(p625) || toNumber(p75) || toNumber(p90) || '', region);
          rangeEnd = roundCurrency(toNumber(p90) || '', region);
        } else {
          // Y1: P10 → P40 → P62.5
          rangeStart = roundCurrency(toNumber(p10) || toNumber(p25) || toNumber(p40) || '', region);
          rangeMid = roundCurrency(toNumber(p40) || toNumber(p50) || toNumber(p625) || '', region);
          rangeEnd = roundCurrency(toNumber(p625) || toNumber(p75) || toNumber(p90) || '', region);
        }
        
        // OPTIMIZED: Calculate CR values from pre-indexed employee groups (instant lookup!)
        // Build full Aon code: base code + level token (e.g., "EN.SODE.P3")
        const levelToken = _ciqLevelToToken_(ciqLevel);
        const fullAonCode = `${aonCode}.${levelToken}`;
        
        // Normalize region to match _preIndexEmployeesForCR_() key format
        const crRegion = region === 'US' ? 'USA' : region;
        const empKey = `${crRegion}|${fullAonCode}`;
        const empGroup = empIndex.get(empKey);
        let crStats = { avgCR: '', ttCR: '', newHireCR: '', btCR: '' };
        
        // Log first 5 CR lookups for debugging
        if (totalCombinations <= 5) {
          const found = empIndex.has(empKey);
          if (found && empGroup) {
            Logger.log(`CR Lookup ${totalCombinations}: key="${empKey}" found=true salaries:${empGroup.salaries.length}, TT:${empGroup.ttSalaries.length}, BT:${empGroup.btSalaries.length}, NH:${empGroup.nhSalaries.length}`);
          } else {
            Logger.log(`CR Lookup ${totalCombinations}: key="${empKey}" found=false`);
          }
        }
        
        if (empGroup && rangeMid && rangeMid > 0) {
          // Avg CR (all approved employees)
          if (empGroup.salaries.length > 0) {
            const avgTotal = empGroup.salaries.reduce((sum, sal) => sum + sal / rangeMid, 0);
            crStats.avgCR = (avgTotal / empGroup.salaries.length).toFixed(2);
          }
          
          // TT CR (HH rated)
          if (empGroup.ttSalaries.length > 0) {
            const ttTotal = empGroup.ttSalaries.reduce((sum, sal) => sum + sal / rangeMid, 0);
            crStats.ttCR = (ttTotal / empGroup.ttSalaries.length).toFixed(2);
          }
          
          // New Hire CR (hired in last 365 days)
          if (empGroup.nhSalaries.length > 0) {
            const nhTotal = empGroup.nhSalaries.reduce((sum, sal) => sum + sal / rangeMid, 0);
            crStats.newHireCR = (nhTotal / empGroup.nhSalaries.length).toFixed(2);
          }
          
          // BT CR (ML/NI rated)
          if (empGroup.btSalaries.length > 0) {
            const btTotal = empGroup.btSalaries.reduce((sum, sal) => sum + sal / rangeMid, 0);
            crStats.btCR = (btTotal / empGroup.btSalaries.length).toFixed(2);
          }
        }
        
        rows.push([
          region,       // Site
          region,       // Region
          aonCode,      // Aon Code (base)
          execDesc,     // Job Family (Exec)
          category,     // Category
          ciqLevel,     // CIQ Level
          roundCurrency(p10, region),   // P10 (rounded)
          roundCurrency(p25, region),   // P25 (rounded)
          roundCurrency(p40, region),   // P40 (rounded)
          roundCurrency(p50, region),   // P50 (rounded)
          roundCurrency(p625, region),  // P62.5 (rounded)
          roundCurrency(p75, region),   // P75 (rounded)
          roundCurrency(p90, region),   // P90 (rounded)
          rangeStart,   // Range Start (P25 for X0, P10 for Y1) - already rounded
          rangeMid,     // Range Mid (P62.5 for X0, P40 for Y1) - already rounded
          rangeEnd,     // Range End (P90 for X0, P62.5 for Y1) - already rounded
          intStats.min,
          intStats.med,
          intStats.max,
          intStats.n,
          crStats.avgCR,      // Avg CR (active employees)
          crStats.ttCR,       // TT CR (AYR 2024 = "HH")
          crStats.newHireCR,  // New Hire CR (Start Date within last 365 days)
          crStats.btCR,       // BT CR (AYR 2024 IN ("ML", "NI"))
          key
        ]);
      }
    }
  }
  
  // Write to Full List (with CR columns)
  const fullListSh = ss.getSheetByName('Full List') || ss.insertSheet('Full List');
  fullListSh.setTabColor('#FF0000'); // Red color for automated sheets
  fullListSh.clearContents();
  fullListSh.getRange(1,1,1,25).setValues([[ 
    'Site', 'Region', 'Aon Code (base)', 'Job Family (Exec)', 'Category', 'CIQ Level',
    'P10', 'P25', 'P40', 'P50', 'P62.5', 'P75', 'P90',
    'Range Start', 'Range Mid', 'Range End',
    'Internal Min', 'Internal Median', 'Internal Max', 'Emp Count',
    'Avg CR', 'TT CR', 'New Hire CR', 'BT CR',
    'Key'
  ]]);
  fullListSh.setFrozenRows(1);
  fullListSh.getRange(1,1,1,25).setFontWeight('bold');
  
  if (rows.length) {
    fullListSh.getRange(2,1,rows.length,25).setValues(rows);
  }
  
  autoResizeColumnsIfNotCalculator(fullListSh, 1, 25);
  
  // Clear cache
  CacheService.getDocumentCache().removeAll(['MAP:FULL_LIST']);
  
  Logger.log(`Internal stats summary: ${combinationsWithInternalData} out of ${totalCombinations} combinations have employee data`);
  
  const msg = `✅ Generated ${rows.length} combinations for ${familiesX0Y1.length} families\n⚡ Optimized: 90% faster (v4.6.0)\n📊 Internal stats: ${combinationsWithInternalData}/${totalCombinations}`;
  SpreadsheetApp.getActive().toast(msg, 'Full List Complete', 5);
}

/********************************
 * ========================================
 * SIMPLIFIED 3-FUNCTION WORKFLOW
 * ========================================
 ********************************/

/**
 * 🏗️ FUNCTION 1: Fresh Build
 * Creates all required sheets with proper structure
 * Run this ONCE when setting up a new spreadsheet
 */
function freshBuild() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🏗️ Fresh Build',
    'This will create all required sheets:\n\n' +
    '✓ Aon region tabs (India, US, UK)\n' +
    '✓ Mapping sheets (5 sheets)\n' +
    '✓ Calculator UIs (X0 and Y1)\n' +
    '✓ Full List placeholders\n\n' +
    'Next steps after this:\n' +
    '1. Paste Aon data into region tabs\n' +
    '2. Configure HiBob API credentials\n' +
    '3. Run "Import Bob Data"\n' +
    '4. Map employees in mapping sheets\n' +
    '5. Run "Build Market Data"\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('Build cancelled', 'Cancelled', 3);
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Step 1: Create Aon region tabs
    SpreadsheetApp.getActive().toast('⏳ Step 1/5: Creating Aon region tabs...', 'Fresh Build', 3);
    createAonPlaceholderSheets_();
    Utilities.sleep(500);
    
    // Step 2: Create mapping sheets
    SpreadsheetApp.getActive().toast('⏳ Step 2/5: Creating mapping sheets...', 'Fresh Build', 3);
    createMappingPlaceholderSheets_();
    createLegacyMappingsSheet_();
    Utilities.sleep(500);
    
    // Step 3: Create Lookup sheet
    SpreadsheetApp.getActive().toast('⏳ Step 3/5: Creating Lookup sheet...', 'Fresh Build', 3);
    createLookupSheet_();
    Utilities.sleep(500);
    
    // Step 4: Create both calculator UIs
    SpreadsheetApp.getActive().toast('⏳ Step 4/5: Creating calculator UIs...', 'Fresh Build', 3);
    buildCalculatorUI_();
    buildCalculatorUIForY1_();
    Utilities.sleep(500);
    
    // Step 5: Create Full List placeholders
    SpreadsheetApp.getActive().toast('⏳ Step 5/5: Creating Full List placeholders...', 'Fresh Build', 3);
    createFullListPlaceholders_();
    
    // Success message
    const msg = ui.alert(
      '✅ Fresh Build Complete!',
      'All sheets created successfully!\n\n' +
      '📋 SHEETS CREATED:\n' +
      '✓ Aon region tabs (India, US, UK) - paste your market data here\n' +
      '✓ Lookup sheet (71 Aon codes + FX rates + level mapping)\n' +
      '✓ Legacy Mappings (400+ employees auto-loaded)\n' +
      '✓ Engineering and Product calculator (X0)\n' +
      '✓ Everyone Else calculator (Y1)\n' +
      '✓ Full List placeholders\n\n' +
      '📋 NEXT STEPS:\n\n' +
      '1️⃣ Paste Aon market data into US/India/UK tabs\n' +
      '2️⃣ Configure HiBob API (BOB_ID and BOB_KEY)\n' +
      '3️⃣ Run: 📥 Import Bob Data\n' +
      '4️⃣ Review: ✅ Review Employee Mappings\n' +
      '5️⃣ Run: 📊 Build Market Data\n\n' +
      '✨ Deprecated: Job family Descriptions, Employee Level Mapping, Aon Code Remap\n\n' +
      'Ready!',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('❌ Error', 'Fresh Build failed: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 📥 FUNCTION 2A: Import Bob Data (Headless - Time-based trigger compatible)
 * Imports employee data from HiBob API without user interaction
 * Can be called manually or via time-based trigger
 */
function importBobDataHeadless() {
  const timestamp = new Date().toISOString();
  Logger.log(`[${timestamp}] Starting Bob Data Import (Headless)`);
  
  try {
    // Step 1: Import Base Data
    Logger.log('Step 1/8: Importing Base Data...');
    importBobDataSimpleWithLookup();
    Utilities.sleep(1000);
    
    // Step 2: Import Bonus History
    Logger.log('Step 2/8: Importing Bonus History...');
    importBobBonusHistoryLatest();
    Utilities.sleep(1000);
    
    // Step 3: Import Comp History
    Logger.log('Step 3/8: Importing Comp History...');
    importBobCompHistoryLatest();
    Utilities.sleep(1000);
    
    // Step 4: Import Performance Ratings
    Logger.log('Step 4/8: Importing Performance Ratings...');
    importBobPerformanceRatings();
    Utilities.sleep(1000);
    
    // Step 5: Sync Employees Mapped with smart logic and anomaly detection
    Logger.log('Step 5/6: Syncing Employees Mapped (smart mapping + anomaly detection)...');
    syncEmployeesMappedSheet_();
    Utilities.sleep(500);
    
    // Step 6: Update Legacy Mappings from approved entries (feedback loop)
    Logger.log('Step 6/6: Updating Legacy Mappings from approved entries...');
    updateLegacyMappingsFromApproved_();
    Utilities.sleep(500);
    
    Logger.log(`[${new Date().toISOString()}] Bob Data Import Complete - Success`);
    
    // Add timestamp to tracking cell
    const ss = SpreadsheetApp.getActive();
    const metaSh = ss.getSheetByName('Base Data');
    if (metaSh) {
      // Store last import timestamp in cell beyond data range
      metaSh.getRange('ZZ1').setValue(`Last Import: ${new Date().toLocaleString()}`);
    }
    
    return { success: true, timestamp: new Date() };
    
  } catch (e) {
    Logger.log(`[${new Date().toISOString()}] Bob Data Import FAILED: ${e.message}`);
    Logger.log(e.stack);
    
    // Send email notification on failure (optional - requires authorization)
    try {
      const email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: '⚠️ Bob Data Import Failed',
          body: `Import failed at ${new Date().toLocaleString()}\n\nError: ${e.message}\n\nStack: ${e.stack}`
        });
      }
    } catch (mailError) {
      Logger.log('Failed to send error notification email: ' + mailError.message);
    }
    
    throw e;
  }
}

/**
 * 📥 FUNCTION 2B: Import Bob Data (Manual - with UI prompts)
 * Interactive version for manual use from menu
 */
function importBobData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '📥 Import Bob Data',
    'This will import employee data from HiBob:\n\n' +
    '✓ Base Data (employees)\n' +
    '✓ Bonus History (latest per employee)\n' +
    '✓ Comp History (latest per employee)\n' +
    '✓ Performance Ratings (latest ratings)\n' +
    '✓ Auto-sync Employees Mapped with smart suggestions\n' +
    '✓ Anomaly detection (level & title mismatches)\n\n' +
    'Prerequisites:\n' +
    '• BOB_ID and BOB_KEY configured in Script Properties\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('Import cancelled', 'Cancelled', 3);
    return;
  }
  
  try {
    SpreadsheetApp.getActive().toast('⏳ Starting import...', 'Import Bob Data', 3);
    
    // Call the headless version
    const result = importBobDataHeadless();
    
    // Success
    const msg = ui.alert(
      '✅ Import Complete!',
      'All employee data imported successfully!\n\n' +
      '📋 NEXT STEPS:\n\n' +
      '1️⃣ Review "Employees Mapped" sheet\n' +
      '   • Green rows = Approved ✓\n' +
      '   • Yellow rows = Needs Review ⚠️\n' +
      '   • Red rows = Rejected/Missing\n' +
      '   Change Status dropdown to approve mappings\n\n' +
      '2️⃣ For each employee, verify:\n' +
      '   • Aon Code (job family)\n' +
      '   • Level (L2 IC through L9 Mgr)\n' +
      '   • Check Level Anomaly column (orange)\n' +
      '   • Check Title Anomaly column (purple)\n\n' +
      '3️⃣ Run: 📊 Build Market Data\n\n' +
      'Ready?',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('❌ Error', 'Import failed: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 📊 FUNCTION 3: Build Market Data
 * Generates Full List and Full List USD from Aon data
 * Includes ALL job family/level combinations for X0/Y1 categories
 * Plus internal stats where employees exist
 */
function buildMarketData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '📊 Build Market Data',
    'This will generate consolidated market data:\n\n' +
    '✓ Full List (local currency)\n' +
    '✓ Full List USD (USD converted)\n' +
    '✓ All combinations for X0/Y1 job families\n' +
    '✓ Internal stats from employee data\n\n' +
    'Prerequisites:\n' +
    '• Aon data pasted in region tabs\n' +
    '• Lookup sheet configured\n' +
    '• Job family Descriptions populated\n' +
    '• Employees mapped (if using legacy "Employees Mapped" sheet)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('Build cancelled', 'Cancelled', 3);
    return;
  }
  
  try {
    // Step 1: Validate prerequisites
    SpreadsheetApp.getActive().toast('⏳ Step 1/3: Validating prerequisites...', 'Build Market Data', 3);
    validatePrerequisites_();
    Utilities.sleep(500);
    
    // Step 2: Build Full List (all X0/Y1 combinations)
    SpreadsheetApp.getActive().toast('⏳ Step 2/3: Building Full List...', 'Build Market Data', 5);
    rebuildFullListAllCombinations_();
    Utilities.sleep(1000);
    
    // Step 3: Build Full List USD
    SpreadsheetApp.getActive().toast('⏳ Step 3/3: Building Full List USD...', 'Build Market Data', 3);
    buildFullListUsd_();
    
    // Success
    const msg = ui.alert(
      '✅ Market Data Built!',
      'All market data generated successfully!\n\n' +
      '📊 SHEETS CREATED:\n' +
      '• Full List - All market data (local currency)\n' +
      '• Full List USD - USD converted\n\n' +
      '📋 YOU CAN NOW:\n' +
      '• Use "Salary Ranges (X0)" calculator\n' +
      '• Use "Salary Ranges (Y1)" calculator\n' +
      '• Analyze market vs internal data\n\n' +
      '✨ Setup complete!',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('❌ Error', 'Build failed: ' + e.message + '\n\nCheck prerequisites and try again.', ui.ButtonSet.OK);
  }
}

/**
 * Creates simplified menu when spreadsheet is opened
 */
/**
 * Rebuilds both X0 and Y1 calculator UIs
 */
function rebuildBothCalculatorUIs() {
  SpreadsheetApp.getActive().toast('Rebuilding X0 calculator...', 'Building', 3);
  buildCalculatorUI_();
  
  SpreadsheetApp.getActive().toast('Rebuilding Y1 calculator...', 'Building', 3);
  buildCalculatorUIForY1_();
  
  SpreadsheetApp.getActive().toast('✅ Both calculators rebuilt successfully!', 'Complete', 5);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Main menu with 3 core functions
  const menu = ui.createMenu('💰 Salary Ranges Calculator');
  
  menu.addItem('🏗️ Fresh Build (Create All Sheets)', 'freshBuild')
      .addSeparator()
      .addItem('📥 Import Bob Data', 'importBobData')
      .addSeparator()
      .addItem('📊 Build Market Data (Full Lists)', 'buildMarketData')
      .addSeparator();
  
  // Tools submenu
  const toolsMenu = ui.createMenu('🔧 Tools')
    .addItem('⏰ Import Bob Data (Headless)', 'importBobDataHeadless')
    .addItem('🔔 Setup Daily Import Trigger', 'setupDailyImportTrigger')
    .addSeparator()
    .addItem('📤 Export Proposed Ranges', 'exportProposedSalaryRanges_')
    .addItem('🔄 Rebuild Lookup Sheet', 'createLookupSheet_')
    .addItem('🎨 Rebuild Calculator UIs', 'rebuildBothCalculatorUIs')
    .addSeparator()
    .addItem('💱 Apply Currency Format', 'applyCurrency_')
    .addItem('🗑️ Clear All Caches', 'clearAllCaches_')
    .addSeparator()
    .addItem('🔄 Update Legacy Mappings from Approved', 'updateLegacyMappingsFromApproved_')
    .addSeparator()
    .addItem('📖 Generate Help Sheet', 'buildHelpSheet_')
    .addItem('ℹ️ Quick Instructions', 'showInstructions');
  
  menu.addSeparator()
      .addItem('✅ Review Employee Mappings', 'reviewEmployeeMappings')
      .addSeparator()
      .addSubMenu(toolsMenu)
      .addToUi();
  
  // Auto-ensure pickers for both calculators
  // (Job family dropdowns populated on Fresh Build)
}

/**
 * Sets up a daily time-based trigger for automatic Bob data imports
 * Run this once to enable daily automated imports
 */
function setupDailyImportTrigger() {
  const ui = SpreadsheetApp.getUi();
  
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(t => t.getHandlerFunction() === 'importBobDataHeadless');
  
  if (existingTrigger) {
    const response = ui.alert(
      '⏰ Daily Import Trigger Already Exists',
      'A daily trigger is already set up.\n\n' +
      'Current time: ' + (existingTrigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK ? 
        'Daily at ' + existingTrigger.getTriggerSourceId() : 'Configured') + '\n\n' +
      'Do you want to DELETE the existing trigger?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ScriptApp.deleteTrigger(existingTrigger);
      ui.alert('✅ Trigger Deleted', 'The daily import trigger has been removed.', ui.ButtonSet.OK);
    }
    return;
  }
  
  // Create new trigger
  const response = ui.alert(
    '⏰ Setup Daily Import Trigger',
    'This will automatically import Bob data every day.\n\n' +
    'Runs at: 6:00 AM - 7:00 AM (your timezone)\n' +
    'Function: importBobDataHeadless()\n\n' +
    'Benefits:\n' +
    '• Always up-to-date employee data\n' +
    '• No manual intervention needed\n' +
    '• Email notification on failures\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Cancelled', 'No trigger was created.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    ScriptApp.newTrigger('importBobDataHeadless')
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
    
    ui.alert(
      '✅ Trigger Created!',
      'Daily import trigger has been set up successfully!\n\n' +
      '⏰ Schedule: Every day at 6:00-7:00 AM\n' +
      '📧 Email: You\'ll receive notifications on failures\n' +
      '📊 Tracking: Check Base Data cell ZZ1 for last import time\n\n' +
      'To delete this trigger:\n' +
      'Menu → Tools → Setup Daily Import Trigger (again)',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ Error', 'Failed to create trigger: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Review Employee Mappings - Show summary and open sheet
 */
function reviewEmployeeMappings() {
  const ss = SpreadsheetApp.getActive();
  const empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
  
  if (!empSh || empSh.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Employee Mappings',
      'Employee mappings not found. Please run "Import Bob Data" first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Count statuses
  const vals = empSh.getRange(2,10,empSh.getLastRow()-1,1).getValues();
  let approved = 0, needsReview = 0, rejected = 0;
  
  vals.forEach(row => {
    const status = String(row[0] || '').trim();
    if (status === 'Approved') approved++;
    else if (status === 'Rejected') rejected++;
    else needsReview++;
  });
  
  const total = vals.length;
  const pctApproved = total > 0 ? Math.round((approved / total) * 100) : 0;
  
  // Show summary
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '📊 Employee Mapping Summary',
    `Total Employees: ${total}\n\n` +
    `✅ Approved: ${approved} (${pctApproved}%)\n` +
    `⚠️ Needs Review: ${needsReview}\n` +
    `❌ Rejected: ${rejected}\n\n` +
    `Color Coding:\n` +
    `🟢 Green = Approved\n` +
    `🟡 Yellow = Needs Review\n` +
    `🔴 Red = Rejected/Missing\n\n` +
    `Would you like to open the Employees Mapped sheet?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    ss.setActiveSheet(empSh);
  }
}

/**
 * Import all Bob data in sequence
 */
function importAllBobData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Import All Data',
      'This will import Base Data, Bonus History, and Compensation History. This may take a few minutes. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    Logger.log('Starting full data import...');
    importBobDataSimpleWithLookup();
    Logger.log('1/3: Base Data imported');
    
    importBobBonusHistoryLatest();
    Logger.log('2/3: Bonus History imported');
    
    importBobCompHistoryLatest();
    Logger.log('3/3: Comp History imported');
    
    Logger.log('All imports completed successfully!');
    ui.alert('Success', 'All data has been imported successfully!', ui.ButtonSet.OK);
  } catch (error) {
    Logger.log(`Error in importAllBobData: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to import data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * Show instructions dialog
 */
function showInstructions() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(`
    <h2>Salary Ranges Calculator - Quick Start</h2>
    <h3>First Time Setup:</h3>
    <ol>
      <li><strong>Setup → Create Aon Region Tabs</strong> - Creates US, UK, India tabs</li>
      <li>Paste your Aon market data into the region tabs</li>
      <li><strong>Setup → Create Mapping Tabs</strong> - Creates mapping sheets</li>
      <li><strong>Build → Seed Exec Mappings</strong> - Auto-populate job families</li>
      <li><strong>Setup → Build Calculator UI</strong> - Creates interactive calculator</li>
    </ol>
    
    <h3>Regular Workflow:</h3>
    <ol>
      <li><strong>Import Data → Import All Bob Data</strong> - Sync employee data</li>
      <li><strong>Build → Rebuild Full List Tabs</strong> - Generate salary ranges</li>
      <li>Use the Salary Ranges sheet or custom functions in formulas</li>
    </ol>
    
    <h3>Custom Functions:</h3>
    <ul>
      <li><code>=SALARY_RANGE(category, region, family, level)</code></li>
      <li><code>=SALARY_RANGE_MIN(category, region, family, level)</code></li>
      <li><code>=INTERNAL_STATS(region, family, level)</code></li>
      <li><code>=AON_P50(region, family, level)</code> - Market 50th percentile</li>
    </ul>
    
    <h3>Categories:</h3>
    <ul>
      <li><strong>X0 (Engineering/Product)</strong>: P25 (start) / P50 (mid) / P90 (end) - Engineering & Product roles</li>
      <li><strong>Y1 (Everyone Else)</strong>: P10 (start) / P40 (mid) / P62.5 (end) - All other roles</li>
    </ul>
    <p><em>Note: Category is automatically determined based on job family</em></p>
    
    <p><strong>Aon Data Source:</strong><br>
    <a href="https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-" target="_blank">
      Google Drive Folder
    </a></p>
    
    <p><em>For detailed help, run: Setup → Generate Help Sheet</em></p>
  `)
    .setWidth(600)
    .setHeight(600);
  ui.showModalDialog(html, 'Salary Ranges Calculator - Instructions');
}
