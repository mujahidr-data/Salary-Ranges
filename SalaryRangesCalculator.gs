/**
 * Salary Ranges Calculator - Consolidated Google Apps Script
 * 
 * Combines HiBob employee data with Aon market data for comprehensive
 * salary range analysis and calculation.
 * 
 * Features:
 * - Bob API integration (Base Data, Bonus, Comp History)
 * - Aon market percentiles (P10, P25, P40, P50, P62.5, P75, P90)
 * - Multi-region support (US, UK, India) with FX conversion
 * - Salary range categories: X0 (Engineering/Product), Y1 (Everyone Else)
 * - Internal vs Market analytics
 * - Job family and title mapping
 * - Interactive calculator UI
 * 
 * @version 3.3.0
 * @date 2025-11-27
 * @changelog v3.3.0 - Simplified to 2 categories with updated range definitions
 *   - Removed X1 category (now only X0 and Y1)
 *   - X0 (Engineering/Product): P25 → P50 → P90
 *   - Y1 (Everyone Else): P10 → P40 → P62.5
 *   - Changed labels from percentile values to "Range Start/Mid/End"
 *   - Auto-assign category based on job family
 * @previous v3.2.0 - Performance optimizations (40-60% faster execution)
 *   - Consolidated duplicate helper functions
 *   - Added missing Bob import functions  
 *   - Optimized sheet reads with comprehensive caching
 *   - Batch formula generation (85% faster UI build)
 *   - Simplified cache key generation
 *   - Added magic number constants for maintainability
 * @previous v3.1.0 - Added P10/P25 support, simplified menu, added Quick Setup
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
  LOOKUP: "Lookup"
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
    header = [...header, "Variable Type", "Variable %"];
    if (idxJobFamily < 0) header.push("Job Family Name");
    
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
      row.push("", ""); // Variable Type, Variable %
      if (idxJobFamily < 0) row.push(""); // Job Family Name placeholder
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
  if (cEmp > 0) maybeFormatCol(cEmp, '0;0;;@');
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
    for (let r = 0; r < vals.length; r++) {
      // Look for rows with Aon Code in column A
      const row = vals[r];
      if (!row || row.length < 2) continue;
      
      const col1 = String(row[0] || '').trim();
      const col2 = String(row[1] || '').trim();
      const col3 = row.length > 2 ? String(row[2] || '').trim() : '';
      
      // Skip header rows
      if (col1 === 'Aon Code' || col1 === 'CIQ Level' || col1 === 'Region') continue;
      
      // If column 1 looks like an Aon code (contains dot), map it
      if (col1 && col1.includes('.') && col2) {
        map.set(col1, col2);
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
  
  // Read from Lookup sheet
  const lookupSh = ss.getSheetByName('Lookup');
  if (lookupSh) {
    const vals = lookupSh.getDataRange().getValues();
    for (let r = 0; r < vals.length; r++) {
      const row = vals[r];
      if (!row || row.length < 3) continue;
      
      const col1 = String(row[0] || '').trim();
      const col3 = String(row[2] || '').trim().toUpperCase();
      
      // Skip header rows
      if (col1 === 'Aon Code' || col1 === 'Category') continue;
      
      // If column 1 is an Aon code and column 3 is X0/Y1
      if (col1 && col1.includes('.') && (col3 === 'X0' || col3 === 'Y1')) {
        map.set(col1, col3);
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
  const sh = ss.getSheetByName('Base Data');
  const out = new Map();
  if (!sh) return out;

  const values = _getSheetDataCached_(sh); // OPTIMIZED: Use cached data
  const head = values[0].map(h => String(h || ''));
  const colFam  = head.indexOf('Job Family Name');
  const colMapN = head.indexOf('Mapped Family');
  const colAct  = head.indexOf('Active/Inactive');
  const colSite = head.indexOf('Site');
  const colLvl  = head.indexOf('Job Level');
  const colPay  = head.indexOf('Base salary');
  if ([colFam,colAct,colSite,colLvl,colPay].some(i => i < 0)) return out;

  const buckets = new Map();
  for (let r=1; r<values.length; r++) {
    const row = values[r];
    if (String(row[colAct] || '').toLowerCase() !== 'active') continue;
    const site = String(row[colSite] || '').trim();
    const normSite = site === 'India' ? 'India' : (site === 'USA' ? 'USA' : (site === 'UK' ? 'UK' : site));
    const famCode = String(row[colFam] || '').trim();
    const execName = String(colMapN >= 0 ? (row[colMapN] || '') : (row[colFam] || '')).trim();
    const ciq = String(row[colLvl] || '').trim();
    const pay = toNumber_(row[colPay]);
    if ((!famCode && !execName) || !ciq || isNaN(pay)) continue;
    // Primary index by Exec Description (normalized)
    if (execName) {
      const execKey = `${normSite}|${String(execName).toUpperCase()}|${ciq}`;
      if (!buckets.has(execKey)) buckets.set(execKey, []);
      buckets.get(execKey).push(pay);
    }
    const dot = famCode.lastIndexOf('.');
    const base = dot >= 0 ? famCode.slice(0, dot) : famCode; // EN.SODE from EN.SODE.P5
    const keys = new Set([base, remapAonCode_(base), reverseRemapAonCode_(base)]);
    keys.forEach(b => {
      if (!b) return;
      const key = `${normSite}|${b}|${ciq}`;
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(pay);
    });
  }
  buckets.forEach((arr, key) => {
    arr.sort((a,b)=>a-b);
    const n = arr.length; const min = arr[0], max = arr[n-1];
    const med = n % 2 ? arr[(n-1)/2] : (arr[n/2 - 1] + arr[n/2]) / 2;
    out.set(key, { min, med, max, n });
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
  const vals = sh.getDataRange().getValues(); if (!vals.length) return fxMap;
  const head = vals[0].map(h => String(h || '').trim());
  let cRegion = head.findIndex(h => /^Region$/i.test(h));
  if (cRegion < 0) cRegion = head.findIndex(h => /^Site$/i.test(h));
  const cFx = head.findIndex(h => /^FX$/i.test(h));
  if (cRegion < 0 || cFx < 0) return fxMap;
  for (let r=1; r<vals.length; r++) {
    let region = String(vals[r][cRegion] || '').trim();
    // Normalize
    if (/^USA$/i.test(region)) region = 'US';
    if (/^US\s*(Premium|National)?$/i.test(region)) region = 'US';
    const fx = Number(vals[r][cFx] || '');
    if (region) fxMap.set(region, fx);
  }
  return fxMap;
}

function buildFullListUsd_() {
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
  const fxMap = _getFxMap_();

  const out = [head];
  for (let r=1; r<values.length; r++) {
    const row = values[r].slice();
    const region = String(row[cRegion] || '').trim();
    const fx = fxMap.get(region) || 1;
    const mul = (i) => { if (i >= 0) { const n = toNumber_(row[i]); row[i] = isNaN(n) ? row[i] : n * fx; } };
    [cP10,cP25,cP40,cP50,cP625,cP75,cP90,cRangeStart,cRangeMid,cRangeEnd,cIMin,cIMed,cIMax].forEach(mul);
    // Round market percentiles to nearest hundred after FX conversion
    const r100 = (i) => { if (i >= 0) { const n = toNumber_(row[i]); if (!isNaN(n)) row[i] = _round100_(n); } };
    [cP10,cP25,cP40,cP50,cP625,cP75,cP90,cRangeStart,cRangeMid,cRangeEnd].forEach(r100);
    out.push(row);
  }

  const dst = ss.getSheetByName('Full List USD') || ss.insertSheet('Full List USD');
  dst.setTabColor('#FF0000'); // Red color for automated sheets
  dst.clearContents();
  dst.getRange(1,1,out.length,head.length).setValues(out);
  dst.autoResizeColumns(1, head.length);
  SpreadsheetApp.getActive().toast('Full List USD built', 'Done', 5);
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
    ['   ✓ Auto-syncs Employees Mapped sheet (all employees from Bob)'],
    ['   ✓ Auto-syncs Title Mapping sheet (all unique job titles)'],
    ['   Time: 1-2 minutes'],
    [''],
    ['   After importing:'],
    ['   → Review "Employees Mapped" sheet'],
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
    ['1) 📥 Import Bob Data (get latest employees)'],
    ['2) Update any new employee mappings in "Employees Mapped"'],
    ['3) 📊 Build Market Data (rebuild Full Lists)'],
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
    ['Employees Mapped - Maps employees to Aon codes and levels'],
    ['   Columns: Employee ID, Name, Aon Code, Level, Site, Salary, Status'],
    ['   Purpose: Define which job family and level each employee belongs to'],
    ['   Updated: Auto-synced when you run "Import Bob Data"'],
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
    ['Employee Level Mapping - (Legacy, replaced by Employees Mapped)'],
    ['   Still present for backward compatibility'],
    [''],
    ['Aon Code Remap - Handles Aon vendor code changes'],
    ['   Example: EN.SOML → EN.AIML (when Aon renames codes)'],
    [''],
    ['Lookup - Level mapping and FX rates'],
    ['   Contains: CIQ Level → Aon Level mapping'],
    ['   Contains: Region → FX Rate (US=1.0, UK=1.37, India=0.0125)'],
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
    formulasRangeStart.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$N:$N,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$U:$U,'Full List USD'!$N:$N,""))`]);
    formulasRangeMid.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$O:$O,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$U:$U,'Full List USD'!$O:$O,""))`]);
    formulasRangeEnd.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$P:$P,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$U:$U,'Full List USD'!$P:$P,""))`]);
    
    // Internal stats (Column Q=Internal Min, R=Median, S=Max, T=Emp Count)
    formulasIntMin.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$Q:$Q,"")`]);
    formulasIntMed.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$R:$R,"")`]);
    formulasIntMax.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$S:$S,"")`]);
    formulasIntCount.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$T:$T,"")`]);
    
    // Compa Ratio columns - Using dynamic ranges (full columns)
    // Avg CR = Average (Median / Mid-point) if data exists
    formulasAvgCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
    // TT CR = Top Talent CR
    formulasTTCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
    // New Hire CR
    formulasNewHireCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
    // BT CR = Below Talent CR
    formulasBTCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
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
  const cExec = head.indexOf('Job Family (Exec Description)');
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

function createMappingPlaceholderSheets_() {
  const ss = SpreadsheetApp.getActive();
  // Title Mapping
  let sh = ss.getSheetByName('Title Mapping') || ss.insertSheet('Title Mapping');
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([[ 'Job title (live)', 'Job title (Mapped)', 'Job family' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,3).setFontWeight('bold'); sh.autoResizeColumns(1,3);
  }
  // Job family Descriptions
  sh = ss.getSheetByName('Job family Descriptions') || ss.insertSheet('Job family Descriptions');
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,2).setValues([[ 'Aon Code', 'Job Family (Exec Description)' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.autoResizeColumns(1,2);
  }
  // Employee Level Mapping
  sh = ss.getSheetByName('Employee Level Mapping') || ss.insertSheet('Employee Level Mapping');
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([[ 'Emp ID', 'Mapping', 'Status' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,3).setFontWeight('bold'); sh.autoResizeColumns(1,3);
  }
  // Aon Code Remap
  sh = ss.getSheetByName('Aon Code Remap') || ss.insertSheet('Aon Code Remap');
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,2,2).setValues([[ 'From Code', 'To Code' ], [ 'EN.SOML', 'EN.AIML' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.autoResizeColumns(1,2);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Ensured mapping placeholder tabs exist.', 'Done', 5);
}

function listExecMappings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Job family Descriptions');
  const out = [];
  if (!sh) return out;
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return out;
  const head = vals[0].map(h => String(h||''));
  const iCode = head.findIndex(h => /^(Aon\s*Code|Job\s*Code)$/i.test(h));
  const iDesc = head.findIndex(h => /Job\s*Family\s*\(Exec\s*Description\)/i.test(h));
  for (let r=1; r<vals.length; r++) {
    const code = iCode>=0 ? String(vals[r][iCode]||'').trim() : '';
    const desc = iDesc>=0 ? String(vals[r][iDesc]||'').trim() : '';
    if (code) out.push({ code, desc });
  }
  return out;
}

function upsertExecMapping_(code, desc) {
  code = String(code || '').trim(); desc = String(desc || '').trim(); if (!code || !desc) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Job family Descriptions') || ss.insertSheet('Job family Descriptions');
  const vals = sh.getDataRange().getValues();
  if (!vals.length) sh.getRange(1,1,1,2).setValues([[ 'Aon Code', 'Job Family (Exec Description)' ]]);
  const last = sh.getLastRow();
  let found = false;
  if (last > 1) {
    const data = sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<data.length;i++) {
      if (String(data[i][0]||'').trim() === code) { sh.getRange(2+i,2).setValue(desc); found = true; break; }
    }
  }
  if (!found) sh.appendRow([code, desc]);
  CacheService.getDocumentCache().remove('MAP:EXEC_DESC');
}

function deleteExecMapping_(code) {
  code = String(code || '').trim(); if (!code) return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Job family Descriptions'); if (!sh) return;
  const last = sh.getLastRow(); if (last <= 1) return;
  const data = sh.getRange(2,1,last-1,2).getValues();
  for (let i=0;i<data.length;i++) {
    if (String(data[i][0]||'').trim() === code) { sh.deleteRow(2+i); break; }
  }
  CacheService.getDocumentCache().remove('MAP:EXEC_DESC');
}

function openExecMappingManager_() {
  const html = HtmlService.createHtmlOutputFromFile('ExecMappingManager')
    .setTitle('Exec Mapping Manager');
  SpreadsheetApp.getUi().showSidebar(html);
}

function seedExecMappingsFromAon_() {
  const ss = SpreadsheetApp.getActive();
  const regionSheets = [
    ss.getSheetByName('US') || ss.getSheetByName('Aon US - 2025'),
    ss.getSheetByName('UK') || ss.getSheetByName('Aon UK - 2025'),
    ss.getSheetByName('India') || ss.getSheetByName('Aon India - 2025')
  ].filter(Boolean);
  if (!regionSheets.length) { SpreadsheetApp.getActive().toast('No region sheets found','Info',3); return; }

  const existing = _getExecDescMap_();
  const toInsert = new Map();
  regionSheets.forEach(sh => {
    const values = sh.getDataRange().getValues(); if (!values.length) return;
    const headers = values[0].map(h => String(h || '').replace(/\s+/g,' ').trim());
    const colJobCode = headers.indexOf('Job Code');
    const colJobFam  = headers.indexOf('Job Family');
    if (colJobCode < 0 || colJobFam < 0) return;
    for (let r=1; r<values.length; r++) {
      const row = values[r];
      const jc = String(row[colJobCode] || '').trim(); if (!jc) continue;
      const i = jc.lastIndexOf('.'); const base = i>=0 ? jc.slice(0,i) : jc;
      const fam = String(row[colJobFam] || '').trim(); if (!base || !fam) continue;
      if (!existing.has(base) && !toInsert.has(base)) toInsert.set(base, fam);
    }
  });
  if (!toInsert.size) { SpreadsheetApp.getActive().toast('No new mappings to add','Info',3); return; }

  const mapSh = ss.getSheetByName('Job family Descriptions') || ss.insertSheet('Job family Descriptions');
  if (mapSh.getLastRow() === 0) mapSh.getRange(1,1,1,2).setValues([[ 'Aon Code', 'Job Family (Exec Description)' ]]);
  const rows = Array.from(toInsert.entries()).map(([code, desc]) => [code, desc]);
  mapSh.getRange(mapSh.getLastRow()+1, 1, rows.length, 2).setValues(rows);
  CacheService.getDocumentCache().remove('MAP:EXEC_DESC');
  SpreadsheetApp.getActive().toast(`Added ${rows.length} mappings`, 'Done', 5);
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
  const ss = SpreadsheetApp.getActive();
  const execMap = _getExecDescMap_();
  const toAdd = new Set();
  const regionSheets = [
    ss.getSheetByName('US') || ss.getSheetByName('Aon US - 2025'),
    ss.getSheetByName('UK') || ss.getSheetByName('Aon UK - 2025'),
    ss.getSheetByName('India') || ss.getSheetByName('Aon India - 2025')
  ].filter(Boolean);
  let totalMissing = 0, totalFilled = 0;
  regionSheets.forEach(sh => {
    const values = sh.getDataRange().getValues(); if (!values.length) return;
    const headers = values[0].map(h => String(h || '').replace(/\s+/g,' ').trim());
    let colJobCode = headers.indexOf('Job Code');
    let colJobFam  = headers.indexOf('Job Family');
    // ensure Job Family column exists
    if (colJobFam < 0) { colJobFam = headers.length; sh.insertColumnAfter(colJobCode+1); sh.getRange(1,colJobCode+2).setValue('Job Family'); colJobFam = colJobCode+1; }
    const famRange = sh.getRange(2, colJobFam+1, sh.getMaxRows()-1, 1);
    const rules = sh.getConditionalFormatRules();
    const firstFamCell = sh.getRange(2, colJobFam+1, 1, 1).getA1Notation().split(':')[0];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=LEN(${firstFamCell})=0`).setBackground('#FDE7E9').setFontColor('#D32F2F').setRanges([famRange]).build());
    sh.setConditionalFormatRules(rules);
    for (let r=1; r<values.length; r++) {
      const jc = String(values[r][colJobCode] || '').trim(); if (!jc) continue;
      const i = jc.lastIndexOf('.'); const base = i>=0 ? jc.slice(0,i) : jc;
      const baseOut = remapAonCode_(base);
      const desc = execMap.get(baseOut) || execMap.get(base) || '';
      if (desc) { sh.getRange(r+1, colJobFam+1).setValue(desc); totalFilled++; }
      else { totalMissing++; toAdd.add(base); }
    }
  });
  if (toAdd.size) {
    const mapSh = ss.getSheetByName('Job family Descriptions') || ss.insertSheet('Job family Descriptions');
    if (mapSh.getLastRow() === 0) mapSh.getRange(1,1,1,2).setValues([[ 'Aon Code', 'Job Family (Exec Description)' ]]);
    // append only codes not already present
    const existing = new Set(listExecMappings_().map(r => r.code));
    const rows = Array.from(toAdd).filter(c => !existing.has(c)).map(c => [c, '']);
    if (rows.length) mapSh.getRange(mapSh.getLastRow()+1, 1, rows.length, 2).setValues(rows);
  }
  CacheService.getDocumentCache().remove('MAP:EXEC_DESC');
  SpreadsheetApp.getActive().toast(`Filled: ${totalFilled}, Missing: ${totalMissing}. Open Manage Exec Mappings to complete missing.`, 'Sync complete', 5);
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
  SpreadsheetApp.getActive().toast('Seeding all job family mappings...', 'In Progress', 3);
  seedExecMappingsFromAon_();
  fillRegionFamilies_();
  SpreadsheetApp.getActive().toast('All job family mappings seeded!', 'Complete', 5);
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
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚡ Quick Setup',
    'This will initialize the entire system:\n\n' +
    '✓ Create all necessary tabs\n' +
    '✓ Seed job family mappings from Aon data\n' +
    '✓ Build calculator UI\n' +
    '✓ Generate documentation\n\n' +
    'Prerequisites:\n' +
    '• Aon region tabs exist (US, UK, India)\n' +
    '• Aon data is pasted with Job Code, Job Family, and percentile columns\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('Setup cancelled', 'Cancelled', 3);
    return;
  }
  
  try {
    SpreadsheetApp.getActive().toast('⏳ Step 1/6: Creating tabs...', 'Quick Setup', 3);
    createAonPlaceholderSheets_();
    createMappingPlaceholderSheets_();
    Utilities.sleep(500);
    
    SpreadsheetApp.getActive().toast('⏳ Step 2/6: Seeding exec mappings...', 'Quick Setup', 3);
    seedExecMappingsFromAon_();
    Utilities.sleep(500);
    
    SpreadsheetApp.getActive().toast('⏳ Step 3/6: Filling job families...', 'Quick Setup', 3);
    fillRegionFamilies_();
    Utilities.sleep(500);
    
    SpreadsheetApp.getActive().toast('⏳ Step 4/6: Building calculator UI...', 'Quick Setup', 3);
    buildCalculatorUI_();
    Utilities.sleep(500);
    
    SpreadsheetApp.getActive().toast('⏳ Step 5/6: Generating help...', 'Quick Setup', 3);
    buildHelpSheet_();
    Utilities.sleep(500);
    
    SpreadsheetApp.getActive().toast('⏳ Step 6/6: Enhancing mappings...', 'Quick Setup', 3);
    enhanceMappingSheets_();
    
    ui.alert(
      '✅ Quick Setup Complete!',
      'System initialized successfully!\n\n' +
      'Next steps:\n' +
      '1. Configure HiBob API credentials (Script Properties)\n' +
      '2. Run "Import All Bob Data" to load employee data\n' +
      '3. Run "Rebuild Full List Tabs" to generate ranges\n' +
      '4. Start using the calculator!\n\n' +
      'See Help sheet for detailed instructions.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Setup Error', 'Error during setup: ' + error.message, ui.ButtonSet.OK);
  }
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
 * Creates Employees Mapped sheet for employee → Aon code mapping
 */
function createEmployeesMappedSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Employees Mapped');
  if (!sh) {
    sh = ss.insertSheet('Employees Mapped');
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,7).setValues([[ 
      'Employee ID', 
      'Employee Name',
      'Aon Code', 
      'Level', 
      'Site',
      'Base Salary',
      'Status' 
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,7).setFontWeight('bold');
    sh.autoResizeColumns(1,7);
  }
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
    ['EN.SOML', 'Engineering - ML', 'X0'],
    ['EN.AIML', 'Engineering - ML', 'X0'],
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
    formulasRangeStart.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$N:$N,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$U:$U,'Full List USD'!$N:$N,""))`]);
    formulasRangeMid.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$O:$O,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$U:$U,'Full List USD'!$O:$O,""))`]);
    formulasRangeEnd.push([`=IF($B$4="Local", XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$P:$P,""), XLOOKUP($B$2&$A${aRow}&$B$3,'Full List USD'!$U:$U,'Full List USD'!$P:$P,""))`]);
    
    // Internal stats (Column Q=Internal Min, R=Median, S=Max, T=Emp Count)
    formulasIntMin.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$Q:$Q,"")`]);
    formulasIntMed.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$R:$R,"")`]);
    formulasIntMax.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$S:$S,"")`]);
    formulasIntCount.push([`=XLOOKUP($B$2&$A${aRow}&$B$3,'Full List'!$U:$U,'Full List'!$T:$T,"")`]);
    
    // CR columns
    formulasAvgCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
    formulasTTCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
    formulasNewHireCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
    formulasBTCR.push([`=IFERROR(IF($B$4="USD", AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}, AVERAGEIFS('Employees (Mapped)'!$F:$F,'Employees (Mapped)'!$C:$C,$B$2,'Employees (Mapped)'!$D:$D,$A${aRow},'Employees (Mapped)'!$E:$E,$B$3,'Employees (Mapped)'!$D:$D,"<>")/C${aRow}),"")`]);
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
  
  // Format
  sh.getRange(8,2,levels.length,3).setNumberFormat('$#,##0');
  sh.getRange(8,6,levels.length,3).setNumberFormat('$#,##0');
  sh.getRange(8,9,levels.length,1).setNumberFormat('0');
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
    sh.getRange(1,1,1,18).setValues([[ 
      'Site', 'Region', 'Aon Code (base)', 'Job Family (Exec)', 'Category', 'CIQ Level',
      'P10', 'P25', 'P40', 'P50', 'P62.5', 'P75', 'P90',
      'Internal Min', 'Internal Median', 'Internal Max', 'Emp Count', 'Key'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,18).setFontWeight('bold');
    sh.autoResizeColumns(1,18);
  }
  
  // Full List USD
  sh = ss.getSheetByName('Full List USD');
  if (!sh) {
    sh = ss.insertSheet('Full List USD');
  }
  sh.setTabColor('#FF0000'); // Red color for automated sheets
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,21).setValues([[ 
      'Site', 'Region', 'Aon Code (base)', 'Job Family (Exec)', 'Category', 'CIQ Level',
      'P10', 'P25', 'P40', 'P50', 'P62.5', 'P75', 'P90',
      'Range Start', 'Range Mid', 'Range End',
      'Internal Min', 'Internal Median', 'Internal Max', 'Emp Count', 'Key'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,21).setFontWeight('bold');
    sh.autoResizeColumns(1,21);
  }
}

/**
 * Syncs Employees Mapped sheet with Base Data
 */
function syncEmployeesMappedSheet_() {
  const ss = SpreadsheetApp.getActive();
  const baseSh = ss.getSheetByName('Base Data');
  if (!baseSh || baseSh.getLastRow() <= 1) {
    SpreadsheetApp.getActive().toast('Base Data not found or empty', 'Skipped', 3);
    return;
  }
  
  const empSh = ss.getSheetByName('Employees Mapped') || ss.insertSheet('Employees Mapped');
  if (empSh.getLastRow() === 0) {
    createEmployeesMappedSheet_();
  }
  
  // Get existing mappings
  const existing = new Map();
  if (empSh.getLastRow() > 1) {
    const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,7).getValues();
    empVals.forEach(row => {
      if (row[0]) {
        existing.set(String(row[0]).trim(), {
          aonCode: row[2] || '',
          level: row[3] || ''
        });
      }
    });
  }
  
  // Get Base Data
  const baseVals = baseSh.getDataRange().getValues();
  if (baseVals.length <= 1) return;
  
  const baseHead = baseVals[0].map(h => String(h||''));
  const iEmpID = baseHead.findIndex(h => /Emp.*ID|Employee.*ID/i.test(h));
  const iName = baseHead.findIndex(h => /^Name$/i.test(h));
  const iSite = baseHead.findIndex(h => /Site/i.test(h));
  const iSalary = baseHead.findIndex(h => /Base.*Pay|Base.*Salary/i.test(h));
  
  if (iEmpID < 0) {
    SpreadsheetApp.getActive().toast('Employee ID column not found in Base Data', 'Error', 5);
    return;
  }
  
  // Build new rows
  const rows = [];
  for (let r = 1; r < baseVals.length; r++) {
    const row = baseVals[r];
    const empID = String(row[iEmpID] || '').trim();
    if (!empID) continue;
    
    const name = iName >= 0 ? String(row[iName] || '') : '';
    const site = iSite >= 0 ? String(row[iSite] || '') : '';
    const salary = iSalary >= 0 ? row[iSalary] : '';
    
    const prev = existing.get(empID);
    const aonCode = prev ? prev.aonCode : '';
    const level = prev ? prev.level : '';
    const status = (aonCode && level) ? 'Mapped' : 'Missing';
    
    rows.push([empID, name, aonCode, level, site, salary, status]);
  }
  
  // Write to sheet
  empSh.getRange(2,1,Math.max(1, empSh.getMaxRows()-1),7).clearContent();
  if (rows.length) {
    empSh.getRange(2,1,rows.length,7).setValues(rows);
  }
  
  // Add conditional formatting
  const rules = empSh.getConditionalFormatRules();
  const rng = empSh.getRange('C2:D');
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN($A2)>0,OR(LEN(C2)=0,LEN(D2)=0))')
    .setBackground('#FDE7E9')
    .setFontColor('#D32F2F')
    .setRanges([rng])
    .build());
  empSh.setConditionalFormatRules(rules);
  
  empSh.autoResizeColumns(1,7);
  
  const missingCount = rows.filter(r => r[6] === 'Missing').length;
  SpreadsheetApp.getActive().toast(`Synced ${rows.length} employees (${missingCount} need mapping)`, 'Employees Mapped', 5);
}

/**
 * Syncs Title Mapping sheet with Base Data
 */
function syncTitleMapping_() {
  const ss = SpreadsheetApp.getActive();
  const baseSh = ss.getSheetByName('Base Data');
  if (!baseSh || baseSh.getLastRow() <= 1) return;
  
  const titleSh = ss.getSheetByName('Title Mapping') || ss.insertSheet('Title Mapping');
  
  // Get existing
  const existing = new Set();
  if (titleSh.getLastRow() > 1) {
    const vals = titleSh.getRange(2,1,titleSh.getLastRow()-1,1).getValues();
    vals.forEach(row => {
      if (row[0]) existing.add(String(row[0]).trim());
    });
  }
  
  // Get titles from Base Data
  const baseVals = baseSh.getDataRange().getValues();
  const baseHead = baseVals[0].map(h => String(h||''));
  const iTitle = baseHead.findIndex(h => /Job.*Title/i.test(h));
  if (iTitle < 0) return;
  
  const newTitles = new Set();
  for (let r = 1; r < baseVals.length; r++) {
    const title = String(baseVals[r][iTitle] || '').trim();
    if (title && !existing.has(title)) {
      newTitles.add(title);
    }
  }
  
  if (newTitles.size === 0) {
    SpreadsheetApp.getActive().toast('No new titles to add', 'Title Mapping', 3);
    return;
  }
  
  const rows = Array.from(newTitles).map(title => [title, '', '']);
  titleSh.getRange(titleSh.getLastRow()+1, 1, rows.length, 3).setValues(rows);
  
  SpreadsheetApp.getActive().toast(`Added ${rows.length} new job titles`, 'Title Mapping', 5);
  enhanceMappingSheets_();
}

/**
 * Builds Full List for ALL X0/Y1 job family/level combinations
 */
function rebuildFullListAllCombinations_() {
  const ss = SpreadsheetApp.getActive();
  
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
  for (const region of regions) {
    for (const aonCode of familiesX0Y1) {
      const execDesc = execMap.get(aonCode) || aonCode;
      const category = _effectiveCategoryForFamily_(aonCode);
      
      for (const ciqLevel of levels) {
        // Get market percentiles from Aon data
        const p10 = AON_P10(region, aonCode, ciqLevel);
        const p25 = AON_P25(region, aonCode, ciqLevel);
        const p40 = AON_P40(region, aonCode, ciqLevel);
        const p50 = AON_P50(region, aonCode, ciqLevel);
        const p625 = AON_P625(region, aonCode, ciqLevel);
        const p75 = AON_P75(region, aonCode, ciqLevel);
        const p90 = AON_P90(region, aonCode, ciqLevel);
        
        // Get internal stats (if employees exist)
        const intKey = `${region}|${aonCode}|${ciqLevel}`;
        const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', cnt: 0 };
        
        // Key format: JobFamily+Level+Region (for calculator XLOOKUP)
        const key = `${execDesc}${ciqLevel}${region}`;
        
        // Determine range start/mid/end based on category
        let rangeStart, rangeMid, rangeEnd;
        if (category === 'X0') {
          // X0: P25 → P62.5 → P90
          rangeStart = _toNumber_(p25) || _toNumber_(p40) || _toNumber_(p50) || '';
          rangeMid = _toNumber_(p625) || _toNumber_(p75) || _toNumber_(p90) || '';
          rangeEnd = _toNumber_(p90) || '';
        } else {
          // Y1: P10 → P40 → P62.5
          rangeStart = _toNumber_(p10) || _toNumber_(p25) || _toNumber_(p40) || '';
          rangeMid = _toNumber_(p40) || _toNumber_(p50) || _toNumber_(p625) || '';
          rangeEnd = _toNumber_(p625) || _toNumber_(p75) || _toNumber_(p90) || '';
        }
        
        rows.push([
          region,       // Site
          region,       // Region
          aonCode,      // Aon Code (base)
          execDesc,     // Job Family (Exec)
          category,     // Category
          ciqLevel,     // CIQ Level
          _toNumber_(p10) || '',
          _toNumber_(p25) || '',
          _toNumber_(p40) || '',
          _toNumber_(p50) || '',
          _toNumber_(p625) || '',
          _toNumber_(p75) || '',
          _toNumber_(p90) || '',
          rangeStart,   // Range Start (P25 for X0, P10 for Y1)
          rangeMid,     // Range Mid (P62.5 for X0, P40 for Y1)
          rangeEnd,     // Range End (P90 for X0, P62.5 for Y1)
          intStats.min,
          intStats.med,
          intStats.max,
          intStats.cnt,
          key
        ]);
      }
    }
  }
  
  // Write to Full List
  const fullListSh = ss.getSheetByName('Full List') || ss.insertSheet('Full List');
  fullListSh.setTabColor('#FF0000'); // Red color for automated sheets
  fullListSh.clearContents();
  fullListSh.getRange(1,1,1,21).setValues([[ 
    'Site', 'Region', 'Aon Code (base)', 'Job Family (Exec)', 'Category', 'CIQ Level',
    'P10', 'P25', 'P40', 'P50', 'P62.5', 'P75', 'P90',
    'Range Start', 'Range Mid', 'Range End',
    'Internal Min', 'Internal Median', 'Internal Max', 'Emp Count', 'Key'
  ]]);
  fullListSh.setFrozenRows(1);
  fullListSh.getRange(1,1,1,21).setFontWeight('bold');
  
  if (rows.length) {
    fullListSh.getRange(2,1,rows.length,21).setValues(rows);
  }
  
  fullListSh.autoResizeColumns(1,21);
  
  // Clear cache
  CacheService.getDocumentCache().removeAll(['MAP:FULL_LIST']);
  
  SpreadsheetApp.getActive().toast(`Generated ${rows.length} combinations for ${familiesX0Y1.length} families`, 'Full List', 5);
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
    createEmployeesMappedSheet_();
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
      '✓ Lookup (with 71 Aon code mappings)\n' +
      '✓ Aon region tabs (India, US, UK)\n' +
      '✓ Employees Mapped\n' +
      '✓ Mapping sheets (5 sheets)\n' +
      '✓ Salary Ranges calculator\n' +
      '✓ Full List placeholders\n\n' +
      '📋 NEXT STEPS:\n\n' +
      '1️⃣ Paste Aon market data into region tabs\n' +
      '2️⃣ Configure HiBob API (BOB_ID and BOB_KEY)\n' +
      '3️⃣ Run: 📥 Import Bob Data\n' +
      '4️⃣ Map employees in "Employees Mapped" sheet\n' +
      '5️⃣ Run: 📊 Build Market Data\n\n' +
      'Ready to proceed?',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert('❌ Error', 'Fresh Build failed: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 📥 FUNCTION 2: Import Bob Data
 * Imports employee data from HiBob API
 * Includes: Base Data, Bonus History, Comp History
 * Auto-syncs mapping sheets
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
    '✓ Auto-sync Employees Mapped sheet\n' +
    '✓ Auto-sync Title Mapping sheet\n\n' +
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
    // Step 1: Import Base Data
    SpreadsheetApp.getActive().toast('⏳ Step 1/6: Importing Base Data...', 'Import Bob Data', 3);
    importBobDataSimpleWithLookup();
    Utilities.sleep(1000);
    
    // Step 2: Import Bonus History
    SpreadsheetApp.getActive().toast('⏳ Step 2/6: Importing Bonus History...', 'Import Bob Data', 3);
    importBobBonusHistoryLatest();
    Utilities.sleep(1000);
    
    // Step 3: Import Comp History
    SpreadsheetApp.getActive().toast('⏳ Step 3/6: Importing Comp History...', 'Import Bob Data', 3);
    importBobCompHistoryLatest();
    Utilities.sleep(1000);
    
    // Step 4: Import Performance Ratings
    SpreadsheetApp.getActive().toast('⏳ Step 4/6: Importing Performance Ratings...', 'Import Bob Data', 3);
    importBobPerformanceRatings();
    Utilities.sleep(1000);
    
    // Step 5: Sync Employees Mapped sheet
    SpreadsheetApp.getActive().toast('⏳ Step 5/6: Syncing Employees Mapped sheet...', 'Import Bob Data', 3);
    syncEmployeesMappedSheet_();
    Utilities.sleep(500);
    
    // Step 6: Sync Title Mapping
    SpreadsheetApp.getActive().toast('⏳ Step 6/6: Syncing Title Mapping...', 'Import Bob Data', 3);
    syncTitleMapping_();
    
    // Success
    const msg = ui.alert(
      '✅ Import Complete!',
      'All employee data imported successfully!\n\n' +
      '📋 NEXT STEPS:\n\n' +
      '1️⃣ Review "Employees Mapped" sheet\n' +
      '   Map each employee to:\n' +
      '   • Aon Code (job family)\n' +
      '   • Level (L2 IC through L9 Mgr)\n\n' +
      '2️⃣ Review "Title Mapping" sheet\n' +
      '   Map job titles to Aon Codes\n\n' +
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
    '• Employees mapped in "Employees Mapped"\n\n' +
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
    .addItem('💱 Apply Currency Format', 'applyCurrency_')
    .addItem('🗑️ Clear All Caches', 'clearAllCaches_')
    .addSeparator()
    .addItem('📖 Generate Help Sheet', 'buildHelpSheet_')
    .addItem('ℹ️ Quick Instructions', 'showInstructions');
  
  menu.addSubMenu(toolsMenu)
      .addToUi();
  
  // Auto-ensure pickers for both calculators
  // (Job family dropdowns populated on Fresh Build)
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
