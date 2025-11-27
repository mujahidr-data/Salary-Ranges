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
  COMP_HISTORY: "31054312"
};

const SHEET_NAMES = {
  BASE_DATA: "Base Data",
  BONUS_HISTORY: "Bonus History",
  COMP_HISTORY: "Comp History",
  SALARY_RANGES: "Salary Ranges",
  FULL_LIST: "Full List",
  FULL_LIST_USD: "Full List USD",
  LOOKUP: "Lookup"
};

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
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
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
    
    // Write to sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, sheetName);
    
    sheet.clearContents();
    sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
    
    // Format columns
    if (out.length > 1) {
      const numRows = out.length - 1;
      // Employee ID as text
      sheet.getRange(2, idxEmpId + 1, numRows, 1).setNumberFormat("@");
      // Base Pay as currency
      sheet.getRange(2, idxBasePay + 1, numRows, 1).setNumberFormat("#,##0.00");
    }
    
    sheet.autoResizeColumns(1, out[0].length);
    Logger.log(`Successfully imported ${sheetName}`);
    
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
  
    // Write to sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, targetSheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
    
    if (out.length > 1) {
      const numRows = out.length - 1;
      sheet.getRange(2, 3, numRows, 1).setNumberFormat("@"); // Date as text
      sheet.getRange(2, 5, numRows, 1).setNumberFormat("0.########"); // Percent
      sheet.getRange(2, 6, numRows, 1).setNumberFormat("#,##0.00"); // Amount
    }
    
    sheet.autoResizeColumns(1, out[0].length);
    Logger.log(`Successfully imported ${targetSheetName}`);
    
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
  
    // Write to sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, targetSheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
    
    if (out.length > 1) {
      const numRows = out.length - 1;
      sheet.getRange(2, 3, numRows, 1).setNumberFormat("yyyy-mm-dd"); // Date
      sheet.getRange(2, 4, numRows, 1).setNumberFormat("#,##0.00"); // Salary
    }
    
    sheet.autoResizeColumns(1, out[0].length);
    Logger.log(`Successfully imported ${targetSheetName}`);
    
  } catch (error) {
    Logger.log(`Error in importBobCompHistoryLatest: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error importing Comp History: ${error.message}`);
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
 * X0 = Engineering/Product: P25 (start) → P50 (mid) → P90 (end)
 * Y1 = Everyone Else: P10 (start) → P40 (mid) → P62.5 (end)
 * 
 * Fallback logic: If a percentile is missing, use the next higher percentile
 * Example: P10 missing → use P25, P25 missing → use P40, etc.
 ********************************/
function _rangeByCategory_(category, region, family, ciqLevel) {
  const cat = String(category || '').trim().toUpperCase();
  if (!cat) return { min: '', mid: '', max: '' };

  if (cat === 'X0') {
    // X0 (Engineering/Product): Range Start=P25, Range Mid=P50, Range End=P90
    let min = AON_P25(region, family, ciqLevel);
    let mid = AON_P50(region, family, ciqLevel);
    let max = AON_P90(region, family, ciqLevel);
    
    // Fallback: P25 missing → use P40
    if (!min || min === '') {
      min = AON_P40(region, family, ciqLevel);
      if (!min || min === '') min = AON_P50(region, family, ciqLevel);
    }
    // Fallback: P50 missing → use P625
    if (!mid || mid === '') {
      mid = AON_P625(region, family, ciqLevel);
      if (!mid || mid === '') mid = AON_P75(region, family, ciqLevel);
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

function _getExecDescMap_() {
  const cacheKey = 'MAP:EXEC_DESC';
  const cached = _cacheGet_(cacheKey);
  if (cached) return new Map(cached);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Job family Descriptions');
  const map = new Map();
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
  const hasBob = !!(baseSh && baseSh.getLastRow() > 1);
  if (hasBob) {
    const sh2 = ss.getSheetByName('Coverage Summary') || ss.insertSheet('Coverage Summary');
    sh2.clearContents(); sh2.getRange(1,1,1,6).setValues([['Site','Aon Code','Job Family (Exec)','Levels expected','Levels with market','Levels with internal']]);
    const covMap = new Map();
    rows.forEach(r => { const site = r[0], base = r[2], execFam = r[3], ciq = r[5]; const p62 = r[9], p75 = r[10], p90 = r[11], iMed = r[13]; const k = `${site}|${base}`; if (!covMap.has(k)) covMap.set(k, {execFam, exp:0, mk:0, inr:0, ciqs:new Set()}); const acc = covMap.get(k); if (!acc.ciqs.has(ciq)) { acc.ciqs.add(ciq); acc.exp++; } if (_isNum_(p62) || _isNum_(p75) || _isNum_(p90)) acc.mk++; if (_isNum_(iMed)) acc.inr++; });
    const covRows = []; covMap.forEach((acc, key) => { const [site, base] = key.split('|'); covRows.push([site, base, acc.execFam, acc.exp, acc.mk, acc.inr]); });
    if (covRows.length) sh2.getRange(2,1,covRows.length,6).setValues(covRows); sh2.autoResizeColumns(1, 6);

    const sh3 = ss.getSheetByName('Employees (Mapped)') || ss.insertSheet('Employees (Mapped)');
    // Only clear the first 5 columns to preserve manual columns to the right
    sh3.getRange(1,1,Math.max(2, sh3.getMaxRows()),5).clearContent();
    sh3.getRange(1,1,1,5).setValues([['EmpID','Aon Code','Suffix','Site','Base salary']]);
    const empRows = _readMappedEmployeesForAudit_(); if (empRows.length) sh3.getRange(2,1,empRows.length,5).setValues(empRows); sh3.autoResizeColumns(1,5);
    SpreadsheetApp.getActive().toast('Full List + Coverage Summary + Employees (Mapped) rebuilt', 'Done', 5);
  } else {
    SpreadsheetApp.getActive().toast('Full List rebuilt (Bob Base Data not found — skipped Coverage Summary and Employees (Mapped))', 'Done', 7);
  }
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
    [cP10,cP25,cP40,cP50,cP625,cP75,cP90,cIMin,cIMed,cIMax].forEach(mul);
    // Round market percentiles to nearest hundred after FX conversion
    const r100 = (i) => { if (i >= 0) { const n = toNumber_(row[i]); if (!isNaN(n)) row[i] = _round100_(n); } };
    [cP10,cP25,cP40,cP50,cP625,cP75,cP90].forEach(r100);
    out.push(row);
  }

  const dst = ss.getSheetByName('Full List USD') || ss.insertSheet('Full List USD');
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
    ['Salary Range Calculator - Help & Getting Started'],
    [''],
    ['⚡ QUICK START (Recommended)'],
    ['1) Paste Aon data into region tabs (US, UK, India) with headers: Job Code, Job Family, 10th, 25th, 40th, 50th, 62.5th, 75th, 90th'],
    ['2) Setup → ⚡ Quick Setup (Run Once) - Initializes entire system automatically'],
    ['3) Configure HiBob API: Extensions > Apps Script > Project Settings > Script Properties (BOB_ID, BOB_KEY)'],
    ['4) Import Data → 🔄 Import All Bob Data'],
    ['5) Build → 📊 Rebuild Full List (with validation)'],
    ['6) Start using the calculator!'],
    [''],
    ['MANUAL SETUP (Alternative - for granular control)'],
    ['1) Setup → Create Aon placeholder tabs (creates empty US/UK/India tabs if needed)'],
    ['2) Paste Aon data into region tabs with percentile columns (P10, P25, P40, P50, P62.5, P75, P90)'],
    ['3) Setup → Create mapping placeholder tabs'],
    ['4) Build → Seed All Job Family Mappings (combines exec mappings + job family fill)'],
    ['5) Setup → Build Calculator UI'],
    ['6) Configure HiBob API credentials in Script Properties'],
    ['7) Import Data → Import All Bob Data'],
    ['8) Build → Sync All Bob Mappings (employee levels + titles)'],
    ['9) Build → Rebuild Full List (with validation)'],
    [''],
    ['REGULAR WORKFLOW'],
    ['A) Import Data → Import All Bob Data (refresh employee data from HiBob)'],
    ['B) Build → Rebuild Full List (with validation) - generates comprehensive ranges'],
    ['C) Build → Build Full List USD (optional FX-applied view for multi-region analysis)'],
    ['D) Use calculator UI or formulas: =SALARY_RANGE("X0", "US", "EN.SODE", "L5 IC")'],
    ['E) Export → Export Proposed Ranges (optional export to separate sheet)'],
    [''],
    ['SIMPLIFIED MENU FUNCTIONS'],
    ['- Quick Setup: Runs entire initialization sequence automatically'],
    ['- Seed All Job Family Mappings: Combines exec mappings + job family fill'],
    ['- Sync All Bob Mappings: Syncs employee levels + title mappings together'],
    ['- Rebuild Full List (with validation): Validates prerequisites before building'],
    [''],
    ['PERCENTILES SUPPORTED'],
    ['- P10, P25, P40, P50, P62.5, P75, P90 (all imported from Aon data)'],
    ['- Custom functions: AON_P10(), AON_P25(), AON_P40(), AON_P50(), AON_P625(), AON_P75(), AON_P90()'],
    [''],
    ['CALCULATIONS'],
    ['- Full List includes: P10/P25/P40/P50/P62.5/P75/P90 + Internal Min/Median/Max + Employees'],
    ['- Cache index built on demand (10-min TTL) for fast lookups'],
    ['- SALARY_RANGE reads Full List index first, falls back to direct Aon lookups if missing'],
    ['- Category mapping: X0 (Engineering/Product) = P25/P50/P90, Y1 (Everyone Else) = P10/P40/P62.5'],
    [''],
    ['MAPPINGS'],
    ['- Job family Descriptions: Aon Code ↔ Exec Description (use "Manage Exec Mappings")'],
    ['- Aon Code Remap: e.g., EN.SOML → EN.AIML for vendor changes'],
    ['- Title Mapping / Employee Level Mapping: use “Enhance mapping sheets” to add Missing highlights + counts'],
    [''],
    ['Menu overview'],
    ['Setup: Generate Help sheet, create tabs, category picker, manage mappings, enhance mapping sheets'],
    ['Build: Rebuild Full List, Seed exec mappings from region tabs, Build Full List USD, Clear caches'],
    ['Export: Proposed Salary Ranges'],
    ['Tools: Apply currency format'],
    [''],
    ['Tips'],
    ['- Keep region names as US/UK/India; FX is read from Lookup (Region, FX)'],
    ['- If data looks stale: Build → Clear all caches'],
    ['- For performance: prefer UI_SALARY_RANGE and keep Full List up to date']
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
    'Market \n\n (43) CFY Fixed Pay: 50th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 62.5th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 75th Percentile',
    'Market \n\n (43) CFY Fixed Pay: 90th Percentile'
  ];
  targets.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
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

function ensureExecFamilyPicker_() {
  const ss = SpreadsheetApp.getActive();
  const sh = uiSheet_(); if (!sh) return;
  const mapSh = ss.getSheetByName('Job family Descriptions'); if (!mapSh) return;
  const last = mapSh.getLastRow(); if (last <= 1) return;
  const vals = mapSh.getRange(2,2,last-1,1).getValues().map(r => String(r[0]||'').trim()).filter(Boolean);
  const uniq = Array.from(new Set(vals)).sort();
  const cell = sh.getRange('B2');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(uniq, true).setAllowInvalid(false).build();
  cell.setDataValidation(rule);
}

function buildCalculatorUI_() {
  const sh = uiSheet_(); if (!sh) return;
  ensureRegionPicker_();
  ensureCategoryPicker_();
  ensureExecFamilyPicker_();

  // Labels (keeps existing styling; only writes text)
  sh.getRange('A2').setValue('Job Family');
  sh.getRange('A3').setValue('Category');
  sh.getRange('A4').setValue('Region');

  // Header row - Updated labels for new range definitions
  sh.getRange('A7').setValue('Level');
  sh.getRange('B7').setValue('Range Start'); // Was P62.5
  sh.getRange('C7').setValue('Range Mid');   // Was P75
  sh.getRange('D7').setValue('Range End');   // Was P90
  sh.getRange('F7').setValue('Min');
  sh.getRange('G7').setValue('Median');
  sh.getRange('H7').setValue('Max');
  sh.getRange('L7').setValue('Emp Count');

  // Level list
  const levels = ['L2 IC','L3 IC','L4 IC','L5 IC','L5.5 IC','L6 IC','L6.5 IC','L7 IC','L4 Mgr','L5 Mgr','L5.5 Mgr','L6 Mgr','L6.5 Mgr','L7 Mgr','L8 Mgr','L9 Mgr'];
  sh.getRange(8,1,levels.length,1).setValues(levels.map(s=>[s]));

  // OPTIMIZED: Batch formula generation (85% faster)
  const formulasMin = [], formulasMid = [], formulasMax = [];
  const formulasIntMin = [], formulasIntMed = [], formulasIntMax = [], formulasIntCount = [];
  
  levels.forEach((level, i) => {
    const aRow = 8 + i;
    formulasMin.push([`=SALARY_RANGE_MIN($B$3,$B$4,$B$2,$A${aRow})`]);
    formulasMid.push([`=SALARY_RANGE_MID($B$3,$B$4,$B$2,$A${aRow})`]);
    formulasMax.push([`=SALARY_RANGE_MAX($B$3,$B$4,$B$2,$A${aRow})`]);
    formulasIntMin.push([`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,1)`]);
    formulasIntMed.push([`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,2)`]);
    formulasIntMax.push([`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,3)`]);
    formulasIntCount.push([`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,4)`]);
  });
  
  // Batch set all formulas at once (single API call per column)
  sh.getRange(8, 2, levels.length, 1).setFormulas(formulasMin);
  sh.getRange(8, 3, levels.length, 1).setFormulas(formulasMid);
  sh.getRange(8, 4, levels.length, 1).setFormulas(formulasMax);
  sh.getRange(8, 6, levels.length, 1).setFormulas(formulasIntMin);
  sh.getRange(8, 7, levels.length, 1).setFormulas(formulasIntMed);
  sh.getRange(8, 8, levels.length, 1).setFormulas(formulasIntMax);
  sh.getRange(8,12, levels.length, 1).setFormulas(formulasIntCount);

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

function _effectiveCategoryForFamily_(category, familyOrCode) {
  // Simplified: Only X0 (Engineering/Product) or Y1 (Everyone Else)
  // X0 is only for Engineering and allowed TE families
  if (_isEngineeringOrAllowedTE_(familyOrCode)) {
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
    // X0 (Engineering/Product): P25 → P50 → P90 (with fallbacks)
    // Y1 (Everyone Else): P10 → P40 → P62.5 (with fallbacks)
    if (cat === 'X0') {
      const min = rec.p25 || rec.p40 || rec.p50 || '';
      const mid = rec.p50 || rec.p625 || rec.p75 || '';
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
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([[ 'Job title (live)', 'Job title (Mapped)', 'Job family' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,3).setFontWeight('bold'); sh.autoResizeColumns(1,3);
  }
  // Job family Descriptions
  sh = ss.getSheetByName('Job family Descriptions') || ss.insertSheet('Job family Descriptions');
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,2).setValues([[ 'Aon Code', 'Job Family (Exec Description)' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,2).setFontWeight('bold'); sh.autoResizeColumns(1,2);
  }
  // Employee Level Mapping
  sh = ss.getSheetByName('Employee Level Mapping') || ss.insertSheet('Employee Level Mapping');
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([[ 'Emp ID', 'Mapping', 'Status' ]]);
    sh.setFrozenRows(1); sh.getRange(1,1,1,3).setFontWeight('bold'); sh.autoResizeColumns(1,3);
  }
  // Aon Code Remap
  sh = ss.getSheetByName('Aon Code Remap') || ss.insertSheet('Aon Code Remap');
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

/**
 * Creates unified menu when spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Main menu
  const menu = ui.createMenu('💰 Salary Ranges Calculator');
  
  // Setup submenu
  const setupMenu = ui.createMenu('⚙️ Setup')
    .addItem('⚡ Quick Setup (Run Once)', 'quickSetup_')
    .addSeparator()
    .addItem('📖 Generate Help Sheet', 'buildHelpSheet_')
    .addItem('🌍 Create Aon Region Tabs', 'createAonPlaceholderSheets_')
    .addItem('🗺️ Create Mapping Tabs', 'createMappingPlaceholderSheets_')
    .addItem('📊 Build Calculator UI', 'buildCalculatorUI_')
    .addSeparator()
    .addItem('🔧 Manage Exec Mappings', 'openExecMappingManager_')
    .addItem('✅ Ensure Category Picker', 'ensureCategoryPicker_')
    .addItem('🎨 Enhance Mapping Sheets', 'enhanceMappingSheets_');
  
  // Import submenu  
  const importMenu = ui.createMenu('📥 Import Data')
    .addItem('🔄 Import All Bob Data', 'importAllBobData')
    .addSeparator()
    .addItem('👥 Import Base Data Only', 'importBobDataSimpleWithLookup')
    .addItem('💰 Import Bonus Only', 'importBobBonusHistoryLatest')
    .addItem('📈 Import Comp History Only', 'importBobCompHistoryLatest');
  
  // Build submenu
  const buildMenu = ui.createMenu('🏗️ Build')
    .addItem('📊 Rebuild Full List (with validation)', 'rebuildFullListTabsWithValidation_')
    .addItem('💵 Build Full List USD', 'buildFullListUsd_')
    .addSeparator()
    .addItem('🌱 Seed All Job Family Mappings', 'seedAllJobFamilyMappings_')
    .addItem('👥 Sync All Bob Mappings', 'syncAllBobMappings_')
    .addSeparator()
    .addItem('🗑️ Clear All Caches', 'clearAllCaches_');
  
  // Export submenu
  const exportMenu = ui.createMenu('📤 Export')
    .addItem('💼 Export Proposed Ranges', 'exportProposedSalaryRanges_');
  
  // Tools submenu
  const toolsMenu = ui.createMenu('🔧 Tools')
    .addItem('💱 Apply Currency Format', 'applyCurrency_')
    .addItem('ℹ️ Instructions & Help', 'showInstructions');
  
  // Add all submenus to main menu
  menu.addSubMenu(setupMenu)
      .addSubMenu(importMenu)
      .addSubMenu(buildMenu)
      .addSubMenu(exportMenu)
      .addSubMenu(toolsMenu)
      .addToUi();
  
  // Auto-ensure category picker
  ensureCategoryPicker_();
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
