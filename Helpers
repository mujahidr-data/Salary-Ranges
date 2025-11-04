/********************************
 * SHARED HELPER FUNCTIONS
 * Consolidated helpers used across all import scripts
 * to reduce code duplication and improve maintainability
 ********************************/

/**
 * Find column index by trying multiple header aliases (case-insensitive)
 * @param {Array} headerRow - The header row array
 * @param {Array<string>} aliases - Array of possible column names
 * @returns {number} Column index or throws error if not found
 */
function findCol(headerRow, aliases) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
  const normalizedHeader = headerRow.map(norm);
  for (const alias of aliases) {
    const i = normalizedHeader.indexOf(norm(alias));
    if (i !== -1) return i;
  }
  throw new Error(
    `Could not find any of the columns [${aliases.join(", ")}]. Available headers: ${headerRow.join(" | ")}`
  );
}

/**
 * Find column index (optional) - returns -1 if not found instead of throwing
 * @param {Array} headerRow - The header row array
 * @param {Array<string>} aliases - Array of possible column names
 * @returns {number} Column index or -1 if not found
 */
function findColOptional(headerRow, aliases) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
  const normalizedHeader = headerRow.map(norm);
  for (const alias of aliases) {
    const i = normalizedHeader.indexOf(norm(alias));
    if (i !== -1) return i;
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
function toNumberSafe(val) {
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

