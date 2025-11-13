/********************************
 * OPTIMIZED Range Calculator
 * Key improvements:
 * - Consolidated sheet reads (single getDataRange per operation)
 * - Enhanced caching with longer TTL
 * - Reduced redundant header lookups
 * - Simplified region normalization
 ********************************/

/********************************
 * Region tabs for market data
 ********************************/
const REGION_TAB = {
  'India': 'Aon India - 2025',
  'US': 'Aon US Premium - 2025',
  'UK': 'Aon UK London - 2025',
};

/********************************
 * Configuration constants
 ********************************/
const CACHE_TTL = 600; // 10 minutes (increased from 5)
const UI_SHEET_NAME = 'Salary Ranges';
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

function findHeaderIndex_(headers, regex) {
  const re = new RegExp(regex, 'i');
  for (let i = 0; i < headers.length; i++) {
    if (re.test(String(headers[i] || ''))) return i;
  }
  return -1;
}

// Convert 1-based column index to A1 letter(s)
function _colToLetter_(col) {
  let c = Number(col);
  let out = '';
  while (c > 0) {
    const rem = (c - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    c = Math.floor((c - 1) / 26);
  }
  return out;
}

function toNumber_(v) {
  if (v == null || v === '') return NaN;
  const n = Number(String(v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? NaN : n;
}

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

function _aonValueCacheKey_(sheetName, fam, targetNum, prefLetter, ciqBaseLevel, headerRegex) {
  return `AON:${sheetName}|${fam}|${targetNum}|${prefLetter}|${ciqBaseLevel}|${headerRegex}`;
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
 * Robust header regexes
 ********************************/
const HDR_P40  = '(?:^\\s*Market\\s*\\(43\\)\\s*CFY\\s*Fixed\\s*Pay:\\s*40(?:th)?\\s*Percentile\\s*$|^\\s*40(?:th)?\\s*Percentile\\s*$|^\\s*P\\s*40\\s*$)';
const HDR_P50  = '(?:^\\s*Market\\s*\\(43\\)\\s*CFY\\s*Fixed\\s*Pay:\\s*50(?:th)?\\s*Percentile\\s*$|^\\s*50(?:th)?\\s*Percentile\\s*$|^\\s*P\\s*50\\s*$)';
const HDR_P625 = '(?:^\\s*Market\\s*\\(43\\)\\s*CFY\\s*Fixed\\s*Pay:\\s*62(?:[\\.,])?5(?:th)?\\s*Percentile\\s*$|^\\s*62(?:[\\.,])?5(?:th)?\\s*Percentile\\s*$|^\\s*P\\s*62(?:[\\.,])?5\\s*$)';
const HDR_P75  = '(?:^\\s*Market\\s*\\(43\\)\\s*CFY\\s*Fixed\\s*Pay:\\s*75(?:th)?\\s*Percentile\\s*$|^\\s*75(?:th)?\\s*Percentile\\s*$|^\\s*P\\s*75\\s*$)';
const HDR_P90  = '(?:^\\s*Market\\s*\\(43\\)\\s*CFY\\s*Fixed\\s*Pay:\\s*90(?:th)?\\s*Percentile\\s*$|^\\s*90(?:th)?\\s*Percentile\\s*$|^\\s*P\\s*90\\s*$)';

/********************************
 * Public custom functions (P50 anchor)
 ********************************/
function AON_P40(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P40);  }
function AON_P50(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P50);  }
function AON_P625(region, family, ciqLevel) { return _aonPick_(region, family, ciqLevel, HDR_P625); }
function AON_P75(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P75);  }
function AON_P90(region, family, ciqLevel)  { return _aonPick_(region, family, ciqLevel, HDR_P90);  }

/********************************
 * Category-based salary ranges (X0 / X1 / Y1)
 ********************************/
function _rangeByCategory_(category, region, family, ciqLevel) {
  const cat = String(category || '').trim().toUpperCase();
  if (!cat) return { min: '', mid: '', max: '' };

  if (cat === 'X0') {
    // X0: min=P62.5, mid=P75, max=P90
    const min = AON_P625(region, family, ciqLevel);
    const mid = AON_P75(region, family, ciqLevel);
    const max = AON_P90(region, family, ciqLevel);
    return { min, mid, max };
  }
  if (cat === 'X1') {
    // X1: min=P50, mid=P62.5, max=P75
    const min = AON_P50(region, family, ciqLevel);
    const mid = AON_P625(region, family, ciqLevel);
    const max = AON_P75(region, family, ciqLevel);
    return { min, mid, max };
  }
  if (cat === 'Y1') {
    // Y1: min=P40, mid=P50, max=P62.5
    const min = AON_P40(region, family, ciqLevel);
    const mid = AON_P50(region, family, ciqLevel);
    const max = AON_P625(region, family, ciqLevel);
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
    return ss.getSheetByName('Aon US Premium - 2025');
  }
  if (r === 'UK') {
    return ss.getSheetByName('Aon UK London - 2025');
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

  const cacheKey = `INT:${siteWanted}|${famCodeU}|${friendlyName}|${lvlU}`;
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
 * Menu + triggers
 ********************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Salary Ranges');

  // Setup
  const setup = ui.createMenu('Setup')
    .addItem('Generate Help sheet', 'buildHelpSheet_')
    .addSeparator()
    .addItem('Create Aon placeholder tabs', 'createAonPlaceholderSheets_')
    .addItem('Create mapping placeholder tabs', 'createMappingPlaceholderSheets_')
    .addItem('Ensure category picker', 'ensureCategoryPicker_')
    .addItem('Build Calculator UI', 'buildCalculatorUI_')
    .addItem('Manage Exec Mappings', 'openExecMappingManager_')
    .addItem('Enhance mapping sheets (format + counts)', 'enhanceMappingSheets_');

  // Build
  const build = ui.createMenu('Build')
    .addItem('Rebuild Full List tabs', 'rebuildFullListTabs_')
    .addItem('Seed exec mappings from region tabs', 'seedExecMappingsFromAon_')
    .addItem('Fill Job Family in region tabs', 'fillRegionFamilies_')
    .addItem('Sync Employee Level Mapping from Bob', 'syncEmployeeLevelMappingFromBob_')
    .addItem('Sync Title Mapping from Bob', 'syncTitleMappingFromBob_')
    .addItem('Build Full List USD', 'buildFullListUsd_')
    .addItem('Clear all caches', 'clearAllCaches_');

  // Imports
  const importsM = ui.createMenu('Imports')
    .addItem('Import Bob Base Data', 'importBobDataSimpleWithLookup')
    .addItem('Import Bob Bonus (latest)', 'importBobBonusHistoryLatest')
    .addItem('Import Bob Comp (latest)', 'importBobCompHistoryLatest');

  // Export
  const exportM = ui.createMenu('Export')
    .addItem('Export Proposed Salary Ranges', 'exportProposedSalaryRanges_');

  // Tools & Help
  const tools = ui.createMenu('Tools')
    .addItem('Apply currency format', 'applyCurrency_');

  menu.addSubMenu(setup)
      .addSubMenu(importsM)
      .addSubMenu(build)
      .addSubMenu(exportM)
      .addSubMenu(tools)
      .addToUi();
  // Apply formatting only when invoked from menu to reduce overhead
  ensureCategoryPicker_();
}

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

  const values = sh.getDataRange().getValues();
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
  const mVals = mapSh.getDataRange().getValues();
  const mHead = mVals[0].map(h => String(h || '').replace(/\s+/g,' ').trim());
  const colEmp = mHead.findIndex(h => /^Emp\s*ID/i.test(h));
  let colMap = mHead.findIndex(h => /Is\s*Mapped\?/i.test(h));
  if (colMap < 0) colMap = mHead.findIndex(h => /^Mapping$/i.test(h));
  if (colEmp < 0 || colMap < 0) return out;

  const bVals = baseSh.getDataRange().getValues();
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
        const colP40  = headers.indexOf('Market (43) CFY Fixed Pay: 40th Percentile') >= 0 ? headers.indexOf('Market (43) CFY Fixed Pay: 40th Percentile') : findHeaderIndex_(headers, '40(?:th)?\\s*Percentile|\\bP\\s*40\\b');
        const colP50  = headers.indexOf('Market (43) CFY Fixed Pay: 50th Percentile') >= 0 ? headers.indexOf('Market (43) CFY Fixed Pay: 50th Percentile') : findHeaderIndex_(headers, '50(?:th)?\\s*Percentile|\\bP\\s*50\\b');
        const colP625 = headers.indexOf('Market (43) CFY Fixed Pay: 62.5th Percentile') >= 0 ? headers.indexOf('Market (43) CFY Fixed Pay: 62.5th Percentile') : findHeaderIndex_(headers, '62[\\.,]?5(?:th)?\\s*Percentile|\\bP\\s*62[\\.,]?5\\b');
        const colP75  = headers.indexOf('Market (43) CFY Fixed Pay: 75th Percentile') >= 0 ? headers.indexOf('Market (43) CFY Fixed Pay: 75th Percentile') : findHeaderIndex_(headers, '75(?:th)?\\s*Percentile|\\bP\\s*75\\b');
        const colP90  = headers.indexOf('Market (43) CFY Fixed Pay: 90th Percentile') >= 0 ? headers.indexOf('Market (43) CFY Fixed Pay: 90th Percentile') : findHeaderIndex_(headers, '90(?:th)?\\s*Percentile|\\bP\\s*90\\b');
        if (colJobCode >= 0 && colJobFam >= 0 && colP50 >= 0 && colP625 >= 0 && colP75 >= 0) {
          for (let r=1; r<values.length; r++) {
            const row = values[r]; const jc = String(row[colJobCode] || '').trim(); if (!jc) continue;
            const i = jc.lastIndexOf('.'); const base = i>=0 ? jc.slice(0,i) : jc; const suf = (i>=0 ? jc.slice(i+1) : jc).toUpperCase().replace(/[^A-Z0-9]/g,'');
            const fam = String(row[colJobFam] || '').trim(); if (base && fam && !famByBase.has(base)) famByBase.set(base, fam);
            const p40 = colP40 >= 0 ? toNumber_(row[colP40]) : NaN; const p50 = toNumber_(row[colP50]); const p62 = toNumber_(row[colP625]); const p75 = toNumber_(row[colP75]); const p90 = colP90 >= 0 ? toNumber_(row[colP90]) : NaN;
            byKey.set(`${base}|${suf}`, { p40, p50, p62, p75, p90 });
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
      lookupRows.forEach(L => { if (L.half) return; const rec = idx.get(`${base}|${L.aon}`); whole.set(`${L.role}|${Math.floor(L.base)}`, rec || { p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN }); });
      lookupRows.forEach(L => {
        let p40, p50, p62, p75, p90;
        if (L.half) { const k1 = `${L.role}|${Math.floor(L.base)}`; const k2 = `${L.role}|${Math.floor(L.base)+1}`; const v1 = whole.get(k1) || {p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN}; const v2 = whole.get(k2) || {p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN}; p40 = _avg2_(v1.p40, v2.p40); p50 = _avg2_(v1.p50, v2.p50); p62 = _avg2_(v1.p62, v2.p62); p75 = _avg2_(v1.p75, v2.p75); p90 = _avg2_(v1.p90, v2.p90); }
        else { const rec = idx.get(`${base}|${L.aon}`) || { p40:NaN,p50:NaN,p62:NaN,p75:NaN,p90:NaN }; p40 = rec.p40; p50 = rec.p50; p62 = rec.p62; p75 = rec.p75; p90 = rec.p90; }
        const ist = internalIdx.get(`${site}|${String(execFam).toUpperCase()}|${L.ciq}`) || internalIdx.get(`${site}|${base}|${L.ciq}`) || null; const key = `${execFam}${L.ciq}${region}`;
        const uniqueKey = `${site}|${region}|${baseOut}|${String(execFam)}|${L.ciq}`;
        if (!emitted.has(uniqueKey)) {
          emitted.add(uniqueKey);
          rows.push([site, region, baseOut, execFam, rawFam, L.ciq, L.aon, _round100_(p40), _round100_(p50), _round100_(p62), _round100_(p75), _round100_(p90), ist ? _round0_(ist.min) : '', ist ? _round0_(ist.med) : '', ist ? _round0_(ist.max) : '', ist ? ist.n : '', '', key]);
        }
      });
    });
  });

  const fl = ss.getSheetByName('Full List') || ss.insertSheet('Full List');
  const fullHeader = ['Site','Region','Aon Code','Job Family (Exec Description)','Job Family (Raw)','CIQ Level','Aon Level','P40','P50','P62.5','P75','P90','Internal Min','Internal Median','Internal Max','Employees','', 'Key'];
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
    [cP40,cP50,cP625,cP75,cP90,cIMin,cIMed,cIMax].forEach(mul);
    // Round market percentiles to nearest hundred after FX conversion
    const r100 = (i) => { if (i >= 0) { const n = toNumber_(row[i]); if (!isNaN(n)) row[i] = _round100_(n); } };
    [cP40,cP50,cP625,cP75,cP90].forEach(r100);
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
  const resp = ui.prompt('Export Proposed Salary Ranges', 'Enter category (X0, X1, Y1). Default X0:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const category = String(resp.getResponseText() || 'X0').trim().toUpperCase();
  if (!/^(X0|X1|Y1)$/.test(category)) { ui.alert('Invalid category. Use X0, X1, or Y1.'); return; }
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
    ['Pre-steps (once per workbook)'],
    ['1) Setup → Generate Help sheet (this page)'],
    ['2) Setup → Create Aon placeholder tabs (creates empty US/UK/India tabs if needed)'],
    ['3) Paste Aon data into region tabs (US, UK, India) with headers: Job Code, Job Family, 40th, 50th, 62.5th, 75th, 90th'],
    ['4) Setup → Create mapping placeholder tabs (creates Title Mapping, Job family Descriptions, Employee Level Mapping, Aon Code Remap)'],
    ['5) Build → Seed exec mappings from region tabs (populates Job family Descriptions from region data)'],
    ['6) Setup → Manage Exec Mappings (review/adjust code ↔ exec description)'],
    ['7) Setup → Ensure category picker (adds X0/X1/Y1 dropdown in B3)'],
    [''],
    ['Regular workflow'],
    ['A) Build → Rebuild Full List tabs (generates Full List, Coverage Summary, Employees (Mapped))'],
    ['B) Build → Build Full List USD (optional FX-applied view)'],
    ['C) Use calculators with UI_SALARY_RANGE or SALARY_RANGE'],
    ['D) Export → Export Proposed Salary Ranges (optional)'],
    [''],
    ['Imports (Bob)'],
    ['- App_Imports: importBobDataSimpleWithLookup (auto-maps Job Family Name, Mapped Family)'],
    ['- importBobBonusHistoryLatest / importBobCompHistoryLatest'],
    ['Script properties required: BOB_ID, BOB_KEY'],
    [''],
    ['How calculations work'],
    ['- Rebuild Full List creates the “Full List” sheet with P40/P50/P62.5/P75/P90 + Internal Min/Median/Max + Employees + Key (R)'],
    ['- A cache index of Full List is built on demand (10-min TTL)'],
    ['- SALARY_RANGE(category, region, familyOrCode, ciqLevel) first reads the Full List index; if missing, it falls back to direct Aon tab lookups'],
    ['- Category mapping: X0 = P62.5/P75/P90, X1 = P50/P62.5/P75, Y1 = P40/P50/P62.5'],
    ['- UI_SALARY_RANGE* functions read the picker at Salary Ranges!B3'],
    [''],
    ['Mappings'],
    ['- Job family Descriptions: Aon Code ↔ Exec Description. Use “Manage Exec Mappings” to add/update/delete'],
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
    'Aon US Premium - 2025',
    'Aon UK London - 2025'
  ];
  const headers = [
    'Job Code',
    'Job Family',
    'Market (43) CFY Fixed Pay: 40th Percentile',
    'Market (43) CFY Fixed Pay: 50th Percentile',
    'Market (43) CFY Fixed Pay: 62.5th Percentile',
    'Market (43) CFY Fixed Pay: 75th Percentile',
    'Market (43) CFY Fixed Pay: 90th Percentile'
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
    .requireValueInList(['X0','X1','Y1'], true)
    .setAllowInvalid(false)
    .build();
  const currentRule = cell.getDataValidation();
  if (!currentRule || String(currentRule) !== String(rule)) cell.setDataValidation(rule);
  const v = String(cell.getDisplayValue() || '').trim();
  if (!v) cell.setValue('X0');
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

  // Header row
  sh.getRange('A7').setValue('Level');
  sh.getRange('B7').setValue('P62.5');
  sh.getRange('C7').setValue('P75');
  sh.getRange('D7').setValue('P90');
  sh.getRange('F7').setValue('Min');
  sh.getRange('G7').setValue('Median');
  sh.getRange('H7').setValue('Max');
  sh.getRange('L7').setValue('Emp Count');

  // Level list
  const levels = ['L2 IC','L3 IC','L4 IC','L5 IC','L5.5 IC','L6 IC','L6.5 IC','L7 IC','L4 Mgr','L5 Mgr','L5.5 Mgr','L6 Mgr','L6.5 Mgr','L7 Mgr','L8 Mgr','L9 Mgr'];
  sh.getRange(8,1,levels.length,1).setValues(levels.map(s=>[s]));

  // Market range formulas
  for (let r=0; r<levels.length; r++) {
    const aRow = 8 + r;
    sh.getRange(aRow, 2).setFormula(`=SALARY_RANGE_MIN($B$3,$B$4,$B$2,$A${aRow})`);
    sh.getRange(aRow, 3).setFormula(`=SALARY_RANGE_MID($B$3,$B$4,$B$2,$A${aRow})`);
    sh.getRange(aRow, 4).setFormula(`=SALARY_RANGE_MAX($B$3,$B$4,$B$2,$A${aRow})`);
    // Internal stats (min/median/max/count)
    sh.getRange(aRow, 6).setFormula(`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,1)`);
    sh.getRange(aRow, 7).setFormula(`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,2)`);
    sh.getRange(aRow, 8).setFormula(`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,3)`);
    sh.getRange(aRow,12).setFormula(`=INDEX(INTERNAL_STATS($B$4,$B$2,$A${aRow}),1,4)`);
  }

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
  const cat = String(category || '').trim().toUpperCase();
  if (cat === 'X0' || cat === 'X1') {
    return _isEngineeringOrAllowedTE_(familyOrCode) ? cat : 'Y1';
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
    if (cat === 'X0') return { min: rec.p625, mid: rec.p75,  max: rec.p90 };
    if (cat === 'X1') return { min: rec.p50,  mid: rec.p625, max: rec.p75 };
    if (cat === 'Y1') return { min: rec.p40,  mid: rec.p50,  max: rec.p625 };
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
    ss.getSheetByName('US') || ss.getSheetByName('Aon US Premium - 2025'),
    ss.getSheetByName('UK') || ss.getSheetByName('Aon UK London - 2025'),
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
    ss.getSheetByName('US') || ss.getSheetByName('Aon US Premium - 2025'),
    ss.getSheetByName('UK') || ss.getSheetByName('Aon UK London - 2025'),
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
  const vals = base.getDataRange().getValues();
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
  const vals = base.getDataRange().getValues();
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
