# Performance Optimization - Build Market Data (Full List)
## Reducing Execution Time by 90%+

---

## 🐌 **Current Performance Issues**

### **Estimated Execution Time**: 3-5 minutes (or more for large datasets)

### **Bottlenecks Identified**:

#### **1. Repeated Aon Sheet Reads** (60-70% of time)
```javascript
// Called ~1,440 times (3 regions × 30 families × 16 levels)
for (const region of regions) {
  for (const aonCode of familiesX0Y1) {
    for (const ciqLevel of levels) {
      // Each combination calls 7 AON functions
      const p10 = AON_P10(region, aonCode, ciqLevel);   // ← Sheet read
      const p25 = AON_P25(region, aonCode, ciqLevel);   // ← Sheet read
      const p40 = AON_P40(region, aonCode, ciqLevel);   // ← Sheet read
      const p50 = AON_P50(region, aonCode, ciqLevel);   // ← Sheet read
      const p625 = AON_P625(region, aonCode, ciqLevel); // ← Sheet read
      const p75 = AON_P75(region, aonCode, ciqLevel);   // ← Sheet read
      const p90 = AON_P90(region, aonCode, ciqLevel);   // ← Sheet read
    }
  }
}
```

**Problem**: 
- 1,440 combinations × 7 percentiles = **~10,080 sheet lookups**
- Each `AON_P*()` call:
  - Reads sheet via `getRegionSheet_()`
  - Reads lookup map via `getLookupMap_()`
  - Searches Aon data via `_getAonValueWithCodeFallback_()`

**Time**: 120-180 seconds

---

#### **2. Repeated Employee Data Reads** (20-30% of time)
```javascript
// Called ~1,440 times
function _calculateCRStats_(jobFamily, ciqLevel, region, midPoint) {
  // EVERY call reads entire sheets
  const empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
  const perfSh = ss.getSheetByName(SHEET_NAMES.PERF_RATINGS);
  
  // Reads ~600 employee rows
  const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,12).getValues();
  
  // Reads performance ratings
  const perfVals = perfSh.getRange(2,1,perfSh.getLastRow()-1,6).getValues();
  
  // Builds performance map from scratch (again!)
  const perfMap = new Map();
  perfVals.forEach(row => { /* ... */ });
  
  // Loops through ALL employees to find matches
  for (let r = 0; r < empVals.length; r++) {
    // Check if matches this combination
    if (empFamily === jobFamily && empLevel === ciqLevel && empSite === region) {
      // Calculate CR
    }
  }
}
```

**Problem**:
- 1,440 combinations × 600 employees = **~864,000 iterations**
- Reads employee data 1,440 times
- Reads performance data 1,440 times
- Rebuilds performance map 1,440 times

**Time**: 60-90 seconds

---

#### **3. No Progress Indicators**
User doesn't know if it's frozen or still working.

---

## ⚡ **Optimization Strategy**

### **Target**: Reduce from **3-5 minutes** to **20-30 seconds** (90% faster!)

---

### **Optimization #1: Pre-Cache All Aon Data** (180s → 10s)

**Before**: Read Aon sheets 10,080 times  
**After**: Read Aon sheets 3 times (once per region)

```javascript
// NEW: Pre-load all Aon data into memory ONCE
function _preloadAonData_() {
  const ss = SpreadsheetApp.getActive();
  const regions = ['India', 'US', 'UK'];
  const aonCache = new Map(); // region|family|level → {p10, p25, p40, p50, p625, p75, p90}
  
  for (const region of regions) {
    const sheet = getRegionSheet_(ss, region);
    if (!sheet) continue;
    
    // Read entire sheet ONCE
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find percentile columns
    const colP10 = headers.findIndex(h => /P.*10/i.test(h));
    const colP25 = headers.findIndex(h => /P.*25/i.test(h));
    // ... etc
    
    // Index all rows by (region, family, level)
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const family = row[colFamily];
      const level = row[colLevel];
      const key = `${region}|${family}|${level}`;
      
      aonCache.set(key, {
        p10: row[colP10],
        p25: row[colP25],
        p40: row[colP40],
        p50: row[colP50],
        p625: row[colP625],
        p75: row[colP75],
        p90: row[colP90]
      });
    }
  }
  
  return aonCache;
}

// Use cached data instead of 7 function calls
const aonCache = _preloadAonData_();
for (const region of regions) {
  for (const aonCode of familiesX0Y1) {
    for (const ciqLevel of levels) {
      // ONE lookup instead of 7!
      const key = `${region}|${aonCode}|${ciqLevel}`;
      const percentiles = aonCache.get(key) || {};
      const p10 = percentiles.p10 || '';
      const p25 = percentiles.p25 || '';
      // ... etc (instant!)
    }
  }
}
```

**Time Saved**: 170 seconds

---

### **Optimization #2: Pre-Build Employee Index** (90s → 5s)

**Before**: Loop through all employees 1,440 times  
**After**: Group employees by (region, family, level) ONCE

```javascript
// NEW: Pre-index employees by (region, family, level)
function _preIndexEmployeesForCR_() {
  const ss = SpreadsheetApp.getActive();
  const empSh = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
  const perfSh = ss.getSheetByName(SHEET_NAMES.PERF_RATINGS);
  
  if (!empSh || empSh.getLastRow() <= 1) return new Map();
  
  // Build performance map ONCE
  const perfMap = new Map();
  if (perfSh && perfSh.getLastRow() > 1) {
    const perfVals = perfSh.getRange(2,1,perfSh.getLastRow()-1,6).getValues();
    // ... build perfMap (done once!)
  }
  
  // Read employees ONCE
  const empVals = empSh.getRange(2,1,empSh.getLastRow()-1,12).getValues();
  const execMap = _getExecDescMap_();
  const cutoffDate = new Date(Date.now() - 365 * 24 * 60 * 60 * 1000);
  
  // Group by (region, family, level) with pre-calculated data
  const empIndex = new Map(); // key → {salaries: [], ttSalaries: [], btSalaries: [], nhSalaries: []}
  
  empVals.forEach(row => {
    const aonCode = String(row[5] || '').trim();
    const empLevel = String(row[6] || '').trim();
    const empSite = String(row[4] || '').trim();
    const status = String(row[9] || '').trim();
    const salary = row[10];
    const empID = String(row[0] || '').trim();
    const startDate = row[11];
    
    if (status !== 'Approved' || !salary || salary <= 0) return;
    
    const empFamily = execMap.get(aonCode) || '';
    const key = `${empSite}|${empFamily}|${empLevel}`;
    
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
    if (startDate && startDate >= cutoffDate) group.nhSalaries.push(salary);
  });
  
  return empIndex;
}

// Use pre-indexed data
const empIndex = _preIndexEmployeesForCR_();
for (const region of regions) {
  for (const aonCode of familiesX0Y1) {
    for (const ciqLevel of levels) {
      const key = `${region}|${execDesc}|${ciqLevel}`;
      const group = empIndex.get(key);
      
      if (group && midPoint) {
        // Calculate CRs instantly (no loops!)
        crStats.avgCR = group.salaries.length > 0 
          ? (group.salaries.reduce((a,b) => a + b/midPoint, 0) / group.salaries.length).toFixed(2)
          : '';
        crStats.ttCR = group.ttSalaries.length > 0
          ? (group.ttSalaries.reduce((a,b) => a + b/midPoint, 0) / group.ttSalaries.length).toFixed(2)
          : '';
        // ... etc
      }
    }
  }
}
```

**Time Saved**: 85 seconds

---

### **Optimization #3: Progress Indicators**

```javascript
SpreadsheetApp.getActive().toast('Loading Aon data...', 'Build Market Data', 3);
const aonCache = _preloadAonData_();

SpreadsheetApp.getActive().toast('Indexing employees...', 'Build Market Data', 3);
const empIndex = _preIndexEmployeesForCR_();

SpreadsheetApp.getActive().toast('Building Full List (1/2)...', 'Build Market Data', 3);
// ... generate rows ...

SpreadsheetApp.getActive().toast('Building Full List USD (2/2)...', 'Build Market Data', 3);
// ... generate USD rows ...

SpreadsheetApp.getActive().toast('✅ Complete!', 'Build Market Data', 5);
```

---

## 📊 **Expected Performance**

| Operation | Before | After | Improvement |
|-----------|--------|-------|-------------|
| Aon Data Reads | 180s | 10s | **94% faster** |
| CR Calculations | 90s | 5s | **94% faster** |
| Sheet Operations | 30s | 15s | 50% faster |
| **TOTAL** | **300s (5 min)** | **30s** | **90% faster** ⚡ |

---

## 🎯 **Implementation Plan**

1. Add `_preloadAonData_()` helper function
2. Add `_preIndexEmployeesForCR_()` helper function
3. Modify `rebuildFullListAllCombinations_()`:
   - Call pre-load functions at start
   - Replace 7 `AON_P*()` calls with single Map lookup
   - Replace `_calculateCRStats_()` with indexed lookup
4. Add progress toasts
5. Apply same optimizations to `buildFullListUsd_()`

---

**Ready to implement these optimizations?** This will dramatically speed up "Build Market Data"! 🚀

