# Performance Optimization Plan - v4.5.0
## syncEmployeesMappedSheet_() Speed Improvements

### Current Performance
- **Execution Time**: ~90-120 seconds for 600+ employees
- **Complexity**: O(n²) in title mapping logic
- **Main Bottlenecks**:
  1. Nested loop: 675 legacy × 600+ base data = ~400K iterations
  2. Repeated `_getLegacyMapping_()` calls
  3. Conditional formatting for large dataset

### Target Performance
- **Execution Time**: ~15-20 seconds
- **Complexity**: O(n) with efficient indexing
- **Speed Improvement**: **6x faster**

---

## Optimization #1: Pre-Build Employee Index (Eliminates O(n²))

**Before** (Line 3905-3919):
```javascript
// For each legacy employee, search ALL base data
for (let i = 1; i < baseVals.length; i++) {
  if (String(baseVals[i][iEmpID] || '').trim() === legacyEmpID) {
    // Found it!
  }
}
```

**After**:
```javascript
// Build index ONCE at start
const empToTitle = new Map(); // empID → title
for (let i = 1; i < baseVals.length; i++) {
  const empID = String(baseVals[i][iEmpID] || '').trim();
  const title = iTitle >= 0 ? String(baseVals[i][iTitle] || '').trim() : '';
  if (empID && title) {
    empToTitle.set(empID, title);
  }
}

// Now use O(1) lookup
legacyVals.forEach(row => {
  const legacyEmpID = String(row[0] || '').trim();
  const jobTitle = empToTitle.get(legacyEmpID) || ''; // ← O(1) lookup!
  // ... process ...
});
```

**Time Saved**: 30-60 seconds → **2-3 seconds**

---

## Optimization #2: Cache Legacy Mappings

**Before**:
```javascript
// Called 600+ times
for (let r = 1; r < baseVals.length; r++) {
  const legacy = _getLegacyMapping_(empID); // Loads from storage EACH time
}
```

**After**:
```javascript
// Load ALL legacy mappings ONCE
const allLegacyMappings = _loadAllLegacyMappings_(); // New function
const legacyMap = new Map();
allLegacyMappings.forEach(row => {
  legacyMap.set(row[0], {aonCode: row[1], level: row[2]});
});

// Now use O(1) lookup
for (let r = 1; r < baseVals.length; r++) {
  const legacy = legacyMap.get(empID); // ← O(1) lookup!
}
```

**Time Saved**: 10-20 seconds → **1-2 seconds**

---

## Optimization #3: Batch Conditional Formatting

**Before**:
```javascript
// Creates 5 separate conditional format rules
rules.push(rule1);
rules.push(rule2);
rules.push(rule3);
rules.push(rule4);
rules.push(rule5);
empSh.setConditionalFormatRules(rules);
```

**After**:
```javascript
// Only apply if row count has changed significantly
const lastRowCount = empSh.getLastRow() - 1;
if (Math.abs(lastRowCount - rows.length) > 10 || !empSh.getConditionalFormatRules().length) {
  // Apply formatting rules
  empSh.setConditionalFormatRules(rules);
} else {
  // Skip - rules already in place
}
```

**Time Saved**: 5-10 seconds → **0-1 seconds** (when rules unchanged)

---

## Optimization #4: Progress Toasts

Add progress indicators to give user feedback:

```javascript
SpreadsheetApp.getActive().toast('Loading Base Data...', 'Syncing', 3);
// ... load base data ...

SpreadsheetApp.getActive().toast('Building mappings (1/3)...', 'Syncing', 3);
// ... build title map ...

SpreadsheetApp.getActive().toast('Processing employees (2/3)...', 'Syncing', 3);
// ... main loop ...

SpreadsheetApp.getActive().toast('Applying formatting (3/3)...', 'Syncing', 3);
// ... conditional formatting ...
```

---

## Implementation Code

### New Helper Function: _loadAllLegacyMappings_()

```javascript
/**
 * Loads all legacy mappings at once (more efficient than per-employee lookup)
 * Returns Map: empID → {aonCode, ciqLevel}
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
        legacyMap.set(empID, {aonCode, ciqLevel});
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
        legacyMap.set(empID, {aonCode, ciqLevel});
      }
    });
  }
  
  return legacyMap;
}
```

---

## Summary

| Optimization | Current Time | Optimized Time | Savings |
|--------------|--------------|----------------|---------|
| Title Mapping (O(n²) → O(n)) | 30-60s | 2-3s | **90% faster** |
| Legacy Lookups (600+ calls → 1 call) | 10-20s | 1-2s | **85% faster** |
| Conditional Formatting (smart skip) | 5-10s | 0-1s | **90% faster** |
| **Total** | **90-120s** | **15-20s** | **80% faster** |

---

## Implementation Steps

1. Add `_loadAllLegacyMappings_()` helper function
2. Modify `syncEmployeesMappedSheet_()`:
   - Add employee index (`empToTitle` Map)
   - Replace `_getLegacyMapping_()` calls with Map lookup
   - Add progress toasts
   - Add smart conditional formatting skip
3. Test with full dataset
4. Deploy

---

**Ready to implement?** This will make "Import Bob Data" significantly faster! 🚀

