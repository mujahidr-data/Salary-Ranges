# HOTFIX v4.6.1 - Lookup Sheet Section Detection Fix

## 🐛 **Bug Discovered**

The Full List was populating with incorrect data:
- ❌ Job Family showing "Avg of P5 and P6" instead of actual job family names
- ❌ All percentile values showing "0"
- ❌ All CR columns showing "0"
- ❌ Key column malformed

## 🔍 **Root Cause**

Three functions were reading from the **wrong sections** of the Lookup sheet:

### **Problem: Indiscriminate Reading**

The Lookup sheet has 3 distinct sections:
1. **CIQ Level → Aon Level** (rows with "L5.5 IC" → "Avg of P5 and P6")
2. **Region/Site → FX Rate**
3. **Aon Code → Job Family + Category** (rows with "EN.SODE" → "Engineering - Software Development")

**Before the fix**, these functions were reading **ALL rows** that met loose criteria:

#### **`_getExecDescMap_()`** (Job Family mapping)
```javascript
// ❌ BEFORE: Read ANY row where column 1 contains a dot
if (col1 && col1.includes('.') && col2) {
  map.set(col1, col2);
}
```

**Result**: Picked up level mappings like:
- `"L5.5 IC"` → `"Avg of P5 and P6"` ❌
- `"L6.5 IC"` → `"Avg of P6 and E1"` ❌

Instead of actual Aon codes like:
- `"EN.SODE"` → `"Engineering - Software Development"` ✅

#### **`_getCategoryMap_()`** (Category X0/Y1 mapping)
```javascript
// ❌ BEFORE: Read ANY row where column 1 contains a dot
if (col1 && col1.includes('.') && (col3 === 'X0' || col3 === 'Y1')) {
  map.set(col1, col3);
}
```

**Result**: Tried to map level codes as if they were Aon codes.

#### **`_getFxMap_()`** (FX Rate mapping)
```javascript
// ❌ BEFORE: Read from first row with "Region" and "FX" headers
const head = vals[0].map(h => String(h || '').trim());
let cRegion = head.findIndex(h => /^Region$/i.test(h));
const cFx = head.findIndex(h => /^FX$/i.test(h));
```

**Result**: Could potentially read wrong columns if multiple sections had similar headers.

---

## ✅ **Fix Applied**

### **1. `_getExecDescMap_()` - Proper Section Detection**

```javascript
// ✅ AFTER: Only read from Aon Code section
let inAonCodeSection = false;

for (let r = 0; r < vals.length; r++) {
  const col1 = String(row[0] || '').trim();
  const col2 = String(row[1] || '').trim();
  
  // Detect Aon Code section header
  if (col1 === 'Aon Code' && /Job.*Family.*Exec/i.test(col2)) {
    inAonCodeSection = true;
    continue;
  }
  
  // Stop at next section
  if (inAonCodeSection && (col1 === 'CIQ Level' || col1 === 'Region')) {
    inAonCodeSection = false;
    continue;
  }
  
  // Only read Aon Code section + validate format
  if (inAonCodeSection && col1 && col2) {
    // Validate: XX.YYYY format (not L5.5 IC)
    if (/^[A-Z]{2}\.[A-Z0-9]{4}$/i.test(col1)) {
      map.set(col1, col2);
    }
  }
}
```

**Now only reads**:
- `"EN.SODE"` → `"Engineering - Software Development"` ✅
- `"EN.PGPG"` → `"Engineering - Product Management/ TPM"` ✅
- `"SA.CRCS"` → `"Sales - Customer Success"` ✅

**Ignores**:
- `"L5.5 IC"` → `"Avg of P5 and P6"` (Level mapping, not Aon code)

---

### **2. `_getCategoryMap_()` - Proper Section Detection**

```javascript
// ✅ AFTER: Only read from Aon Code section (has Category in column 3)
let inAonCodeSection = false;

for (let r = 0; r < vals.length; r++) {
  const col1 = String(row[0] || '').trim();
  const col2 = String(row[1] || '').trim();
  const col3 = String(row[2] || '').trim().toUpperCase();
  
  // Detect Aon Code section header
  if (col1 === 'Aon Code' && /Job.*Family.*Exec/i.test(col2) && col3 === 'Category') {
    inAonCodeSection = true;
    continue;
  }
  
  // Stop at next section
  if (inAonCodeSection && (col1 === 'CIQ Level' || col1 === 'Region')) {
    break;
  }
  
  // Only read Aon Code section + validate format + validate category
  if (inAonCodeSection && col1 && (col3 === 'X0' || col3 === 'Y1')) {
    if (/^[A-Z]{2}\.[A-Z0-9]{4}$/i.test(col1)) {
      map.set(col1, col3);
    }
  }
}
```

---

### **3. `_getFxMap_()` - Proper Section Detection**

```javascript
// ✅ AFTER: Only read from FX section
let inFxSection = false;

for (let r = 0; r < vals.length; r++) {
  const col1 = String(row[0] || '').trim();
  const col2 = String(row[1] || '').trim();
  const col3 = String(row[2] || '').trim();
  
  // Detect FX section header (Region, Site, FX Rate)
  if (col1 === 'Region' && col2 === 'Site' && /FX.*Rate/i.test(col3)) {
    inFxSection = true;
    continue;
  }
  
  // Stop at next section
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
```

---

### **4. `_preloadAonData_()` - Enhanced Logging**

Added comprehensive debug logging to help troubleshoot Aon data loading:

```javascript
Logger.log(`Region ${region}: Headers = ${headers.slice(0, 10).join(', ')}`);
Logger.log(`Region ${region}: JobCode col=${colJobCode}, P10=${colP10}, P25=${colP25}, P625=${colP625}`);
Logger.log(`Sample: ${jobCode} → ${family}, ${ciqLevel}, P25=${row[colP25]}, P625=${row[colP625]}`);
Logger.log(`Region ${region}: Indexed ${rowCount} job codes`);
Logger.log(`Pre-loaded ${aonCache.size} total Aon data combinations`);
```

**This will help identify**:
- If Aon sheets are being found
- If columns are being detected correctly
- If data is being parsed correctly
- Sample values for verification

---

## 📊 **Expected Results After Fix**

### **Before (v4.6.0)**:
| Region | Job Family | Category | P10 | P25 | P625 | P90 |
|--------|------------|----------|-----|-----|------|-----|
| India | Avg of P5 and P6 | Y1 | 0 | 0 | 0 | 0 |
| India | Avg of M4 and M5 | Y1 | 0 | 0 | 0 | 0 |

### **After (v4.6.1)**:
| Region | Job Family | Category | P10 | P25 | P625 | P90 |
|--------|------------|----------|-----|-----|------|-----|
| India | Engineering - Software Development | X0 | 500000 | 750000 | 1200000 | 1800000 |
| India | Sales - Customer Success | Y1 | 400000 | 600000 | 900000 | 1200000 |

---

## 🚀 **Deployment Steps**

1. **Commit and push** to GitHub
2. **Deploy to Apps Script** via `clasp push`
3. **Test in Google Sheets**:
   - Open the sheet and refresh (F5)
   - Go to **Bob Data** → **Build Market Data**
   - Check Full List has correct data

---

## 🧪 **Testing Checklist**

After deploying, verify:
- [ ] Full List shows actual job family names (not "Avg of P5 and P6")
- [ ] Percentile columns (P10, P25, P40, P50, P62.5, P75, P90) have numeric values
- [ ] CR columns have calculated values (or blank if no employees)
- [ ] Key column format: `Job FamilyLevelRegion` (e.g., "Engineering - Software DevelopmentL5 ICIndia")
- [ ] Check Logger output for debug messages confirming data loading

---

## 📝 **Technical Details**

### **Validation Regex**

```javascript
/^[A-Z]{2}\.[A-Z0-9]{4}$/i
```

**Matches**:
- `EN.SODE` ✅
- `SA.CRCS` ✅
- `FI.ACCO` ✅
- `EN.AIML` ✅

**Rejects**:
- `L5.5 IC` ❌ (level code, not Aon code)
- `L6.5 IC` ❌
- `P5` ❌
- `M4` ❌

### **Section Detection Pattern**

Each function now:
1. Scans for the **exact header row** of its section
2. Sets a flag `inSection = true`
3. Reads data **only while flag is true**
4. Stops when encountering **next section's header**
5. Additionally validates data format for extra safety

---

## 🔧 **Files Modified**

- `SalaryRangesCalculator.gs`:
  - `_getExecDescMap_()` - Fixed to read only Aon Code section
  - `_getCategoryMap_()` - Fixed to read only Aon Code section
  - `_getFxMap_()` - Fixed to read only FX section
  - `_preloadAonData_()` - Added debug logging

---

**Version**: 4.6.1  
**Priority**: CRITICAL (data integrity issue)  
**Impact**: Fixes incorrect Full List generation  
**Backward Compatible**: Yes

