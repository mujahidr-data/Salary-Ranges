# Internal Stats Diagnosis

## 🐛 **Current Issue**

Internal Min, Median, Max, and Emp Count columns showing 0 or blank.

## 🔍 **Key Differences: Old vs New**

### **Old INTERNAL_STATS Function**:
```javascript
// Matching logic:
const rowFamCodeU = String(row[colFam] || '').trim().toUpperCase();
const rowMapNameU = String(row[colMap] || '').trim().toUpperCase();

// Match if EITHER Job Family Name OR Mapped Family matches
if (!(rowFamCodeU === famCodeU || (friendlyName && rowMapNameU === friendlyName))) continue;
```

**Pros**:
- Simple, direct matching
- Flexible: matches by Job Family Name OR Mapped Family
- Called from calculator at runtime (always uses current data)

**Cons**:
- Slow (reads Base Data on every calculator change)
- No pre-indexing

---

### **New _buildInternalIndex_ Function**:
```javascript
// Key creation:
const famCode = String(row[colFam] || '').trim(); // e.g., "EN.SODE.P5"
const dot = famCode.lastIndexOf('.');
const base = dot >= 0 ? famCode.slice(0, dot) : famCode; // Extract "EN.SODE"

// Creates keys like: "USA|EN.SODE|L5 IC"
const key = `${normSite}|${base}|${ciq}`;
```

**Pros**:
- Fast (pre-indexed during Build Market Data)
- Efficient lookups

**Cons**:
- **CRITICAL**: Assumes Job Family Name has format "EN.SODE.P5"
- If format is different, extraction fails → no keys created!

---

## 🎯 **Potential Root Causes**

### **Cause #1: Job Family Name Format Mismatch**

**Expected format**: `EN.SODE.P5`  
**What extraction does**: Split by last `.` → get `EN.SODE`

**Problem**: If Base Data has Job Family Name as:
- Empty
- Just "EN.SODE" (no level suffix)
- Something else entirely

Then `base` will be empty or wrong, and no keys are created!

---

### **Cause #2: Key Format Mismatch in Lookup**

**Keys created**: `USA|EN.SODE|L5 IC`  
**Lookup uses**: `${intRegion}|${aonCode}|${ciqLevel}`  
Where `aonCode` comes from Lookup sheet's X0/Y1 families

**Problem**: If Lookup sheet has different Aon codes than Base Data Job Family Names, no matches!

---

### **Cause #3: Region Normalization**

**Index uses**: `USA` (normalized from "USA" or "US")  
**Lookup uses**: `intRegion = region === 'US' ? 'USA' : region`

**Fixed in v4.6.2**, but if Build Market Data wasn't re-run after fix, old data persists!

---

## 🧪 **Diagnostic Steps**

### **Step 1: Check Base Data Format**

Run "Build Market Data" and check logs for:
```
Base Data headers (first 15): [column names]
Sample employee 1: site=USA, famCode=?, execName=?, level=L5 IC, pay=85000
  → Created aon key: USA|?|L5 IC
```

**What to look for**:
- Is `famCode` populated? Or empty?
- What format is it? (e.g., "EN.SODE.P5" or something else?)
- Are AON keys being created?

---

### **Step 2: Check Index Creation**

Look for:
```
Processed X active employees, skipped Y inactive, skipped Z with missing data
Built internal index: N combinations with employee data
  USA|EN.SODE|L5 IC → min=70000, med=85000, max=100000, n=12
```

**What to look for**:
- How many employees processed?
- How many combinations created? (should be > 0)
- Sample keys - do they look correct?

---

### **Step 3: Check Lookup Attempts**

Look for:
```
Lookup 1: key="India|EN.AIML|L2 IC" found=true stats={"min":1500000,"med":1750000,"max":2000000,"n":5}
Lookup 2: key="India|EN.AIML|L3 IC" found=false stats={"min":"","med":"","max":"","n":0}
```

**What to look for**:
- Are lookups finding matches? (found=true vs found=false)
- Do the lookup keys match the format of the created keys?

---

## 🔧 **Potential Fixes**

### **Fix Option 1: Improve Key Extraction**

If Base Data Job Family Name doesn't have the level suffix:
```javascript
// Current:
const base = dot >= 0 ? famCode.slice(0, dot) : famCode;

// Enhanced:
let base = famCode;
if (dot >= 0) {
  // Check if last part looks like a level token (P5, M4, E3, etc.)
  const lastPart = famCode.slice(dot + 1);
  if (/^[PME]\d+$/i.test(lastPart)) {
    base = famCode.slice(0, dot);
  }
}
```

---

### **Fix Option 2: Add Fallback to Exec Name**

If Job Family Name extraction fails, use Mapped Family:
```javascript
// If we can't extract AON code from Job Family Name,
// try to reverse-lookup from Mapped Family
if (!base && execName) {
  // Find AON code that maps to this exec description
  const execMap = _getExecDescMap_();
  execMap.forEach((desc, code) => {
    if (desc.toUpperCase() === execName.toUpperCase()) {
      base = code;
    }
  });
}
```

---

### **Fix Option 3: Revert to Runtime Calculation**

Use the old INTERNAL_STATS approach for calculator:
- Pre-index is still used for Full List generation (fast bulk operation)
- But calculator calls INTERNAL_STATS at runtime (always accurate)

---

## 📋 **Next Steps**

1. **Run "Build Market Data"**
2. **Go to Extensions → Apps Script → Executions**
3. **Find "buildMarketData" execution**
4. **Copy ALL logs** and share them
5. Logs will show which of the 3 potential causes is the issue
6. Then we can apply the appropriate fix

---

**Status**: Waiting for log output to diagnose exact root cause


