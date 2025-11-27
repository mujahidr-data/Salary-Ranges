# HOTFIX v4.6.2 - Internal Stats Bug Fix (Min/Med/Max/Count)

## 🐛 **Bug Discovered**

The Full List's **Internal Min, Internal Median, Internal Max, and Employee Count** columns were showing:
- ❌ All values blank/empty
- ❌ Even when employees existed in Base Data

## 🔍 **Root Cause Analysis**

Found **TWO critical bugs** preventing internal stats from working:

### **Bug #1: Region Key Mismatch**

**The Problem**: Different region codes in different parts of the system!

#### **`_buildInternalIndex_()`** (builds the index):
```javascript
// Normalizes to "USA"
const normSite = site === 'India' ? 'India' : (site === 'USA' ? 'USA' : (site === 'UK' ? 'UK' : site));
const key = `${normSite}|${b}|${ciq}`;
// Creates keys like: "USA|EN.SODE|L5 IC"
```

#### **`rebuildFullListAllCombinations_()`** (looks up stats):
```javascript
const regions = ['India', 'US', 'UK'];  // ❌ Uses "US", not "USA"
const intKey = `${region}|${aonCode}|${ciqLevel}`;
// Tries to lookup: "US|EN.SODE|L5 IC"  ← DOES NOT EXIST!
```

**Result**: Lookups always failed because:
- Index contains: `"USA|EN.SODE|L5 IC"`
- Lookup searches for: `"US|EN.SODE|L5 IC"`
- **No match!** Returns default: `{ min: '', med: '', max: '', cnt: 0 }`

---

### **Bug #2: Property Name Mismatch**

**The Problem**: Different property names in create vs access!

#### **`_buildInternalIndex_()`** (returns data):
```javascript
buckets.forEach((arr, key) => {
  arr.sort((a,b)=>a-b);
  const n = arr.length;
  const min = arr[0], max = arr[n-1];
  const med = n % 2 ? arr[(n-1)/2] : (arr[n/2 - 1] + arr[n/2]) / 2;
  out.set(key, { min, med, max, n });  // ← Returns "n"
});
```

#### **`rebuildFullListAllCombinations_()`** (accesses data):
```javascript
const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', cnt: 0 };
// ...
rows.push([
  // ...
  intStats.min,
  intStats.med,
  intStats.max,
  intStats.cnt,  // ← Accesses "cnt" (WRONG!)
  // ...
]);
```

**Result**: Even if lookup succeeded (it didn't because of Bug #1), the count would still be wrong:
- Returns: `{ min: 50000, med: 75000, max: 100000, n: 5 }`
- Accesses: `intStats.cnt` → **undefined**

---

## ✅ **Fixes Applied**

### **Fix #1: Region Normalization**

```javascript
// ✅ AFTER (v4.6.2): Normalize region before lookup
const intRegion = region === 'US' ? 'USA' : region;
const intKey = `${intRegion}|${aonCode}|${ciqLevel}`;
const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', n: 0 };
```

**Now the keys match**:
- Index key: `"USA|EN.SODE|L5 IC"`
- Lookup key: `"USA|EN.SODE|L5 IC"` ✅

---

### **Fix #2: Property Name Correction**

```javascript
// ✅ AFTER (v4.6.2): Use correct property name
rows.push([
  // ...
  intStats.min,
  intStats.med,
  intStats.max,
  intStats.n,  // ← Changed from `cnt` to `n`
  // ...
]);
```

**Now the property names match**:
- Returns: `{ min: 50000, med: 75000, max: 100000, n: 5 }`
- Accesses: `intStats.n` → `5` ✅

---

### **Bonus: Debug Logging**

Added comprehensive logging to `_buildInternalIndex_()`:

```javascript
Logger.log(`Built internal index: ${out.size} combinations with employee data`);
// Log first 5 for verification
let count = 0;
out.forEach((stats, key) => {
  if (count < 5) {
    Logger.log(`  ${key} → min=${stats.min}, med=${stats.med}, max=${stats.max}, n=${stats.n}`);
    count++;
  }
});
```

**Example Output**:
```
Built internal index: 234 combinations with employee data
  USA|EN.SODE|L5 IC → min=70000, med=85000, max=100000, n=12
  USA|EN.PGPG|L6 IC → min=120000, med=140000, max=160000, n=8
  UK|SA.CRCS|L5 IC → min=45000, med=55000, max=65000, n=15
  India|TE.DADA|L4 IC → min=600000, med=750000, max=900000, n=6
  USA|FI.ACCO|L5 Mgr → min=80000, med=95000, max=110000, n=4
```

**This helps verify**:
- How many employee groups were indexed
- Sample keys and values
- If region normalization is working

---

## 📊 **Expected Results After Fix**

### **Before (v4.6.1)**:
| Job Family | Level | Internal Min | Internal Median | Internal Max | Emp Count |
|------------|-------|--------------|-----------------|--------------|-----------|
| Engineering - Software Development | L5 IC | (blank) | (blank) | (blank) | 0 |
| Sales - Customer Success | L5 IC | (blank) | (blank) | (blank) | 0 |

### **After (v4.6.2)**:
| Job Family | Level | Internal Min | Internal Median | Internal Max | Emp Count |
|------------|-------|--------------|-----------------|--------------|-----------|
| Engineering - Software Development | L5 IC | 70000 | 85000 | 100000 | 12 |
| Sales - Customer Success | L5 IC | 50000 | 65000 | 80000 | 8 |

---

## 🧪 **Testing Checklist**

After deploying, verify:
- [ ] **Internal Min** column has numeric values (or blank if no employees)
- [ ] **Internal Median** column has numeric values (or blank if no employees)
- [ ] **Internal Max** column has numeric values (or blank if no employees)
- [ ] **Emp Count** column shows numbers > 0 where employees exist
- [ ] Check Logger output to see how many combinations were indexed
- [ ] Sample log entries show correct region keys ("USA", not "US")

---

## 🔧 **Technical Details**

### **Region Normalization Logic**

| Input Region | Normalized for Internal Index | Used in Lookup |
|-------------|-------------------------------|----------------|
| "India" | "India" | "India" ✅ |
| "UK" | "UK" | "UK" ✅ |
| "US" | "USA" | "USA" ✅ (now normalized) |
| "USA" | "USA" | "USA" ✅ |

### **Object Property Names**

| Property | Internal Index (`_buildInternalIndex_()`) | Full List Usage (`rebuildFullListAllCombinations_()`) |
|----------|-------------------------------------------|-----------------------------------------------------|
| **Minimum** | `min` | `intStats.min` ✅ |
| **Median** | `med` | `intStats.med` ✅ |
| **Maximum** | `max` | `intStats.max` ✅ |
| **Count** | `n` | `intStats.n` ✅ (fixed from `cnt`) |

### **Default Values**

When no employees exist for a combination:
```javascript
const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', n: 0 };
```

**Result**:
- `min`, `med`, `max` → empty string (displays as blank in sheet)
- `n` → 0 (displays as 0 in sheet)

---

## 📝 **Code Changes Summary**

### **File: `SalaryRangesCalculator.gs`**

#### **Line ~4922-4924** (Key Lookup):
```javascript
// BEFORE (v4.6.1)
const intKey = `${region}|${aonCode}|${ciqLevel}`;
const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', cnt: 0 };

// AFTER (v4.6.2)
const intRegion = region === 'US' ? 'USA' : region;
const intKey = `${intRegion}|${aonCode}|${ciqLevel}`;
const intStats = internalIndex.get(intKey) || { min: '', med: '', max: '', n: 0 };
```

#### **Line ~4991-4994** (Data Access):
```javascript
// BEFORE (v4.6.1)
intStats.min,
intStats.med,
intStats.max,
intStats.cnt,  // ← WRONG

// AFTER (v4.6.2)
intStats.min,
intStats.med,
intStats.max,
intStats.n,    // ← FIXED
```

#### **Line ~1350-1357** (Debug Logging):
```javascript
// ADDED (v4.6.2)
Logger.log(`Built internal index: ${out.size} combinations with employee data`);
// Log first 5 for verification
let count = 0;
out.forEach((stats, key) => {
  if (count < 5) {
    Logger.log(`  ${key} → min=${stats.min}, med=${stats.med}, max=${stats.max}, n=${stats.n}`);
    count++;
  }
});
```

---

## 🚀 **Deployment Steps**

1. ✅ **Code updated** with region normalization and property name fix
2. ✅ **Debug logging added** to `_buildInternalIndex_()`
3. ✅ **Version updated** to 4.6.2
4. ⏳ **Commit to Git**
5. ⏳ **Push to GitHub**
6. ⏳ **Deploy to Apps Script** via `clasp push`
7. ⏳ **Test in Google Sheets**

---

## 🔍 **How to View Debug Logs**

After running "Build Market Data":

1. Go to **Extensions** → **Apps Script**
2. Click **Executions** (left sidebar)
3. Find the most recent "buildMarketData" execution
4. Click on it to expand
5. Look for logs like:
   ```
   Built internal index: 234 combinations with employee data
     USA|EN.SODE|L5 IC → min=70000, med=85000, max=100000, n=12
     USA|EN.PGPG|L6 IC → min=120000, med=140000, max=160000, n=8
     ...
   ```

This will confirm:
- ✅ How many employee groups were found
- ✅ Region keys are correct ("USA", not "US")
- ✅ Stats are being calculated correctly

---

## 📈 **Impact**

| Metric | Before | After | Fix |
|--------|--------|-------|-----|
| **Internal Min** | All blank | Populated | ✅ Region normalization |
| **Internal Median** | All blank | Populated | ✅ Region normalization |
| **Internal Max** | All blank | Populated | ✅ Region normalization |
| **Emp Count** | All 0 | Correct counts | ✅ Property name fix (`n` not `cnt`) |

---

**Version**: 4.6.2  
**Priority**: CRITICAL (data accuracy)  
**Impact**: Fixes internal employee statistics in Full List  
**Backward Compatible**: Yes  
**Deployment**: Immediate

