# DEBUG: Internal Stats Not Populating (v4.6.3-debug)

## 🐛 **Current Issue**

**Symptoms**:
- ✅ CR columns (Avg CR, TT CR, New Hire CR, BT CR) are calculating correctly
- ❌ Internal stats columns (Internal Min, Internal Median, Internal Max, Emp Count) are all showing 0

**What This Tells Us**:
- `_preIndexEmployeesForCR_()` is working (reads from Employees Mapped sheet)
- `_buildInternalIndex_()` is NOT working (reads from Base Data sheet)

---

## 🔍 **Deployed Debug Version**

I've deployed **v4.6.3-debug** with comprehensive logging to diagnose the issue.

### **What It Logs**

#### **1. When Building Internal Index** (`_buildInternalIndex_()`):
```
Base Data headers (first 15): [list of column names]
Base Data column indices: Job Family Name=X, Mapped Family=X, Active/Inactive=X, Site=X, Job Level=X, Base salary=X
Sample employee 1: site=USA, famCode=EN.SODE.P5, execName=..., level=L5 IC, pay=85000
  → Created exec key: USA|ENGINEERING - SOFTWARE DEVELOPMENT|L5 IC
  → Created aon key: USA|EN.SODE|L5 IC
Sample employee 2: ...
Sample employee 3: ...
Processed 234 active employees, skipped 45 inactive, skipped 12 with missing data
Built internal index: 456 combinations with employee data
  USA|EN.SODE|L5 IC → min=70000, med=85000, max=100000, n=12
  USA|EN.PGPG|L6 IC → min=120000, med=140000, max=160000, n=8
  ...
```

#### **2. When Looking Up Stats** (`rebuildFullListAllCombinations_()`):
```
Lookup 1: key="India|EN.AIML|L2 IC" found=true stats={"min":1500000,"med":1750000,"max":2000000,"n":5}
Lookup 2: key="India|EN.AIML|L3 IC" found=false stats={"min":"","med":"","max":"","n":0}
Lookup 3: key="India|EN.AIML|L4 IC" found=true stats={"min":3000000,"med":3500000,"max":4000000,"n":8}
...
Internal stats summary: 234 out of 1440 combinations have employee data
```

---

## 🧪 **What You Need to Do**

### **Step 1: Refresh & Run**
1. **Reload your Google Sheet** (press F5)
2. Go to **Bob Data** menu → **"📊 Build Market Data (Full Lists)"**
3. Click **Yes** to confirm
4. Wait for it to complete (~30 seconds)

### **Step 2: View the Logs**
1. Go to **Extensions** → **Apps Script**
2. Click **"Executions"** (left sidebar)
3. Find the most recent **"buildMarketData"** execution
4. Click on it to expand
5. **Copy ALL the log output** and send it to me

---

## 🔎 **What I'm Looking For**

### **Scenario A: Column Not Found**
```
Base Data column indices: Job Family Name=-1, ...
ERROR: One or more required columns not found in Base Data!
```
**This means**: Base Data sheet doesn't have a column called exactly "Job Family Name"

**Solution**: We need to check what the actual column name is and update the code.

---

### **Scenario B: Wrong Data Format**
```
Sample employee 1: site=USA, famCode=, execName=Engineering - Software Development, level=L5 IC, pay=85000
  → Created exec key: USA|ENGINEERING - SOFTWARE DEVELOPMENT|L5 IC
```
**This means**: `famCode` is empty (no Aon Code in Base Data)

**Solution**: The code will fall back to exec keys only. Lookups need to match.

---

### **Scenario C: Key Format Mismatch**
```
Sample employee 1: famCode=EN.SODE.P5
  → Created aon key: USA|EN.SODE|L5 IC
...
Lookup 1: key="India|EN.AIML|L2 IC" found=false
```
**This means**: Keys being created don't match keys being looked up

**Solution**: Adjust the key generation or lookup logic.

---

### **Scenario D: No Active Employees**
```
Processed 0 active employees, skipped 567 inactive, skipped 0 with missing data
Built internal index: 0 combinations with employee data
```
**This means**: Base Data has no employees marked as "Active"

**Solution**: Check the Active/Inactive column values.

---

## 📋 **Questions to Answer from Logs**

Please check the logs and tell me:

1. **Column Detection**:
   - What column indices were found? (especially Job Family Name)
   - Any ERROR messages about columns?

2. **Sample Employee Data**:
   - What does `famCode` look like? (e.g., "EN.SODE.P5" or empty or something else?)
   - What does `execName` look like?
   - What keys are being created?

3. **Processing Summary**:
   - How many active employees were processed?
   - How many combinations in the final index?

4. **Lookup Results**:
   - Are the first 5 lookups showing `found=true` or `found=false`?
   - Do the lookup keys match the format of the created keys?

5. **Final Summary**:
   - What's the ratio? (e.g., "234 out of 1440 combinations have employee data")

---

## 🎯 **Expected Behavior**

### **If Everything Works Correctly**:
```
Base Data column indices: Job Family Name=5, Mapped Family=-1, Active/Inactive=8, Site=3, Job Level=7, Base salary=12
Sample employee 1: site=USA, famCode=EN.SODE.P5, execName=Engineering - Software Development, level=L5 IC, pay=85000
  → Created exec key: USA|ENGINEERING - SOFTWARE DEVELOPMENT|L5 IC
  → Created aon key: USA|EN.SODE|L5 IC
Processed 234 active employees, skipped 45 inactive, skipped 12 with missing data
Built internal index: 456 combinations with employee data
...
Lookup 1: key="India|EN.AIML|L2 IC" found=true stats={"min":1500000,"med":1750000,"max":2000000,"n":5}
Lookup 2: key="India|EN.AIML|L3 IC" found=false stats={"min":"","med":"","max":"","n":0}
...
Internal stats summary: 234 out of 1440 combinations have employee data
```

---

## 🚀 **Next Steps**

Once you send me the logs, I'll be able to:
1. **Identify the exact issue** (column name, data format, key mismatch, etc.)
2. **Implement the fix** (update column detection, key generation, or lookup logic)
3. **Deploy v4.6.4** with the actual fix
4. **Test to confirm** internal stats populate correctly

---

**Status**: Waiting for log output  
**Version**: 4.6.3-debug  
**Action Required**: Run "Build Market Data" and send me the execution logs

