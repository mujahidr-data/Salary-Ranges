# 🔒 Persistent Legacy Mappings Storage

## 🎯 Overview

Legacy Mappings are now stored in **Script Properties** (permanent storage), not just in the sheet. This means your approved employee mappings **survive sheet deletion** and persist across Fresh Builds.

---

## ✨ What Changed

### **Before (v4.2.0 and earlier):**
```
Legacy Mappings stored in: Sheet only
If you delete the sheet: Data lost forever
Fresh Build loads from: Embedded code (400 employees, static)
Approved mappings: Lost if sheet deleted
```

### **After (v4.3.0):**
```
Legacy Mappings stored in: Script Properties (permanent)
If you delete the sheet: Data survives in storage
Fresh Build loads from: Script Properties (latest approved data)
Approved mappings: Persist forever ✓
```

---

## 🔄 Complete Flow

### **Initial Setup (First Time):**
```
1. Fresh Build runs
   ↓
2. Checks Script Properties
   → No legacy data found
   ↓
3. Loads embedded data (400 employees)
   ↓
4. Saves to Script Properties
   ↓
5. Creates Legacy Mappings sheet
   ✅ Storage: 400 employees
   ✅ Sheet: 400 employees
```

### **Approve New Mappings:**
```
1. User approves 50 new employees in Employees Mapped
   ↓
2. Import Bob Data runs (or manual trigger)
   ↓
3. Update Legacy Mappings from Approved runs
   ↓
4. Updates Legacy Mappings sheet
   ↓
5. Saves to Script Properties
   ✅ Storage: 450 employees (updated!)
   ✅ Sheet: 450 employees
```

### **Delete Sheets & Rebuild:**
```
1. User deletes all sheets (including Legacy Mappings)
   ↓
2. Fresh Build runs
   ↓
3. Checks Script Properties
   → 450 employees found! ✓
   ↓
4. Loads from Script Properties (not embedded data!)
   ↓
5. Creates Legacy Mappings sheet
   ✅ Storage: 450 employees (preserved!)
   ✅ Sheet: 450 employees (restored!)
```

**Your approved mappings are back!** 🎉

---

## 📊 Where Data Is Stored

### **Script Properties:**
```
Location: Project Settings → Script Properties
Keys:
  - LEGACY_MAPPINGS_CHUNKS: "3"
  - LEGACY_MAPPINGS_0: "[JSON chunk 1]"
  - LEGACY_MAPPINGS_1: "[JSON chunk 2]"
  - LEGACY_MAPPINGS_2: "[JSON chunk 3]"
  - LEGACY_MAPPINGS_UPDATED: "2025-11-27T14:30:00.000Z"

Format: JSON array chunked into 8KB segments
Limit: 500KB total (enough for ~5000+ employees)
Persistence: Forever (unless manually deleted)
```

### **Legacy Mappings Sheet:**
```
Purpose: Visual display of storage data
Updates: In sync with Script Properties
Can be deleted: Yes (will be recreated from storage)
Source of truth: No (Script Properties is)
```

---

## 🔧 Technical Implementation

### **Save to Storage:**
```javascript
_saveLegacyMappingsToStorage_(legacyData)
  ↓
1. Convert array to JSON string
2. Chunk into 8KB segments (Script Properties limit)
3. Save each chunk: LEGACY_MAPPINGS_0, LEGACY_MAPPINGS_1, ...
4. Save chunk count: LEGACY_MAPPINGS_CHUNKS
5. Save timestamp: LEGACY_MAPPINGS_UPDATED
```

### **Load from Storage:**
```javascript
_loadLegacyMappingsFromStorage_()
  ↓
1. Read chunk count
2. Reconstruct JSON from all chunks
3. Parse JSON → array
4. Return legacy data
```

### **Auto-Save Triggers:**
- ✅ `updateLegacyMappingsFromApproved_()` - After syncing approved mappings
- ✅ `createLegacyMappingsSheet_()` - On first load (if storage empty)
- ✅ Part of Import Bob Data workflow (automatic)

---

## 🎯 Benefits

| Scenario | Before | After |
|----------|--------|-------|
| **Delete sheets** | Lose all approved mappings | Mappings restored from storage ✓ |
| **Fresh Build** | Always starts with 400 static employees | Loads latest approved data ✓ |
| **Approve 100 new** | Must keep sheet safe | Auto-saved to storage ✓ |
| **Collaborate** | Hard to sync approved mappings | Storage shared across users ✓ |
| **Disaster recovery** | Manual re-mapping needed | Automatic restoration ✓ |

---

## 📋 User Experience

### **Scenario 1: Accidental Sheet Deletion**
```
User: "Oh no! I accidentally deleted all sheets!"
  ↓
Run: Fresh Build
  ↓
System: Loads 450 employees from storage
  ↓
User: "All my approved mappings are back!" ✅
```

### **Scenario 2: Clean Slate Testing**
```
User: "I want to start fresh and test the setup"
  ↓
Delete all sheets
  ↓
Run: Fresh Build
  ↓
System: Your approved mappings still there (from storage)
  ↓
User: "Perfect, I don't have to re-map everything!" ✅
```

### **Scenario 3: Continuous Improvement**
```
Week 1: Approve 50 mappings → Saved to storage
Week 2: Delete & rebuild → 50 mappings restored
Week 3: Approve 30 more → Saved to storage (now 80)
Week 4: Delete & rebuild → 80 mappings restored
```

**Your work compounds over time and never gets lost!** 🚀

---

## 🔍 Viewing Storage Data

### **Option 1: Extensions → Apps Script**
```
1. Extensions → Apps Script
2. Project Settings (gear icon)
3. Script Properties
4. See: LEGACY_MAPPINGS_CHUNKS, LEGACY_MAPPINGS_0, etc.
```

### **Option 2: Execution Log**
```
After Fresh Build or Import:
1. Extensions → Apps Script → Executions
2. Click on latest execution
3. View logs:
   "Loaded 450 legacy mappings from storage (last updated: 2025-11-27...)"
```

---

## 🔧 Managing Storage

### **View Last Update Time:**
```javascript
// From Apps Script editor:
const updated = PropertiesService.getScriptProperties().getProperty('LEGACY_MAPPINGS_UPDATED');
Logger.log(updated); // "2025-11-27T14:30:00.000Z"
```

### **Manually Reset Storage:**
```javascript
// Only if you want to completely reset (rare!)
const props = PropertiesService.getScriptProperties();
const chunkCount = parseInt(props.getProperty('LEGACY_MAPPINGS_CHUNKS') || '0');
for (let i = 0; i < chunkCount; i++) {
  props.deleteProperty(`LEGACY_MAPPINGS_${i}`);
}
props.deleteProperty('LEGACY_MAPPINGS_CHUNKS');
props.deleteProperty('LEGACY_MAPPINGS_UPDATED');
```

### **Force Reload from Embedded Data:**
```
1. Delete Script Properties (see above)
2. Delete Legacy Mappings sheet
3. Run Fresh Build
4. Will load embedded data → Save to storage
```

---

## 🐛 Troubleshooting

### **Legacy Mappings sheet is empty after Fresh Build**

**Possible causes:**
1. Script Properties storage corrupted
2. Storage limit exceeded (500KB)
3. JSON parse error

**Solution:**
```
1. Extensions → Apps Script → Executions
2. Check for errors in Fresh Build execution
3. If corrupted, manually delete Script Properties
4. Run Fresh Build again (will reload from embedded data)
```

### **Old data keeps coming back**

**This is expected!** Storage is persistent. If you want to use newer embedded data:
1. Manually delete Script Properties
2. Run Fresh Build

---

## 📚 Related Functions

| Function | Purpose | Updates Storage? |
|----------|---------|------------------|
| `_saveLegacyMappingsToStorage_()` | Save to Script Properties | ✅ Yes |
| `_loadLegacyMappingsFromStorage_()` | Load from Script Properties | ❌ No (read-only) |
| `createLegacyMappingsSheet_()` | Create sheet from storage | ⚠️ Only if empty |
| `updateLegacyMappingsFromApproved_()` | Sync approved mappings | ✅ Yes |
| `_getLegacyMapping_(empID)` | Get mapping for one employee | ❌ No (read-only) |

---

## 🎯 Summary

✅ **Persistent** - Survives sheet deletion  
✅ **Automatic** - No manual intervention  
✅ **Efficient** - Cached and chunked  
✅ **Safe** - Fallback to sheet if storage fails  
✅ **Transparent** - Works seamlessly in background  
✅ **Recoverable** - Can reset if needed  

Your approved mappings are now **truly permanent**! 🔒

---

**Version**: 4.3.0  
**Date**: November 27, 2025  
**Status**: ✅ Active

