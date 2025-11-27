# Legacy Mappings Feedback Loop

## 🔄 How It Works

The system creates a **continuous improvement cycle** where approved employee mappings automatically update the legacy data for future imports.

```
┌─────────────────────────────────────────────────────────┐
│                     FEEDBACK LOOP                        │
└─────────────────────────────────────────────────────────┘

1. Import Bob Data
   ↓
   Uses Legacy Mappings (400+ employees)
   ↓
2. Employees Mapped Sheet
   - 400 employees: 100% confidence (Legacy source)
   - New employees: 95% confidence (Title-Based)
   - Status: "Needs Review"
   ↓
3. User Reviews Mappings
   - Approves correct suggestions
   - Fixes incorrect mappings
   - Status: "Needs Review" → "Approved"
   ↓
4. Auto-Update Legacy Mappings
   - All approved entries synced back
   - Existing entries updated
   - New employees added to legacy
   ↓
5. Next Import Bob Data
   - Uses updated Legacy Mappings
   - Even more employees have 100% confidence
   - Less manual work required
   ↓
   (Loop continues...)
```

---

## 📊 Benefits Over Time

### **First Import**
- Legacy: 400 employees (100% confidence)
- New: 50 employees (95% title-based or 0% unmapped)
- Manual work: Review 50 employees

### **Second Import (After Approvals)**
- Legacy: 450 employees (100% confidence) ← Updated!
- New: 10 employees (95% or 0%)
- Manual work: Review 10 employees

### **Third Import**
- Legacy: 460 employees (100% confidence)
- New: 5 employees
- Manual work: Review 5 employees

**Result**: Each import requires less manual work as the legacy data improves!

---

## 🔧 Technical Implementation

### **Auto-Sync During Import**
```javascript
importBobData()
  ↓
  Step 6: syncEmployeesMappedSheet_()
    - Reads Legacy Mappings
    - Creates Employees Mapped with suggestions
  ↓
  Step 7: updateLegacyMappingsFromApproved_()
    - Reads approved entries from Employees Mapped
    - Updates Legacy Mappings sheet
    - Converts CIQ Level → Aon token (L5 IC → P5)
```

### **Manual Trigger (Optional)**
```
Menu → 🔧 Tools → Update Legacy Mappings from Approved
```
Use this if you:
- Approved more mappings after the import
- Want to manually sync back to legacy data
- Need to rebuild legacy reference

---

## 📝 Mapping Conversion

### **CIQ Level → Aon Token**

| CIQ Level | Aon Token | Example |
|-----------|-----------|---------|
| L2 IC | P2 | EN.SODE.P2 |
| L3 IC | P3 | EN.SODE.P3 |
| L4 IC | P4 | EN.SODE.P4 |
| L5 IC | P5 | EN.SODE.P5 |
| L6 IC | P6 | EN.SODE.P6 |
| L7 IC | E1 | EN.SODE.E1 |
| L4 Mgr | M3 | SA.CRCS.M3 |
| L5 Mgr | M4 | SA.CRCS.M4 |
| L6 Mgr | M5 | SA.CRCS.M5 |
| L6.5 Mgr | M6 | SA.CRCS.M6 |
| L7 Mgr | E1 | EN.PGPG.E1 |
| L8 Mgr | E3 | EN.PGHC.E3 |
| L9 Mgr | E5 | LE.GLEC.E5 |

### **Example Flow**

**Employee Approved**:
- Employee ID: 20999
- Aon Code: EN.SODE
- Level: L5 IC
- Status: Approved

**Synced to Legacy Mappings**:
```
Employee ID | Job Family | Full Mapping
20999       | EN.SODE    | EN.SODE.P5
```

**Next Import**:
- Employee 20999 found in Legacy Mappings
- Confidence: 100%
- Source: Legacy
- Auto-mapped without manual review

---

## 🎯 Best Practices

### **1. Regular Approval**
After each Import Bob Data:
- Run **Review Employee Mappings**
- Approve all yellow rows that look correct
- Fix any incorrect suggestions
- Changes auto-sync to Legacy Mappings

### **2. Bulk Approval**
For employees with high confidence:
- Filter Status = "Needs Review"
- Filter Confidence = "100%" or "95%"
- Review quickly
- Bulk change Status to "Approved"

### **3. Manual Updates**
If you manually approve 50 employees:
```
Tools → Update Legacy Mappings from Approved
```
This immediately syncs them to Legacy Mappings (no need to wait for next import)

### **4. Data Quality**
The feedback loop means:
- Correct mappings get reinforced
- New employees benefit from historical patterns
- Legacy data stays current automatically
- Less manual work over time

---

## 🐛 Troubleshooting

### **Legacy Mappings not updating**
- Check that Status = "Approved" in Employees Mapped
- Ensure Aon Code and Level are filled in
- Run manual sync: Tools → Update Legacy Mappings from Approved

### **Confidence stuck at low %**
- Employee not in Legacy Mappings
- Title not in Title Mapping
- Approve the mapping to add to legacy data

### **Old legacy data not overwriting**
- System ONLY updates from approved entries
- Updates existing + adds new
- Never deletes from Legacy Mappings
- To reset: manually delete rows from Legacy Mappings

---

## 📌 Summary

✅ **Auto-updating**: Approved mappings sync back to Legacy Mappings  
✅ **Continuous improvement**: Each import easier than the last  
✅ **100% confidence**: Legacy data always current  
✅ **Less manual work**: System learns from your approvals  
✅ **Feedback loop**: Your work compounds over time  

---

**Version**: 3.9.0  
**Date**: November 27, 2025  
**Status**: ✅ Active

