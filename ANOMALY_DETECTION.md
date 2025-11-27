# 🔍 Anomaly Detection Guide

## Overview

The Employees Mapped sheet now includes **automatic anomaly detection** to flag inconsistencies and outliers in employee mappings. This helps identify data quality issues and incorrect mappings.

---

## ✨ New Columns

### **Column N: Level Anomaly** 🟧
**Purpose**: Flags when CIQ level doesn't match expected Aon level

**Example Anomalies:**
```
Employee: John Doe
CIQ Level: L5 IC
Aon Code: EN.SODE.M4
Level Anomaly: "Expected P5, got M4"
```

This indicates:
- Employee is L5 IC (individual contributor level 5)
- But mapped to M4 (manager level 4)
- **Should be**: EN.SODE.P5

**Color**: 🟧 Orange background

---

### **Column O: Title Anomaly** 🟪
**Purpose**: Flags when employee's mapping differs from others with same job title

**Example Anomalies:**
```
Employee: Jane Smith
Job Title: Senior Software Engineer
Aon Code: SA.CRCS
Level: L4 IC
Title Anomaly: "15 others: EN.SODE L5 IC"
```

This indicates:
- 15 other "Senior Software Engineer" employees are mapped to EN.SODE L5 IC
- But Jane is mapped to SA.CRCS L4 IC
- **Possible issue**: Wrong title, wrong mapping, or legitimate exception

**Color**: 🟪 Purple background

---

## 🎯 Why Anomaly Detection?

### **Problem 1: Level Mismatches**
```
❌ Wrong: Individual Contributor (IC) mapped to Manager (M) level
❌ Wrong: Manager mapped to IC level
❌ Wrong: L5 IC mapped to P3 (should be P5)
```

### **Problem 2: Title Inconsistencies**
```
❌ Wrong: Same title, different mappings across team
❌ Wrong: Employee promoted but old mapping retained
❌ Wrong: Copy-paste error from another employee
```

### **Solution:**
Automated detection highlights these issues immediately, no manual cross-checking needed!

---

## 🔍 How It Works

### **Level Anomaly Detection:**

```javascript
1. Extract CIQ Level: "L5 IC"
2. Convert to expected Aon token: "P5"
3. Extract actual Aon token from mapping: "EN.SODE.M4" → "M4"
4. Compare: "P5" vs "M4"
5. If different: Flag anomaly
```

**Conversion Rules:**
| CIQ Level | Expected Token | Example Mapping |
|-----------|----------------|-----------------|
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

---

### **Title Anomaly Detection:**

```javascript
1. Group employees by job title
2. Count frequency of each (Aon Code + Level) combination per title
3. Find most common mapping for each title
4. Compare each employee to their title's common mapping
5. If different: Flag anomaly with count
```

**Example:**
```
Job Title: "Senior Software Engineer"
Mappings found:
  - EN.SODE L5 IC: 15 employees ← Most common
  - EN.SODE L4 IC: 2 employees
  - SA.CRCS L5 IC: 1 employee  ← Flagged

Result: 1 employee flagged with "15 others: EN.SODE L5 IC"
```

---

## 📊 Using Anomalies to Review Mappings

### **Workflow:**

```
1. Run Import Bob Data
   ↓
2. Open Employees Mapped sheet
   ↓
3. Filter by anomalies:
   - Sort by Column N (Level Anomaly)
   - Sort by Column O (Title Anomaly)
   ↓
4. Review each flagged row:
   - Is this a legitimate exception?
   - Is the mapping wrong?
   - Is the title wrong?
   ↓
5. Fix issues:
   - Update Aon Code/Level if wrong
   - Add notes explaining exceptions
   - Change Status to "Approved" if correct
```

---

### **Common Scenarios:**

#### **Scenario 1: Promoted Employee**
```
Employee: John Doe
Title: "Senior Software Engineer"
Level: L6 IC (promoted!)
Mapping: EN.SODE.P5 (old)
Title Anomaly: "15 others: EN.SODE L5 IC"
Level Anomaly: "Expected P6, got P5"

Action: Update mapping to EN.SODE.P6
```

#### **Scenario 2: Manager with IC Title**
```
Employee: Jane Smith
Title: "Senior Software Engineer" (title not updated)
Level: L5 Mgr (actually a manager)
Mapping: EN.SODE.M4
Title Anomaly: "20 others: EN.SODE L5 IC"
Level Anomaly: None (M4 is correct for L5 Mgr)

Action: Update title in HiBob or add note explaining exception
```

#### **Scenario 3: Unique Role**
```
Employee: Alex Johnson
Title: "Principal Data Science Architect"
Level: L7 IC
Mapping: TE.DADA.E1
Title Anomaly: None (only employee with this title)
Level Anomaly: None (E1 is correct for L7 IC)

Action: No action needed, unique role is correctly mapped
```

#### **Scenario 4: Wrong Mapping**
```
Employee: Maria Garcia
Title: "Customer Success Manager"
Level: L5 IC
Mapping: EN.SODE.P5 (wrong family!)
Title Anomaly: "10 others: SA.CRCS L5 IC"
Level Anomaly: None (P5 matches L5 IC)

Action: Fix Aon Code to SA.CRCS.P5
```

---

## 🎨 Visual Indicators

### **Color Coding:**

| Color | Meaning | Column | Example |
|-------|---------|--------|---------|
| 🟩 Green | Approved | All | Employee mapping is approved |
| 🟨 Yellow | Needs Review | All | Default status, needs verification |
| 🟥 Red | Rejected/Missing | All | Rejected or blank mapping |
| 🟧 Orange | Level Anomaly | N | CIQ level doesn't match Aon level |
| 🟪 Purple | Title Anomaly | O | Mapping differs from title peers |

### **Quick Filters:**

**View only anomalies:**
1. Click column N header
2. Filter → Filter by condition → Is not empty
3. Repeat for column O

**View only severe issues:**
1. Filter Status = "Needs Review"
2. Filter Level Anomaly → Is not empty
3. OR Title Anomaly → Is not empty

---

## 📋 Best Practices

### **✅ DO:**
- Review all orange (level) anomalies first - often critical errors
- Investigate purple (title) anomalies with high counts ("20 others")
- Document legitimate exceptions with notes
- Update HiBob data if titles are wrong
- Approve correct mappings after review

### **❌ DON'T:**
- Ignore anomalies without investigation
- Assume all anomalies are errors (some are legitimate)
- Bulk approve without reviewing flagged rows
- Change mappings without understanding why they differ
- Delete anomaly columns (they're auto-regenerated)

---

## 🔧 Troubleshooting

### **Too Many False Positives?**

**Level Anomalies:**
- Check if CIQ levels in HiBob are correct
- Verify Lookup sheet has correct level mappings
- Some executives may have non-standard mappings (expected)

**Title Anomalies:**
- If job titles are inconsistent in HiBob, many false positives
- Example: "Sr Software Engineer" vs "Senior Software Engineer"
- Solution: Standardize job titles in HiBob

### **No Anomalies Detected?**

**Possible causes:**
1. All mappings are correct (great!)
2. Not enough legacy data to compare
3. All employees have unique titles
4. Fresh Build not run (Legacy Mappings empty)

**Check:**
- Legacy Mappings sheet has 400+ rows
- Employees Mapped has multiple employees with same title
- Level and Aon Code columns are populated

---

## 📊 Example Sheet

```
| Emp ID | Name       | Title                  | Aon Code | Job Family         | Level   | Status         | Level Anomaly       | Title Anomaly           |
|--------|------------|------------------------|----------|-------------------|---------|----------------|---------------------|-------------------------|
| 20616  | John Doe   | Software Engineer II   | EN.SODE  | Engineering - SW  | L4 IC   | Approved       |                     |                         |
| 20999  | Jane Smith | Software Engineer II   | EN.SODE  | Engineering - SW  | L4 IC   | Needs Review   |                     |                         |
| 21000  | Alex Wong  | Software Engineer II   | SA.CRCS  | Sales - CS        | L4 IC   | Needs Review   |                     | 15 others: EN.SODE L4 IC| 🟪
| 21001  | Maria Lee  | Engineering Manager    | EN.PGPG  | Engineering - PM  | L5 IC   | Needs Review   | Expected M4, got P5 |                         | 🟧
| 21002  | Tom Chen   | Senior Data Scientist  | TE.DADA  | Technology - Data | L6 IC   | Approved       |                     |                         |
```

**Issues Identified:**
- Row 3 (Alex): Wrong job family (SA.CRCS instead of EN.SODE)
- Row 4 (Maria): Manager role but mapped to IC level

---

## 📚 Related Documentation

- `EMPLOYEES_MAPPED_GUIDE.md` - Full guide to Employees Mapped sheet
- `SMART_MAPPING_GUIDE.md` - Smart mapping and confidence scores
- `LEGACY_MAPPINGS.md` - Legacy mapping system

---

## 🔄 Updates

| Version | Date | Changes |
|---------|------|---------|
| 4.1.0 | Nov 27, 2025 | Added anomaly detection columns |
| 4.0.0 | Nov 27, 2025 | Added time-based trigger support |

---

**Status**: ✅ Active  
**Last Updated**: November 27, 2025

