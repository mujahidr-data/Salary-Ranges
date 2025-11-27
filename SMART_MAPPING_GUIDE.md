# Smart Employee Mapping Guide

## Overview
The system now uses intelligent mapping with multiple data sources and approval workflows to ensure accurate employee-to-Aon code assignments.

---

## 📊 Data Sources (Priority Order)

### 1. **Legacy Mappings** (Highest Priority)
- **Sheet**: `Legacy Mappings`
- **Confidence**: 100%
- **Source**: Your provided historical data
- **Columns**: `Employee ID | Job Family (Base) | Full Mapping`
- **Example**: `20560 | EN.SODE | EN.SODE.P5`

### 2. **Title-Based Mapping**
- **Sheet**: `Title Mapping`
- **Confidence**: 95%
- **Source**: Manual mappings by job title
- **Columns**: `Job Title | Aon Code | Level`
- **Example**: `Senior Software Engineer | EN.SODE | L5 IC`

### 3. **Manual Entries**
- **Confidence**: 50% (unless approved)
- **Source**: Direct manual input in Employees Mapped sheet

---

## 🔄 Workflow

### **Step 1: Import Bob Data**
```
Menu → 📥 Import Bob Data
```
This will:
- Import Base Data, Bonus, Comp, Performance Ratings from HiBob
- Auto-sync **Employees Mapped** sheet with smart suggestions
- Show counts: Approved, Legacy, Title-Based, Needs Review

### **Step 2: Review Mappings**
```
Menu → ✅ Review Employee Mappings
```
Shows summary:
- ✅ **Approved**: X employees (Y%)
- ⚠️ **Needs Review**: Z employees
- ❌ **Rejected**: W employees

Color coding in sheet:
- 🟢 **Green rows** = Approved
- 🟡 **Yellow rows** = Needs Review
- 🔴 **Red rows** = Rejected or Missing mapping

### **Step 3: Approve/Reject Mappings**
1. Open `Employees Mapped` sheet
2. Review each row:
   - **Aon Code**: Job family (e.g., EN.SODE, SA.CRCS)
   - **Level**: CIQ level (e.g., L5 IC, L4 Mgr)
   - **Confidence**: Mapping confidence (100%, 95%, 50%)
   - **Source**: Where mapping came from (Legacy, Title-Based, Manual)
3. Use **Status** dropdown:
   - `Needs Review` → `Approved` (if correct)
   - `Needs Review` → `Rejected` (if incorrect)
4. For rejected/missing:
   - Manually enter correct Aon Code and Level
   - Change Status to `Approved`

### **Step 4: Build Market Data**
```
Menu → 📊 Build Market Data
```
This will:
- Generate Full List with all X0/Y1 combinations
- Calculate CR values for each combination:
  * **Avg CR**: All approved active employees
  * **TT CR**: Employees with AYR 2024 = "HH"
  * **BT CR**: Employees with AYR 2024 IN ("ML", "NI")
  * **New Hire CR**: Employees with Start Date within last 365 days
- Populate both calculators with data

---

## 📋 Employees Mapped Sheet Columns

| Column | Description | Example |
|--------|-------------|---------|
| Employee ID | From Base Data | 20560 |
| Employee Name | Display name | Aashi Mittal |
| Job Title | Current title | Senior Software Engineer |
| Department | Department | Engineering |
| Site | Location | India |
| **Aon Code** | Job family code | EN.SODE |
| **Level** | CIQ level | L5 IC |
| **Confidence** | Mapping confidence | 100%, 95%, 50% |
| **Source** | Origin of mapping | Legacy, Title-Based, Manual |
| **Status** | Review status | Needs Review, Approved, Rejected |
| Base Salary | Current salary | 5000000 |
| Start Date | Hire date | 2025-05-06 |

---

## 🎯 CR Calculations Explained

### **Avg CR** (Average Compa Ratio)
- **Filter**: Status = "Approved" AND Active employees
- **Formula**: Average(Salary / Range Mid)
- **Purpose**: Shows average positioning vs. market mid-point

### **TT CR** (Top Talent Compa Ratio)
- **Filter**: Status = "Approved" AND AYR 2024 = "HH"
- **Formula**: Average(Salary / Range Mid) for top performers
- **Purpose**: Shows how we compensate high performers

### **New Hire CR**
- **Filter**: Status = "Approved" AND Start Date within last 365 days
- **Formula**: Average(Salary / Range Mid) for recent hires (last 365 days)
- **Purpose**: Shows new hire positioning

### **BT CR** (Below Talent Compa Ratio)
- **Filter**: Status = "Approved" AND AYR 2024 IN ("ML", "NI")
- **Formula**: Average(Salary / Range Mid) for underperformers
- **Purpose**: Shows compensation for needs improvement

---

## 🔧 Legacy Mapping Format

The `Legacy Mappings` sheet uses Aon level tokens:

### Level Token Format
- **P** = Professional (IC): `P2` → L2 IC, `P5` → L5 IC
- **M** = Manager: `M4` → L4 Mgr, `M5` → L5 Mgr
- **E** = Executive:
  - `E1` → L9 Mgr
  - `E2` → L8 Mgr
  - `E3` → L7 Mgr
  - `E4` → L6.5 Mgr
  - `E5` → L6 Mgr
  - `E6` → L5.5 Mgr

### Example Mappings
```
Employee ID | Job Family | Full Mapping  → Parsed As
20560       | EN.SODE    | EN.SODE.P5   → EN.SODE, L5 IC
20151       | EN.SODE    | EN.SODE.M4   → EN.SODE, L4 Mgr
102539      | EN.PGHC    | EN.PGHC.E3   → EN.PGHC, L7 Mgr
```

---

## 🚀 Performance

### **Before** (Custom Functions)
- Calculator recalculates on every change
- Each CR cell scans all employees
- Slow for large datasets (500+ employees)

### **After** (Pre-calculated + XLOOKUP)
- CR values calculated once during "Build Market Data"
- Stored in Full List sheet
- Calculators use instant XLOOKUP
- Fast even with 1000+ employees

---

## 📌 Best Practices

### 1. **Keep Legacy Mappings Updated**
- Add new employees to Legacy Mappings sheet if you have historical data
- Format: `EmpID | EN.SODE | EN.SODE.P5`

### 2. **Maintain Title Mapping**
- Map common job titles to Aon codes
- This auto-suggests mappings for new employees
- Update when new titles are added

### 3. **Regular Review Cycle**
- Run **Import Bob Data** weekly/monthly
- Review new "Needs Review" employees
- Approve/reject mappings promptly
- Rebuild Market Data after approvals

### 4. **Approval Workflow**
```
Import Bob Data
  ↓
Review Employee Mappings
  ↓
Approve good suggestions
  ↓
Manually fix rejected ones
  ↓
Build Market Data
  ↓
CR values updated in calculators
```

---

## 🐛 Troubleshooting

### **No mappings showing up**
- Check that Base Data was imported successfully
- Verify Legacy Mappings sheet has data
- Ensure Title Mapping sheet is populated

### **All mappings show "Needs Review"**
- This is normal for first-time setup
- Legacy Mappings and Title Mapping need to be populated
- Manually approve correct mappings

### **CR columns are empty**
- Ensure employees have Status = "Approved"
- Check that Performance Ratings sheet has AYR 2024 column
- Verify Start Date is populated in Base Data
- Rebuild Market Data after fixing

### **Confidence is always 0%**
- Employee not found in Legacy or Title mapping
- Add to Title Mapping sheet for future auto-suggestions

---

## 📖 Next Steps

1. **Populate Legacy Mappings**: Copy your provided data into the sheet
2. **Run Import Bob Data**: Let system suggest mappings
3. **Review & Approve**: Go through Employees Mapped sheet
4. **Build Market Data**: Generate Full List with CR values
5. **Use Calculators**: CR columns now populated with accurate data

---

**Version**: 3.8.0  
**Date**: November 27, 2025  
**Status**: ✅ Deployed

