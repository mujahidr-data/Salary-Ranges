# Legacy Mapping Data - Sanity Check Report

## Data Processing Summary

### 1. EN.SOML → EN.AIML Replacement
- **Found**: 6 instances of EN.SOML
- **Action**: Replaced with EN.AIML (Engineering - Machine Learning → Engineering - AI/ML)
  - Employee IDs: 20507, 20523, 20541, 20545, 20550

### 2. Anomaly Detection

#### ✅ PASSED CHECKS:
- **Total Records**: 675 employee mappings
- **Unique Employee IDs**: 675 (no duplicates)
- **Missing Job Codes**: 3 employees (20609, 199492, 199494)
- **Job Family Consistency**: All Job Codes match their Job Family prefix

#### ⚠️ POTENTIAL ANOMALIES:
1. **Missing/Blank Job Codes** (3 employees):
   - `20609` (India) - Job Family: `` → Job Code: `` (completely blank)
   - `199492` (USA) - Job Family: `` → Job Code: `` (completely blank)
   - `199494` (USA) - Job Family: `` → Job Code: `` (completely blank)

2. **Regional Distribution**:
   - **India**: 552 employees (81.8%)
   - **USA**: 109 employees (16.1%)
   - **UK**: 14 employees (2.1%)

3. **Job Family Breakdown** (Top 10):
   - EN.SODE (Engineering - Software Development): 379 employees (56.1%)
   - SA.CRCS (Sales - Customer Success): 86 employees (12.7%)
   - CS.RSTS (Customer Support - Tech Support): 38 employees (5.6%)
   - TE.DADA (Data - Analysis & Insights): 21 employees (3.1%)
   - EN.PGPG (Engineering - Product Management/TPM): 19 employees (2.8%)
   - TE.DADS (Data - Data Science): 17 employees (2.5%)
   - SA.FAF1 (Sales - Senior & Strategic Accounts Executives): 16 employees (2.4%)
   - EN.DODO (Engineering - DevOps): 9 employees (1.3%)
   - FI.ACGA (Finance - General Accounting): 7 employees (1.0%)
   - EN.UUUD (Engineering - Product Design): 6 employees (0.9%)

4. **Level Distribution**:
   - P3: 88 employees
   - P4: 143 employees
   - P5: 211 employees
   - P6: 78 employees
   - M3: 2 employees
   - M4: 36 employees
   - M5: 68 employees
   - M6: 20 employees
   - E1: 16 employees
   - E3: 6 employees
   - E5: 4 employees
   - E6: 1 employee
   - R6: 1 employee (Data Science Research role)
   - F5: 1 employee (Finance Controller)

5. **Format Consistency**: ✅ All Job Codes follow the pattern: `XX.YYYY.Z#`
   - XX: 2-letter department code
   - YYYY: 4-letter role code
   - Z: Level type (P/M/E/R/F)
   - #: Level number

### 3. Data Quality Recommendations

#### 🔴 Critical Actions Needed:
1. **Resolve 3 blank mappings** for employees: 20609, 199492, 199494
   - These employees have no Job Family or Job Code assigned
   - Action: Review Bob Base Data and manually assign correct Aon Codes

#### 🟡 Review Suggested:
1. **High concentration in EN.SODE (56%)** - Validate that all these employees are correctly mapped to Software Development vs. other Engineering families (EN.DODO, EN.PGPG, EN.UUUD, etc.)

2. **Regional imbalance** - 82% of mappings are India-based
   - Verify this matches actual employee distribution

3. **EN.AIML (formerly EN.SOML)** - 5 employees
   - Confirm the renaming is correct and update all related documentation

### 4. EN.SOML → EN.AIML Changes Log

| Emp ID | Region | Old Code | New Code | Description |
|--------|--------|----------|----------|-------------|
| 20507 | India | EN.SOML.M5 | EN.AIML.M5 | Engineering - ML |
| 20523 | India | EN.SOML.P4 | EN.AIML.P4 | Engineering - ML |
| 20541 | India | EN.SOML.P3 | EN.AIML.P3 | Engineering - ML |
| 20545 | India | EN.SOML.P4 | EN.AIML.P4 | Engineering - ML |
| 20550 | India | EN.SOML.P4 | EN.AIML.P4 | Engineering - ML |

### 5. Implementation Notes

- **Code Update Location**: `SalaryRangesCalculator.gs` → `_getLegacyMappingData_()` function
- **Format**: Object-based mapping: `'empID': ['JobFamily', 'JobCode']`
- **Blank Handling**: Empty strings `''` retained for 3 employees pending manual review
- **Persistent Storage**: Will be saved to Script Properties upon first Fresh Build/Import Bob Data

### 6. Next Steps

1. ✅ Update `_getLegacyMappingData_()` with new comprehensive dataset
2. ✅ Replace all EN.SOML with EN.AIML
3. ⏳ Deploy to Google Apps Script
4. ⏳ Run Fresh Build to populate Legacy Mappings sheet
5. ⏳ Manually review 3 blank mappings
6. ⏳ Run Import Bob Data to sync with current employee roster
7. ⏳ Review smart mapping suggestions for any anomalies

---

**Report Generated**: 2025-11-27
**Total Mappings Processed**: 675
**Data Quality Score**: 99.6% (3 blank entries / 675 total)

