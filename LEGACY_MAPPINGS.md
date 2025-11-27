# Legacy Employee Mappings

## Purpose
This file stores the legacy employee mappings provided by the user for reference and validation.

## Usage
The system will:
1. Load these mappings on first sync
2. Use them as a baseline for existing employees
3. Flag any discrepancies for review
4. Suggest mappings for new employees based on Title Mapping

## Format
Two mapping types:
- **Job Family Only**: EmpID → Base Aon Code (e.g., "EN.SODE")
- **Full Mapping**: EmpID → Full Aon Code with Level (e.g., "EN.SODE.P5")

## Instructions
1. Copy this data to the "Lookup" sheet under "Legacy Employee Mappings" section
2. Run "Import Bob Data" to sync with current employees
3. Review suggested mappings in "Employees Mapped" sheet
4. Approve or modify mappings as needed

---

## Implementation Notes
- The syncEmployeesMappedSheet_() function will reference these mappings
- New employees will get suggested mappings based on Title Mapping
- Status column will show: "Legacy", "Title-Based", "Needs Review", "Approved"
- CR calculations will use Performance Ratings and Start Date filters

