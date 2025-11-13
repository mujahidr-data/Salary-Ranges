#!/bin/bash
# Build consolidated salary ranges script

OUTPUT="SalaryRangesCalculator.gs"

# Start with header
cat > "$OUTPUT" << 'HEADER'
/**
 * Salary Ranges Calculator - Consolidated Google Apps Script
 * 
 * Combines HiBob employee data with Aon market data for comprehensive
 * salary range analysis and calculation.
 * 
 * Features:
 * - Bob API integration (Base Data, Bonus, Comp History)
 * - Aon market percentiles (P40, P50, P62.5, P75, P90)
 * - Multi-region support (US, UK, India) with FX conversion
 * - Salary range categories (X0, X1, Y1)
 * - Internal vs Market analytics
 * - Job family and title mapping
 * - Interactive calculator UI
 * 
 * @version 3.0.0
 * @date 2025-11-13
 * 
 * Aon Data Source: https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const BOB_REPORT_IDS = {
  BASE_DATA: "31048356",
  BONUS_HISTORY: "31054302",
  COMP_HISTORY: "31054312"
};

const SHEET_NAMES = {
  BASE_DATA: "Base Data",
  BONUS_HISTORY: "Bonus History",
  COMP_HISTORY: "Comp History",
  SALARY_RANGES: "Salary Ranges",
  FULL_LIST: "Full List",
  FULL_LIST_USD: "Full List USD",
  LOOKUP: "Lookup"
};

const REGION_TAB = {
  'India': 'Aon India - 2025',
  'US': 'Aon US Premium - 2025',
  'UK': 'Aon UK London - 2025'
};

const CACHE_TTL = 600; // 10 minutes
const ALLOWED_EMP_TYPES = new Set(["Permanent", "Regular Full-Time"]);

HEADER

# Add Helpers first (foundational functions)
echo "" >> "$OUTPUT"
echo "// ============================================================================" >> "$OUTPUT"
echo "// HELPER FUNCTIONS" >> "$OUTPUT"
echo "// ============================================================================" >> "$OUTPUT"
echo "" >> "$OUTPUT"
tail -n +7 Helpers.gs >> "$OUTPUT"

# Add AppImports (Bob data integration)
echo "" >> "$OUTPUT"
echo "// ============================================================================" >> "$OUTPUT"
echo "// BOB DATA IMPORTS" >> "$OUTPUT"
echo "// ============================================================================" >> "$OUTPUT"
echo "" >> "$OUTPUT"
tail -n +6 AppImports.gs | head -n -1 >> "$OUTPUT"

# Add RangeCalculator (main logic)
echo "" >> "$OUTPUT"
echo "// ============================================================================" >> "$OUTPUT"
echo "// SALARY RANGE CALCULATIONS" >> "$OUTPUT"
echo "// ============================================================================" >> "$OUTPUT"
echo "" >> "$OUTPUT"
tail -n +24 RangeCalculator.gs >> "$OUTPUT"

echo "✅ Created consolidated script: $OUTPUT"
wc -l "$OUTPUT"
