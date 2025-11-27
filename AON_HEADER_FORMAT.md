# Aon Data Header Format

## Standard Headers (Consistent Across All Regions)

The Aon data sheets (India, US, UK) now use a **standardized header format** across all regions:

### Column Structure:

| Column | Header Name |
|--------|-------------|
| A | `Job Code` |
| B | `Job Family` |
| C | `Market \n\n (43) CFY Fixed Pay: 10th Percentile` |
| D | `Market \n\n (43) CFY Fixed Pay: 25th Percentile` |
| E | `Market \n\n (43) CFY Fixed Pay: 40th Percentile` |
| F | `Market \n\n (43) CFY Fixed Pay: 50th Percentile` |
| G | `Market \n\n (43) CFY Fixed Pay: 62.5th Percentile` |
| H | `Market \n\n (43) CFY Fixed Pay: 75th Percentile` |
| I | `Market \n\n (43) CFY Fixed Pay: 90th Percentile` |

### Notes on Header Format:

1. **Newlines in Headers**: The percentile column headers contain newlines (`\n`). When displayed in Google Sheets, they appear as multi-line headers:
   ```
   Market
   
    (43) CFY Fixed Pay: 10th Percentile
   ```

2. **Regex Pattern**: The code uses flexible regex patterns to match these headers:
   ```javascript
   'Market[\\s\\n]*(\\(43\\))?[\\s\\n]*CFY[\\s\\n]*Fixed[\\s\\n]*Pay:[\\s\\n]*10(?:th)?[\\s\\n]*Percentile'
   ```
   - `[\\s\\n]*` matches any whitespace including newlines
   - `(\\(43\\))?` makes the "(43)" optional
   - `(?:th)?` makes the "th" suffix optional

3. **Case Insensitive**: All header matching is case-insensitive

## How It Works

### 1. Job Code Column
- Contains Aon job codes (e.g., "1234.P6", "5678.M4")
- Format: `BaseCode.LevelLetter+Number`
- Example: "1234.P6" = Base code 1234, Professional level 6

### 2. Job Family Column
- Contains executive job family descriptions
- Populated automatically by the system from the Lookup sheet
- Maps Aon codes to friendly names (e.g., "Engineering - Software Development")

### 3. Percentile Columns (P10, P25, P40, P50, P62.5, P75, P90)
- Market salary data at various percentiles
- All values should be numeric (no commas or currency symbols)
- Empty cells are handled gracefully with fallback logic

## Importing Aon Data

When pasting Aon data into the sheets:

1. **Preserve Headers**: Keep the exact header format (including newlines)
2. **Paste Data**: Paste your Aon data starting from row 2
3. **Job Family Column**: Will be auto-populated when you run "Import Bob Data"
4. **Numeric Format**: Ensure percentile values are numbers (no text or currency formatting)

## Validation

The system validates that:
- `Job Code` column exists
- `Job Family` column exists (created if missing)
- All 7 percentile columns are present and parseable

If headers don't match, the system will show an error with available column names.

## Region-Specific Sheets

Each region has its own Aon sheet with identical structure:

- **Aon US - 2025**: US market data
- **Aon UK - 2025**: UK market data  
- **Aon India - 2025**: India market data

All three sheets use the same header format for consistency.

## Fallback Logic

If a required percentile is missing for a job code:

### X0 (Engineering & Product)
- Range Start (P25) missing → use P40 → P50
- Range Mid (P62.5) missing → use P75 → P90
- Range End (P90) missing → no fallback

### Y1 (Everyone Else)
- Range Start (P10) missing → use P25 → P40
- Range Mid (P40) missing → use P50 → P62.5
- Range End (P62.5) missing → use P75 → P90

## Troubleshooting

### "Missing Job Family/Job Code/Percentile header" Error

**Cause**: Headers don't match expected format

**Solution**:
1. Check that row 1 has the exact headers listed above
2. Verify newlines are preserved in percentile headers
3. Run "Fresh Build" to recreate placeholder sheets with correct headers
4. Copy data from old sheets to new sheets

### Percentile Values Show as #REF!

**Cause**: Header regex not matching due to format differences

**Solution**:
1. Check that percentile headers match the expected format
2. Verify no extra spaces or characters
3. Check that "(43)" is present in the header
4. Ensure "CFY Fixed Pay" is spelled correctly

### Job Family Column Missing

**Cause**: Column not created yet

**Solution**:
1. Run "Import Bob Data" - this auto-creates the Job Family column
2. OR manually insert column B and name it "Job Family"

---

**Last Updated**: v3.7.1  
**Date**: November 27, 2025

