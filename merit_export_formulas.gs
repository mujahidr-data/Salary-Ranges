/**
 * NEW: Export Merit Data with XLOOKUP formulas (v5.0)
 * Uses formulas instead of calculated values for transparency and debugging
 */
function exportMeritDataWithFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '📊 Export Merit Data (Formula Version)',
    'This will create "Internal Data Analysis" with XLOOKUP formulas:\n\n' +
    '✅ Active employees only\n' +
    '✅ Formulas link to source sheets\n' +
    '✅ Easy to debug and verify\n' +
    '✅ Auto-updates when source data changes\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    SpreadsheetApp.getActive().toast('📊 Building formula-based export...', '', -1);
    
    // Get source sheets
    const baseData = ss.getSheetByName(SHEET_NAMES.BASE_DATA);
    const employeesMapped = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_MAPPED);
    const fullList = ss.getSheetByName(SHEET_NAMES.FULL_LIST);
    const bonusHistory = ss.getSheetByName(SHEET_NAMES.BONUS_HISTORY);
    const compHistorySummary = ss.getSheetByName(SHEET_NAMES.COMP_HISTORY_SUMMARY);
    const lookup = ss.getSheetByName(SHEET_NAMES.LOOKUP);
    
    if (!baseData || !employeesMapped || !fullList) {
      throw new Error('Missing required sheets: Base Data, Employees Mapped, or Full List');
    }
    
    // Load Base Data to filter active employees
    const baseVals = baseData.getDataRange().getValues();
    const baseHeaders = baseVals[0].map(h => String(h || '').trim());
    
    // Find Active column
    const bActive = baseHeaders.findIndex(h => /active.*inactive|status/i.test(h));
    const bEmpId = baseHeaders.findIndex(h => /emp.*id/i.test(h));
    
    if (bEmpId < 0 || bActive < 0) {
      throw new Error(`Cannot find Emp ID or Active column in Base Data`);
    }
    
    // Build list of active employee IDs
    const activeEmpIds = [];
    for (let r = 1; r < baseVals.length; r++) {
      const activeStatus = String(baseVals[r][bActive] || '').trim().toLowerCase();
      if (activeStatus && activeStatus !== 'inactive' && activeStatus !== 'terminated') {
        const empId = String(baseVals[r][bEmpId] || '').trim();
        if (empId) {
          activeEmpIds.push({empId, rowNum: r + 1}); // +1 for 1-based indexing
        }
      }
    }
    
    Logger.log(`Found ${activeEmpIds.length} active employees`);
    SpreadsheetApp.getActive().toast(`Building formulas for ${activeEmpIds.length} employees...`, '', -1);
    
    // Create/update output sheet
    let outputSheet = ss.getSheetByName('Internal Data Analysis');
    if (!outputSheet) {
      outputSheet = ss.insertSheet('Internal Data Analysis');
    }
    
    outputSheet.clear();
    outputSheet.setTabColor('#4285F4');
    
    // Write header
    const outputHeader = [
      'Emp ID', 'Emp Name', 'Start Date', 'Job Level', 'Title', 'Manager', 
      'Department', 'ELT', 'Site', 'Email', 'Base Salary', 'Base Salary in USD',
      'Aon Code (Full)', 'Applicable Internal Min', 'Applicable Internal Median', 
      'Applicable Internal Max', 'Applicable Market Range Start', 
      'Applicable Market Range Median', 'Applicable Market Range End',
      'Compa Ratio (Market Median)', 'Position in Range', 'Distance from Market Median',
      'Current Variable Type', 'Variable %', 'Last Increase Date', 
      'Last Promotion Date', 'Last Increase %'
    ];
    
    outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]);
    
    // Format header
    const headerRange = outputSheet.getRange(1, 1, 1, outputHeader.length);
    headerRange.setBackground('#4285F4')
               .setFontColor('#FFFFFF')
               .setFontWeight('bold')
               .setWrap(true);
    
    // Build formulas for each active employee
    const formulas = [];
    const bdSheet = `'${SHEET_NAMES.BASE_DATA}'`;
    const emSheet = `'${SHEET_NAMES.EMPLOYEES_MAPPED}'`;
    const flSheet = `'${SHEET_NAMES.FULL_LIST}'`;
    const bhSheet = `'${SHEET_NAMES.BONUS_HISTORY}'`;
    const csSheet = `'${SHEET_NAMES.COMP_HISTORY_SUMMARY}'`;
    const lkSheet = `'${SHEET_NAMES.LOOKUP}'`;
    
    for (let i = 0; i < activeEmpIds.length; i++) {
      const {empId, rowNum} = activeEmpIds[i];
      const outputRow = i + 2; // +2 because header is row 1
      const empIdCell = `A${outputRow}`;
      
      const rowFormulas = [
        // A: Emp ID (direct from Base Data)
        `=INDEX(${bdSheet}!A:A,${rowNum})`,
        
        // B: Emp Name
        `=INDEX(${bdSheet}!B:B,${rowNum})`,
        
        // C: Start Date
        `=INDEX(${bdSheet}!C:C,${rowNum})`,
        
        // D: Job Level (XLOOKUP from Employees Mapped)
        `=IFERROR(XLOOKUP(${empIdCell},${emSheet}!A:A,${emSheet}!H:H),"")`,
        
        // E: Title
        `=INDEX(${bdSheet}!D:D,${rowNum})`,
        
        // F: Manager
        `=INDEX(${bdSheet}!F:F,${rowNum})`,
        
        // G: Department
        `=INDEX(${bdSheet}!G:G,${rowNum})`,
        
        // H: ELT
        `=INDEX(${bdSheet}!H:H,${rowNum})`,
        
        // I: Site
        `=INDEX(${bdSheet}!I:I,${rowNum})`,
        
        // J: Email
        `=INDEX(${bdSheet}!J:J,${rowNum})`,
        
        // K: Base Salary
        `=INDEX(${bdSheet}!K:K,${rowNum})`,
        
        // L: Base Salary in USD (Base Salary * FX Rate)
        `=K${outputRow}*IFERROR(XLOOKUP(INDEX(${bdSheet}!L:L,${rowNum}),${lkSheet}!A:A,${lkSheet}!B:B),1)`,
        
        // M: Aon Code (Full) - from Employees Mapped
        `=IFERROR(XLOOKUP(${empIdCell},${emSheet}!A:A,${emSheet}!I:I),"")`,
        
        // N: Internal Min - from Full List (Site|Aon Code lookup)
        `=IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!O:O),"")`,
        
        // O: Internal Median
        `=IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!P:P),"")`,
        
        // P: Internal Max
        `=IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!Q:Q),"")`,
        
        // Q: Market Range Start (P40)
        `=IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!I:I),"")`,
        
        // R: Market Range Median (P62.5 for X0, P50 for Y1)
        `=IF(OR(LEFT(M${outputRow},3)="EN.",LEFT(M${outputRow},3)="DA.",M${outputRow}="TE.DADS"),IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!L:L),""),IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!K:K),""))`,
        
        // S: Market Range End (P75 for X0, P62.5 for Y1)
        `=IF(OR(LEFT(M${outputRow},3)="EN.",LEFT(M${outputRow},3)="DA.",M${outputRow}="TE.DADS"),IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!M:M),""),IFERROR(XLOOKUP(I${outputRow}&"|"&M${outputRow},${flSheet}!A:A&"|"&${flSheet}!C:C,${flSheet}!L:L),""))`,
        
        // T: Compa Ratio
        `=IF(AND(L${outputRow}>0,R${outputRow}>0),L${outputRow}/R${outputRow},"")`,
        
        // U: Position in Range
        `=IF(AND(L${outputRow}>0,Q${outputRow}>0,S${outputRow}>Q${outputRow}),(L${outputRow}-Q${outputRow})/(S${outputRow}-Q${outputRow}),"")`,
        
        // V: Distance from Market Median
        `=IF(AND(L${outputRow}>0,R${outputRow}>0),L${outputRow}-R${outputRow},"")`,
        
        // W: Current Variable Type
        `=IFERROR(XLOOKUP(${empIdCell},${bhSheet}!A:A,${bhSheet}!D:D),"")`,
        
        // X: Variable %
        `=IFERROR(XLOOKUP(${empIdCell},${bhSheet}!A:A,${bhSheet}!E:E),"")`,
        
        // Y: Last Increase Date
        `=IFERROR(XLOOKUP(${empIdCell},${csSheet}!A:A,${csSheet}!C:C),"")`,
        
        // Z: Last Promotion Date
        `=IFERROR(XLOOKUP(${empIdCell},${csSheet}!A:A,${csSheet}!B:B),"")`,
        
        // AA: Last Increase %
        `=IFERROR(XLOOKUP(${empIdCell},${csSheet}!A:A,${csSheet}!D:D),"")`
      ];
      
      formulas.push(rowFormulas);
      
      if ((i + 1) % 100 === 0) {
        SpreadsheetApp.getActive().toast(`Built formulas for ${i + 1} of ${activeEmpIds.length} employees...`, '', -1);
      }
    }
    
    // Write all formulas at once
    SpreadsheetApp.getActive().toast('Writing formulas to sheet...', '', -1);
    if (formulas.length > 0) {
      outputSheet.getRange(2, 1, formulas.length, outputHeader.length).setFormulas(formulas);
    }
    
    // Format columns
    SpreadsheetApp.getActive().toast('Formatting...', '', -1);
    if (formulas.length > 0) {
      const numRows = formulas.length;
      
      // Date formats
      outputSheet.getRange(2, 3, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Start Date
      outputSheet.getRange(2, 25, numRows, 2).setNumberFormat('yyyy-mm-dd'); // Last Increase/Promotion Dates
      
      // Number formats
      outputSheet.getRange(2, 11, numRows, 2).setNumberFormat('#,##0'); // Base Salaries
      outputSheet.getRange(2, 14, numRows, 3).setNumberFormat('#,##0'); // Internal Min/Med/Max
      outputSheet.getRange(2, 17, numRows, 3).setNumberFormat('#,##0'); // Market ranges
      outputSheet.getRange(2, 20, numRows, 1).setNumberFormat('0.00'); // Compa Ratio
      outputSheet.getRange(2, 21, numRows, 1).setNumberFormat('0.0%'); // Position in Range
      outputSheet.getRange(2, 22, numRows, 1).setNumberFormat('#,##0'); // Distance from median
    }
    
    // Auto-resize columns
    SpreadsheetApp.getActive().toast('Auto-resizing columns...', '', -1);
    outputSheet.autoResizeColumns(1, outputHeader.length);
    
    // Freeze header row
    outputSheet.setFrozenRows(1);
    
    // Sort by Start Date
    if (formulas.length > 1) {
      outputSheet.getRange(2, 1, formulas.length, outputHeader.length).sort(3); // Sort by column C (Start Date)
    }
    
    SpreadsheetApp.getActive().toast('', '', 1); // Clear toast
    
    ui.alert(
      '✅ Merit Export Complete (Formula Version)',
      `Successfully exported ${formulas.length} active employees.\n\n` +
      'Sheet: "Internal Data Analysis"\n\n' +
      'All columns use XLOOKUP formulas:\n' +
      '  ✅ Easy to debug\n' +
      '  ✅ Auto-updates with source data\n' +
      '  ✅ Transparent calculations\n\n' +
      'Sorted by Start Date (oldest first).',
      ui.ButtonSet.OK
    );
    
    Logger.log(`Formula-based merit export complete: ${formulas.length} employees`);
    
  } catch (error) {
    SpreadsheetApp.getActive().toast('', '', 1);
    Logger.log(`Error in exportMeritDataWithFormulas: ${error}`);
    ui.alert('❌ Error', `Failed to export merit data:\n\n${error.message}`, ui.ButtonSet.OK);
    throw error;
  }
}

