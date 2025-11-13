/********************************
 * Combined Imports: Bob, Bonus, Comp
 * Relies on helpers from Shared_Helpers.js
 ********************************/

function importBobDataSimpleWithLookup() {
  const reportId = "31048356";
  const sheetName = "Base Data";
  const bonusSheetName = "Bonus History";
  const apiUrl = `https://api.hibob.com/v1/company/reports/${reportId}/download?format=csv&locale=en-CA`;

  const bobId = PropertiesService.getScriptProperties().getProperty("BOB_ID");
  const bobKey = PropertiesService.getScriptProperties().getProperty("BOB_KEY");
  if (!bobId || !bobKey) throw new Error("Missing BOB_ID / BOB_KEY.");
  const auth = Utilities.base64Encode(`${bobId}:${bobKey}`);

  const res = UrlFetchApp.fetch(apiUrl, {
    method: "GET",
    headers: { Authorization: `Basic ${auth}`, accept: "text/csv" },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch CSV: ${res.getResponseCode()} - ${res.getContentText()}`);
  }

  const rows = Utilities.parseCsv(res.getContentText());
  if (!rows.length) throw new Error("CSV contains no data.");

  const srcHeader = rows[0];
  const idxEmpId       = findCol(srcHeader, ["Employee ID", "Emp ID", "Employee Id"]);
  const idxJobLevel    = findCol(srcHeader, ["Job Level", "Job level"]);
  const idxBasePay     = findCol(srcHeader, ["Base Pay", "Base salary", "Base Salary"]);
  const idxEmpType     = findCol(srcHeader, ["Employment Type", "Employment type"]);
  const idxStartDate   = findCol(srcHeader, ["Start Date", "Start date", "Original start date", "Original Start Date"]);
  const idxTermination = findCol(srcHeader, ["Termination Date", "Termination date"]);
  const idxTitleOpt    = findColOptional(srcHeader, ["Job title", "Job Title", "Title", "Job name"]);

  // ===== BUILD BONUS LOOKUP MAP =====
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bonusSheet = ss.getSheetByName(bonusSheetName);
  const bonusMap = new Map(); // empId -> {type, percent}
  if (bonusSheet) {
    const bonusData = bonusSheet.getDataRange().getValues();
    if (bonusData.length > 1) {
      const bonusHdr = bonusData[0];
      const bIdxEmpId = bonusHdr.indexOf("Employee ID");
      const bIdxType  = bonusHdr.indexOf("Variable type");
      const bIdxPct   = bonusHdr.indexOf("Commission/Bonus %");
      if (bIdxEmpId >= 0 && bIdxType >= 0 && bIdxPct >= 0) {
        for (let r = 1; r < bonusData.length; r++) {
          const empId = String(bonusData[r][bIdxEmpId] || '').trim();
          if (empId) {
            bonusMap.set(empId, { type: bonusData[r][bIdxType] || '', percent: bonusData[r][bIdxPct] || '' });
          }
        }
      }
    }
  }

  // ===== Build Title→Aon Code and Code→Exec Description maps =====
  const titleMap = buildTitleToFamilyMap_(ss);
  const codeToExec = buildCodeToExecDescMap_(ss);

  // ===== PROCESS MAIN DATA =====
  let header = srcHeader.slice();
  header = [...header, "Variable Type", "Variable %", "Job Family Name", "Mapped Family"];

  const allowedEmpTypes = new Set(["Permanent", "Regular Full-Time"]);
  const out = [header];

  for (let r = 1; r < rows.length; r++) {
    const src = rows[r];
    if (!src || src.length === 0) continue;
    const row = src.slice();
    const empType = safeCell(row, idxEmpType);
    if (!allowedEmpTypes.has(empType)) continue;

    const empId  = safeCell(row, idxEmpId);
    const jobLvl = safeCell(row, idxJobLevel);
    if (!empId || !jobLvl) continue;

    const basePayNum = toNumberSafe(safeCell(row, idxBasePay));
    if (!isFinite(basePayNum) || basePayNum === 0) continue;
    row[idxBasePay] = basePayNum;

    const bonus = bonusMap.get(empId);
    row.push(bonus ? bonus.type : "", bonus ? bonus.percent : "");

    // Auto-map Aon Job Family code from Title Mapping
    const title = idxTitleOpt >= 0 ? safeCell(row, idxTitleOpt) : "";
    const norm = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    const titleNorm = norm(title);
    let code = titleMap.get(title) || titleMap.get(titleNorm) || "";
    if (!code && titleNorm) {
      for (const [k, v] of titleMap.entries()) {
        const kn = norm(k);
        if (kn && (titleNorm.includes(kn) || kn.includes(titleNorm))) { code = v; break; }
      }
    }
    const exec = code ? (codeToExec.get(code) || code) : "";
    row.push(code, exec);
    out.push(row);
  }

  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const WRITE_COLS = out[0].length;
  const sliced = out.map(r => r.slice(0, WRITE_COLS));
  const rowsToClear = Math.max(sheet.getMaxRows(), sliced.length);
  sheet.getRange(1, 1, rowsToClear, WRITE_COLS).clearContent();
  sheet.getRange(1, 1, sliced.length, sliced[0].length).setValues(sliced);

  const numRows = sliced.length - 1;
  const firstDataRow = 2;
  if (numRows > 0) {
    if (idxBasePay + 1 <= WRITE_COLS) sheet.getRange(firstDataRow, idxBasePay + 1, numRows, 1).setNumberFormat("#,##0.00");
    const vpCol = header.indexOf("Variable %") + 1;
    if (vpCol > 0 && vpCol <= WRITE_COLS) sheet.getRange(firstDataRow, vpCol, numRows, 1).setNumberFormat("0.########");
  }
  Logger.log(`Imported ${numRows} employees successfully with bonus data + family mapping`);
}

// ===== Bonus Import =====
function importBobBonusHistoryLatest() {
  const reportId = "31054302";
  const targetSheetName = "Bonus History";
  const rows = fetchBobReport(reportId);
  const header = rows[0];
  const iEmpId   = findCol(header, ["Employee ID", "Emp ID", "Employee Id"]);
  const iName    = findCol(header, ["Display name", "Emp Name", "Display Name", "Name"]);
  const iEffDate = findCol(header, ["Effective date", "Effective Date", "Effective"]);
  const iType    = findCol(header, ["Variable type", "Variable Type", "Type"]);
  const iPct     = findCol(header, ["Commission/Bonus %", "Variable %", "Commission %", "Bonus %"]);
  const iAmt     = findCol(header, ["Amount", "Variable Amount", "Commission/Bonus Amount"]);
  const iCurr    = findColOptional(header, [
    "Variable Amount currency", "Variable Amount Currency",
    "Amount currency", "Amount Currency",
    "Currency", "Currency code", "Currency Code"
  ]);
  const latest = new Map();
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]; if (!row || row.length === 0) continue;
    const empId = safeCell(row, iEmpId);
    const effRaw = safeCell(row, iEffDate);
    const effKey = (effRaw.match(/^\d{4}-\d{2}-\d{2}/) || [])[0];
    if (!empId || !effKey) continue;
    const existing = latest.get(empId);
    if (!existing || effKey > existing.effKey) latest.set(empId, { row, effKey });
  }
  const outHeader = ["Employee ID","Display name","Effective date","Variable type","Commission/Bonus %","Amount","Amount currency"];
  const out = [outHeader];
  latest.forEach(({ row, effKey }) => {
    const empId = safeCell(row, iEmpId);
    const name  = safeCell(row, iName);
    const type  = safeCell(row, iType);
    const pctVal = toNumberSafe(safeCell(row, iPct));
    const amtVal = toNumberSafe(safeCell(row, iAmt));
    const curr   = iCurr === -1 ? "" : safeCell(row, iCurr);
    out.push([empId, name, effKey, type, isFinite(pctVal) ? pctVal : "", isFinite(amtVal) ? amtVal : "", curr]);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, targetSheetName);
  writeSheetData(sheet, out, { 3: "@", 5: "0.########", 6: "#,##0.00" });
  const recordCount = out.length - 1;
  Logger.log(`Imported ${recordCount} bonus records (latest per employee)`);
}

// ===== Comp Import =====
function importBobCompHistoryLatest() {
  const reportId = "31054312"; // Comp History report
  const targetSheetName = "Comp History";
  const rows = fetchBobReport(reportId);
  const header = rows[0];
  const iEmpId   = findCol(header, ["Emp ID", "Employee ID", "Employee Id"]);
  const iName    = findCol(header, ["Emp Name", "Display name", "Display Name", "Name"]);
  const iEffDate = findCol(header, ["History effective date", "Effective date", "Effective Date", "Effective"]);
  const iBase    = findCol(header, ["History base salary", "Base salary", "Base Salary", "Base pay", "Salary"]);
  const iCurr    = findCol(header, ["History base salary currency", "Base salary currency", "Currency"]);
  const iReason  = findCol(header, ["History reason", "Reason", "Change reason"]);
  const latest = new Map();
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]; if (!row || row.length === 0) continue;
    const empId = safeCell(row, iEmpId);
    const effStr = safeCell(row, iEffDate);
    const ymd = toYmd(effStr);
    if (!empId || !ymd) continue;
    const existing = latest.get(empId);
    if (!existing || ymd > existing.ymd) latest.set(empId, { row, ymd });
  }
  const outHeader = ["Emp ID","Emp Name","Effective date","Base salary","Base salary currency","History reason"];
  const out = [outHeader];
  latest.forEach(({ row, ymd }) => {
    const empId  = safeCell(row, iEmpId);
    const name   = safeCell(row, iName);
    const base   = toNumberSafe(safeCell(row, iBase));
    const curr   = safeCell(row, iCurr);
    const reason = safeCell(row, iReason);
    const effDate = parseDateSmart(ymd);
    out.push([empId, name, effDate, isFinite(base) ? base : "", curr, reason]);
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, targetSheetName);
  writeSheetData(sheet, out, { 3: "yyyy-mm-dd", 4: "#,##0.00" });
  const recordCount = out.length - 1;
  Logger.log(`Imported ${recordCount} comp history records (latest per employee)`);
}

// ===== Mapping helpers =====
function buildTitleToFamilyMap_(ss) {
  const sh = ss.getSheetByName('Title Mapping');
  const map = new Map();
  if (!sh) return map;
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return map;
  const head = vals[0].map(h => String(h || '').trim());
  const iLive   = findColOptional(head, ['Job title (live)', 'Job Title (Live)', 'Job title live']);
  const iMapped = findColOptional(head, ['Job title (Mapped)', 'Job title mapped', 'Mapped Title']);
  const iFam    = findColOptional(head, ['Job family', 'Job Family', 'Aon Code', 'Job Code']);
  const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  for (let r = 1; r < vals.length; r++) {
    const live = iLive >= 0 ? String(vals[r][iLive] || '').trim() : '';
    const mapped = iMapped >= 0 ? String(vals[r][iMapped] || '').trim() : '';
    const fam = iFam >= 0 ? String(vals[r][iFam] || '').trim() : '';
    if (!fam) continue;
    if (live) { map.set(live, fam); map.set(norm(live), fam); }
    if (mapped) { map.set(mapped, fam); map.set(norm(mapped), fam); }
  }
  return map;
}

function buildCodeToExecDescMap_(ss) {
  const sh = ss.getSheetByName('Job family Descriptions');
  const map = new Map();
  if (!sh) return map;
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return map;
  const head = vals[0].map(h => String(h || '').trim());
  const iCode = findColOptional(head, ['Aon Code', 'Job Code', 'Aon code', 'Job family code']);
  const iDesc = findColOptional(head, ['Job Family (Exec Description)', 'Exec Description', 'Description']);
  for (let r = 1; r < vals.length; r++) {
    const code = iCode >= 0 ? String(vals[r][iCode] || '').trim() : '';
    const desc = iDesc >= 0 ? String(vals[r][iDesc] || '').trim() : '';
    if (code) map.set(code, desc || code);
  }
  return map;
}
