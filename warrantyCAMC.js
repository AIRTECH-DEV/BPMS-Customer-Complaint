function generateWarrantyCAMC() {
  const ss = SpreadsheetApp.getActive();
  const report = ss.getSheetByName("Complaint Report1");
  const cleanedData = ss.getSheetByName("Cleaned Data");
  const warrantySheet = ss.getSheetByName("BPMS - Warranty / CAMC");

  if (!report || !cleanedData || !warrantySheet) return;

  // 1. GET THE LATEST SUBMISSION ONLY
  const lastRowReport = report.getLastRow();
  if (lastRowReport < 2) return;
  
  const lastEntry = report.getRange(lastRowReport, 1, 1, report.getLastColumn()).getValues()[0];
  const service = (lastEntry[14] || "").toString().toUpperCase().trim();
  const cid = lastEntry[1] ? lastEntry[1].toString().trim() : "";

  if (!service.includes("WARRANTY") || cid === "") return;

  // 2. LOOKUP COMPANY/CLIENT INFO
  const cleanVals = cleanedData.getDataRange().getValues();
  let info = { company: "Not Found", client: "Not Found" };
  for (let i = 0; i < cleanVals.length; i++) {
    if (cleanVals[i][1].toString().trim() === cid) {
      info.company = cleanVals[i][2];
      info.client = cleanVals[i][3];
      break;
    }
  }

  // 3. REMOVE PREVIOUS ENTRY (IF EXISTS)
  // We find the real last row with data to avoid scanning empty rows
  const realLastRow = getRealLastRow(warrantySheet, "B"); 
  if (realLastRow >= 7) {
    const existingCids = warrantySheet.getRange(7, 2, realLastRow - 6, 1).getValues();
    for (let i = existingCids.length - 1; i >= 0; i--) {
      if (existingCids[i][0].toString().trim() === cid) {
        warrantySheet.deleteRow(i + 7);
        Logger.log("Deleted old entry for CID: " + cid);
      }
    }
  }

  // 4. APPEND TO NEXT AVAILABLE ROW
  const rowData = [[
    lastEntry[0], // A: Timestamp
    cid,          // B: Complaint ID
    info.company, // C: Company Name
    info.client,  // D: Contact Person
    lastEntry[14],// E: Service Type
    lastEntry[12]  // F: Issue
  ]];

  // Recalculate last row after deletion to find the "Next to last" spot
  const nextRow = Math.max(getRealLastRow(warrantySheet, "B") + 1, 7);
  
  warrantySheet.getRange(nextRow, 1, 1, 6).setValues(rowData);
  
  SpreadsheetApp.flush();
  Logger.log("Moved CID " + cid + " to row " + nextRow);
}

/**
 * Helper function to find the actual last row with content in a specific column
 */
function getRealLastRow(sheet, columnAlphabet) {
  const range = sheet.getRange(columnAlphabet + ":" + columnAlphabet);
  const values = range.getValues();
  let lastRow = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }
  return lastRow;
}
