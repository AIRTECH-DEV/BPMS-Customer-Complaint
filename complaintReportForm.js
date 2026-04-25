/**
 * Updates the Dispatch sheet for a specific Complaint ID
 * @param {string} reportId - The ID from the submitted form
 * @param {string} isResolved - The resolution status from the form
 */
function syncComplaintStatus(reportId, isResolved) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dispatchSheet = ss.getSheetByName("BPMS - Complaint Intake & Dispatch");
  const schedulingSheet = ss.getSheetByName("Complaint Scheduling"); // Access Scheduling sheet
  const targetSheet = ss.getSheetByName("Complaint_Close"); // Added target sheet

  // 1. Get all IDs from Column B of the Dispatch sheet
  const dispatchData = dispatchSheet.getRange("B:B").getValues();
  
  // 2. Find the row index where the ID matches
  const rowIndex = dispatchData.findIndex(row => row[0].toString().trim() === reportId.toString().trim());

  if (rowIndex !== -1) {
    const dispatchRow = rowIndex + 1; // Convert 0-index to 1-index row number
    
    // 3. Update the values for that specific row
    const today = new Date();
    dispatchSheet.getRange(dispatchRow, 17).setValue(today);          // Column Q: Date
    dispatchSheet.getRange(dispatchRow, 18).setValue("Done");           // Column R: Status
    
    // Column T: Yes/No based on the resolution status
    const resolvedStatus = (isResolved === "Yes") ? "Yes" : "No";
    dispatchSheet.getRange(dispatchRow, 20).setValue(resolvedStatus);
    
    // --- NEW LOGIC: Move row if Resolved is "Yes" ---
    if (resolvedStatus === "Yes") {
      // Check if it has already been copied to avoid duplicates (Column U / Index 20)
      const alreadyCopied = dispatchSheet.getRange(dispatchRow, 21).getValue();
      
      if (alreadyCopied !== "COPIED") {
        // Get the entire row data
        const rowData = dispatchSheet.getRange(dispatchRow, 1, 1, dispatchSheet.getLastColumn()).getValues()[0];
        
        // Append the row to the Complaint_Close sheet
        targetSheet.appendRow(rowData);
        
        // Mark Column U as "COPIED" in the source sheet
        dispatchSheet.getRange(dispatchRow, 21).setValue("COPIED");
        
        Logger.log("Row for ID " + reportId + " moved to Complaint_Close.");
      }
      // --- MOVE FROM SCHEDULING TO CALL SCHEDULING CLOSE ---
      // We look for the ID in Column B of the Scheduling sheet
      const schedulingData = schedulingSheet.getRange("B:B").getValues();
      const schedRowIndex = schedulingData.findIndex(row => row[0].toString().trim() === reportId.toString().trim());
      
      if (schedRowIndex !== -1) {
        const schedActualRow = schedRowIndex + 1;
        const schedTargetSheet = ss.getSheetByName("Call Scheduling Close"); // Ensure this sheet name matches exactly
        
        // 1. Get the data from the Scheduling row
        const schedRowData = schedulingSheet.getRange(schedActualRow, 1, 1, schedulingSheet.getLastColumn()).getValues()[0];
        
        // 2. Append to the Close sheet
        schedTargetSheet.appendRow(schedRowData);
        
        // 3. Delete from the active Scheduling sheet
        schedulingSheet.deleteRow(schedActualRow);
        
        Logger.log("Moved ID " + reportId + " from Complaint Scheduling to Call Scheduling Close.");
      }

    }
    // ------------------------------------------------
    
    Logger.log("Synced ID " + reportId + " to Dispatch Row " + dispatchRow);
  } else {
    Logger.log("Sync Failed: ID " + reportId + " not found in Dispatch sheet.");
  }
}

/*function handleComplaintReportForm(e) {
  const ss = e.source;
  const sheet = ss.getSheetByName("Complaint Report");
  const bpms  = ss.getSheetByName("BPMS - Complaint Intake & Dispatch");

  const row = e.range.getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const complaintId = values[1];  // column B of Complaint Report
  const timestamp = values[0];    // column A (Timestamp)

  if (!complaintId) return;

  const BPMS_HEADER_ROW = 6;
  const BPMS_ID_COL = 2;
  const BPMS_STATUS_COL = 24;  // X
  const BPMS_ACTUAL_COL = 22;  // Example: V

  const lastRow = bpms.getLastRow();
  const numRows = lastRow - BPMS_HEADER_ROW;
  if (numRows <= 0) return;

  const bpmsIds = bpms
      .getRange(BPMS_HEADER_ROW + 1, BPMS_ID_COL, numRows, 1)
      .getValues()
      .flat();

  const index = bpmsIds.indexOf(complaintId);
  if (index === -1) {
    Logger.log(`Complaint ID ${complaintId} not found in BPMS`);
    return;
  }

  const bpmsRow = BPMS_HEADER_ROW + 1 + index;

  // Set status to Done
  bpms.getRange(bpmsRow, BPMS_STATUS_COL).setValue("Done");

  // Set actual timestamp
  if (timestamp) {
    bpms.getRange(bpmsRow, BPMS_ACTUAL_COL).setValue(timestamp);
  }

  Logger.log(`Updated Complaint ID ${complaintId} → Done @ row ${bpmsRow}`);
}



function backfillComplaintReports() {
  const ss = SpreadsheetApp.getActive();
  const reports = ss.getSheetByName("Complaint Report");
  const bpms = ss.getSheetByName("BPMS - Complaint Intake & Dispatch");

  const REPORT_ID_COL = 2;   // Complaint ID in Complaint Report
  const REPORT_TS_COL = 1;   // Timestamp column A
  const BPMS_HEADER_ROW = 6;
  const BPMS_ID_COL = 2;     // Complaint ID in BPMS
  const BPMS_STATUS_COL = 24; // eXample Column X = Status
  const BPMS_ACTUAL_COL = 22; // e.g. Column V = Actual (adjust if needed)

  const lastReportRow = reports.getLastRow();
  if (lastReportRow < 2) return;

  const reportData = reports.getRange(2, 1, lastReportRow - 1, reports.getLastColumn()).getValues();

  const bpmsLast = bpms.getLastRow();
  const bpmsNumRows = bpmsLast - BPMS_HEADER_ROW;
  if (bpmsNumRows < 1) return;

  const bpmsIds = bpms.getRange(BPMS_HEADER_ROW + 1, BPMS_ID_COL, bpmsNumRows, 1).getValues().flat();

  reportData.forEach((row) => {
    const timestamp = row[REPORT_TS_COL - 1];
    const complaintId = row[REPORT_ID_COL - 1];

    if (!complaintId) return;

    const idx = bpmsIds.indexOf(complaintId);
    if (idx === -1) return;

    const bpmsRow = BPMS_HEADER_ROW + 1 + idx;

    // Mark Done
    bpms.getRange(bpmsRow, BPMS_STATUS_COL).setValue("Done");

    // Fill Actual timestamp
    if (timestamp) {
      bpms.getRange(bpmsRow, BPMS_ACTUAL_COL).setValue(timestamp);
    }
  });

  Logger.log("Backfill complete!");
}*/