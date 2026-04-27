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
