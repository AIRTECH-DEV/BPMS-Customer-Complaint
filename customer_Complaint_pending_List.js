/**
 * Automatically adds the complaint to the Pending List if not resolved.
 * @param {string} complaintId - The ID generated/matched during submission.
 * @param {object} formData - The data object from the form.
 */
function syncPendingComplaints(complaintId, formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pendingSheet = ss.getSheetByName("Customer_Complaint_Pending_List");
  
  // 1. Prevent Duplicates: Check if this ID is already in the Pending List
  const lastRowPending = pendingSheet.getLastRow();
  if (lastRowPending >= 2) {
    const pendingIds = pendingSheet.getRange(2, 2, lastRowPending - 1, 1).getValues().flat();
    if (pendingIds.includes(complaintId)) {
      Logger.log("ID: " + complaintId + " is already in Pending List. Skipping.");
      return; 
    }
  }

  // 2. Append the new row to the Pending List
  // Using formData directly ensures we don't have to wait for the sheet to update
  pendingSheet.appendRow([
    new Date(),             // Timestamp
    complaintId,            // Complaint ID
    formData.company,       // Company / Client Name
    formData.contact,       // Contact Person (Adjust based on your mapping)
    formData.serviceType,   // Service Type
    formData.model,         // Machine Model
    "No"                    // Status (Is it Resolved? No)
  ]);

  Logger.log("Successfully added ID " + complaintId + " to Customer_Complaint_Pending_List");
}



/**
 * Removes a complaint from the Pending List once it is resolved.
 * @param {string} complaintId - The ID to be removed.
 */
function removeFromPendingList(complaintId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pendingSheet = ss.getSheetByName("Customer_Complaint_Pending_List");
  const lastRow = pendingSheet.getLastRow();
  
  if (lastRow < 2) return; // Sheet is empty

  // 1. Get all IDs from Column B (index 1)
  const data = pendingSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  
  // 2. Loop backwards through the data (standard practice when deleting rows)
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0].toString().trim() === complaintId.toString().trim()) {
      // Row index is i + 2 because data starts at row 2 and is 0-indexed
      pendingSheet.deleteRow(i + 2);
      Logger.log("Removed ID " + complaintId + " from Pending List.");
    }
  }
}