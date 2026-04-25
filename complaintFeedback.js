function processFeedbackFromReport(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = e.range.getSheet();
  
  // 1. Security: Only run if the edit is in "Complaint Report1"
  if (sourceSheet.getName() !== "Complaint Report1") return;

  const feedbackSheet = ss.getSheetByName("BPMS - Customer Feedback");

  // 2. Get the specific row that was edited in Report sheet
  const row = e.range.getRow();
  
  // Fetch columns A through L (12 columns) from Complaint Report1
  const rowData = sourceSheet.getRange(row, 1, 1, 12).getValues()[0]; 

 // --- Corrected Mapping from Complaint Report1 ---
const timestamp      = rowData[0];  // Col A
const complaintId    = rowData[1];  // Col B
const customerName   = rowData[2];  // Col C (FIXED: changed row.Data to rowData)
const serviceType    = rowData[7];  // Col H
const resolvedStatus = rowData[11]; // Col L

// 3. Only proceed if "Complaint Resolved?" is "Yes"
if (resolvedStatus === "Yes") {
  
  const feedbackSheet = ss.getSheetByName("BPMS - Customer Feedback");
  const startRow = 7;
  const lastRow = feedbackSheet.getLastRow();
  const destRow = Math.max(startRow, lastRow + 1);

  // 5. Prepare data for the 5 columns you listed:
  // Timestamp, Complaint ID, Customer Name, Service Type, Complaint Resolved?
  const finalData = [
    timestamp, 
    complaintId, 
    customerName, 
    serviceType, 
    resolvedStatus
  ];

  // 6. Write data (5 columns)
  feedbackSheet.getRange(destRow, 1, 1, 5).setValues([finalData]);
  
  Logger.log("Data added to Feedback Sheet at Row: " + destRow);
    // Optional: Mark as processed in the Report sheet
    // sourceSheet.getRange(row, 13).setValue("Moved to Feedback");
  }
}