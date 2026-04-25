function generateLabourAMC() {
  const ss = SpreadsheetApp.getActive();
  const report = ss.getSheetByName("Complaint Report1");
  const cleanedData = ss.getSheetByName("Cleaned Data");
  const labourSheet = ss.getSheetByName("BPMS - Labour AMC");

  if (!report || !cleanedData || !labourSheet) return;

  // 1. MAP CLEANED DATA (Reference for Company, Client, and Type)
  const cleanVals = cleanedData.getDataRange().getValues();
  const cleanLookup = new Map();
  cleanVals.forEach(row => {
    const cid = row[1] ? row[1].toString().trim() : "";
    if (cid) {
      cleanLookup.set(cid, { 
        company: row[2], // Col C
        client: row[3],  // Col D
        type: row[7]     // Col H
      });
    }
  });

  // 2. PROCESS REPORT DATA
  const lastRow = report.getLastRow();
  if (lastRow < 2) return;
  const repVals = report.getRange(2, 1, lastRow - 1, report.getLastColumn()).getValues();
  
  /**
   * uniqueAmcMap will store the data.
   * By using .delete() before .set(), we force the updated ID to move 
   * to the end of the Map, which translates to the bottom of the sheet.
   */
  const uniqueAmcMap = new Map();

  repVals.forEach(r => {
    const service = (r[14] || "").toString().toUpperCase().trim();
    const cid = r[1] ? r[1].toString().trim() : "";
    
    if (service === "AMC" && cid !== "") {
      const info = cleanLookup.get(cid) || { company: "N/A", client: "N/A", type: "N/A" };
      
      const rowData = [
        r[0],         // Timestamp
        cid,          // Complaint ID
        info.company, // Company Name (from Cleaned)
        info.client,  // Client Name (from Cleaned)
        r[14],        // Service Type (from Report)
        r[4],         // Issues (from Report - gets updated)
        info.type     // Customer Type (from Cleaned Col H)
      ];

      // THE "MOVE TO BOTTOM" LOGIC:
      // If the ID already exists in our list, delete the old one first.
      if (uniqueAmcMap.has(cid)) {
        uniqueAmcMap.delete(cid);
      }
      // Re-inserting it puts it at the "end" of the Map
      uniqueAmcMap.set(cid, rowData);
    }
  });

  // 3. CONVERT MAP VALUES TO ARRAY
  const finalRows = Array.from(uniqueAmcMap.values());

  // 4. WRITE TO SHEET
  if (finalRows.length > 0) {
    // Clear everything from Row 7 down to remove old order
    const currentLast = labourSheet.getLastRow();
    if (currentLast >= 7) {
      labourSheet.getRange(7, 1, currentLast - 6, 7).clearContent();
    }
    
    // Paste the data. The updated IDs will naturally appear at the bottom.
    labourSheet.getRange(7, 1, finalRows.length, 7).setValues(finalRows);
    SpreadsheetApp.flush();
  }
}
