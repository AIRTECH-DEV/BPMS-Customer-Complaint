function generatePaymentFollowup() {
  const ss = SpreadsheetApp.getActive();
  const clean = ss.getSheetByName("Cleaned Data");
  const report = ss.getSheetByName("Complaint Report1");
  const pay = ss.getSheetByName("BPMS - Payment Followup");

  if (!clean || !report || !pay) {
    Logger.log("Error: One or more sheets not found.");
    return;
  }

  // ---- 1. Clear old rows (Starting from Row 7) ----
  const payLastRow = pay.getLastRow();
  if (payLastRow >= 7) {
    pay.getRange(7, 1, payLastRow - 6, 8).clearContent();
  }

  // ---- Helper: Normalize text ----
  const normalize = (str) => {
    if (!str) return "";
    return String(str).replace(/[–—]/g, "-").trim().toLowerCase();
  };

  // ---- 2. Read Cleaned Data (Lookup) ----
  const cleanLast = clean.getLastRow();
  if (cleanLast < 2) return;
  const cleanVals = clean.getRange(2, 1, cleanLast - 1, 5).getValues();

  const cleanMap = new Map();
  cleanVals.forEach(r => {
    cleanMap.set(normalize(r[1]), {
      company: r[2],
      address: r[3],
      serviceType: r[4]
    });
  });

  // ---- 3. Read Complaint Report1 ----
  const repLast = report.getLastRow();
  if (repLast < 2) return;
  const repVals = report.getRange(2, 1, repLast - 1, report.getLastColumn()).getValues();

  // ---- 4. Filter and Process ----
  const finalMap = new Map();

  repVals.forEach(r => {
    const serviceType = normalize(r[14]); // Column K (Service Type)
    const cid = normalize(r[1]);          // Column B (Complaint ID)
    
    // Only include if Service Type is "Paid"
    if (serviceType === "paid" && cid !== "") {
      const cleanInfo = cleanMap.get(cid);
      
      finalMap.set(cid, [
        r[0],                          // Col A: Timestamp
        r[1],                          // Col B: ID
        cleanInfo ? cleanInfo.company : r[2], // Col C: Customer
        cleanInfo ? cleanInfo.address : r[3], // Col D: Address
        r[14],                         // Col E: Service Type
        r[18] || "Pending",            // Col F: Resolved? (Column O)
        r[19] || "N/A",                // Col G: Tech Name (Column N)
        r[4] || ""                     // Col H: Issue (Column E)
      ]);
    }
  });

  // ---- 5. Convert to Array and Sort ----
  const finalRows = Array.from(finalMap.values());

  // Sort by Date (Oldest at top)
  finalRows.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  // ---- 6. Write to sheet ----
  if (finalRows.length > 0) {
    pay.getRange(7, 1, finalRows.length, 8).setValues(finalRows);
    SpreadsheetApp.flush();
  }

  Logger.log("Payment follow-up updated. Total Rows: " + finalRows.length);
}