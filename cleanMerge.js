/****************************************************
 * BPMS Complaint Form Cleaner
 * - Reads new submissions from Form Responses 1
 * - Appends a cleaned, unified row to Cleaned Data
 * - Preserves the raw sheet for backup
 ****************************************************/
function handleComplaintForm(e) {
  const ss = e.source;
  const raw = ss.getSheetByName("Complaint Entry");
  const clean = ss.getSheetByName("Cleaned Data");
  const row = e.range.getRow();
  const vals = raw.getRange(row, 1, 1, raw.getLastColumn()).getValues()[0];

  // Helper functions
  const v = (idx) => vals[idx - 1] ? String(vals[idx - 1]).trim() : "";
  const pick = (...idxs) => idxs.map(v).find(x => x !== "") || "";
  // const joinAddr = (...parts) => parts.filter(Boolean).join(", ");
  const joinAddr = (...parts) => parts.filter(p => p && p.trim() !== "").join(", ");
  // const formatDateTime = (value) =>
  //   value ? Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "dd-MMM-yyyy HH:mm") : "";

  const map = {
    timestamp: 1,
    complaintID: 2,
    companyName: 3,
    name: 4,
    serviceType1: 5,
    emailAddress: 6,
    complaintType: 7,
    customerType: 8,
    siteFlat1: 9,
    siteBldg1: 10,
    siteStreet1: 11,
    siteCity1: 12,
    sitePin1: 13,
    machineBrand1: 14,
    machineType1: 15,
    contactName1: 16,
    contactNumber1: 17,
    contactEmail1: 18,
    phoneNumber: 19,
    emailID: 20,
    flat2: 21,
    bldg2: 22,
    street2: 23,
    city2: 24,
    pin2: 25,
    serviceType2: 26,
    machineBrand2: 27,
    machineType2: 28,
    contactName2: 29,
    contactNumber2: 30,
    contactEmail2: 31
  };

  const serviceType = pick(map.serviceType1, map.serviceType2);
  const fullAddress = joinAddr(pick(map.siteFlat1, map.flat2), pick(map.siteBldg1, map.bldg2),
                               pick(map.siteStreet1, map.street2), pick(map.siteCity1, map.city2),
                               pick(map.sitePin1, map.pin2));
  const machineBrand = pick(map.machineBrand1, map.machineBrand2);
  const machineType  = pick(map.machineType1, map.machineType2);
  const contactName  = pick(map.contactName1, map.contactName2);
  const contactNum   = pick(map.contactNumber1, map.contactNumber2);
  const contactEmail = pick(map.contactEmail1, map.contactEmail2);

  const customerType = v(map.customerType);
  // const companyName = /commercial/i.test(customerType) ? v(map.companyName) : ""; 

  // Append only the merged columns (example: starting from column E onward)
  const record = [
    serviceType,
    v(map.emailAddress),
    v(map.complaintType),
    customerType,
    fullAddress,
    machineBrand,
    machineType,
    contactName,
    contactNum,
    contactEmail
  ];

  clean.getRange(row, 5, 1, record.length).setValues([record]);
}

/****************************************************
 * Bulk Conversion Utility
 * - Converts all existing form data into Cleaned Data
 * - Useful for backfilling old records
 ****************************************************/
function convertExistingFormResponses() {
  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName("Complaint Entry");
  const clean = ss.getSheetByName("Cleaned Data");

  // Clear existing cleaned data (except headers)
  const lastCleanRow = clean.getLastRow();
  if (lastCleanRow > 1) clean.getRange(2, 1, lastCleanRow - 1, clean.getLastColumn()).clearContent();

  const lastRow = raw.getLastRow();
  const lastCol = raw.getLastColumn();
  if (lastRow < 2) {
    Logger.log("No data found in Complaint Entry.");
    return;
  }

  const data = raw.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Helper to format timestamps
  function formatDateTime(value) {
    if (!value) return "";
    const d = new Date(value);
    if (isNaN(d)) return value;
    const options = {
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    };
    return d.toLocaleString("en-GB", options).replace(",", "");
  }

  const map = {
    timestamp: 1,
    complaintID: 2,
    companyName: 3,
    name: 4,
    serviceType1: 5,
    emailAddress: 6,
    complaintType: 7,
    customerType: 8,
    siteFlat1: 9,
    siteBldg1: 10,
    siteStreet1: 11,
    siteCity1: 12,
    sitePin1: 13,
    machineBrand1: 14,
    machineType1: 15,
    contactName1: 16,
    contactNumber1: 17,
    contactEmail1: 18,
    phoneNumber: 19,
    emailID: 20,
    flat2: 21,
    bldg2: 22,
    street2: 23,
    city2: 24,
    pin2: 25,
    serviceType2: 26,
    machineBrand2: 27,
    machineType2: 28,
    contactName2: 29,
    contactNumber2: 30,
    contactEmail2: 31
  };

  const pick = (vals, ...idxs) => {
    for (const i of idxs) {
      const x = vals[i - 1];
      if (x && String(x).trim() !== "") return String(x).trim();
    }
    return "";
  };
  const joinAddr = (...parts) => parts.filter(p => p && p.trim() !== "").join(", ");

  const cleanedRows = data.map(vals => {
    const customerType = pick(vals, map.customerType);
    const serviceType = pick(vals, map.serviceType1, map.serviceType2);
    const fullAddress = joinAddr(
      pick(vals, map.siteFlat1, map.flat2),
      pick(vals, map.siteBldg1, map.bldg2),
      pick(vals, map.siteStreet1, map.street2),
      pick(vals, map.siteCity1, map.city2),
      pick(vals, map.sitePin1, map.pin2)
    );
    const machineBrand = pick(vals, map.machineBrand1, map.machineBrand2);
    const machineType = pick(vals, map.machineType1, map.machineType2);
    const contactName = pick(vals, map.contactName1, map.contactName2);
    const contactNum = pick(vals, map.contactNumber1, map.contactNumber2);
    const contactEmail = pick(vals, map.contactEmail1, map.contactEmail2);
    const companyName = /commercial/i.test(customerType) ? pick(vals, map.companyName) : "";

    return [
      formatDateTime(pick(vals, map.timestamp)), // ✅ formatted Timestamp
      pick(vals, map.complaintID),
      companyName,
      pick(vals, map.name),
      serviceType,
      pick(vals, map.emailAddress),
      pick(vals, map.complaintType),
      customerType,
      fullAddress,
      machineBrand,
      machineType,
      contactName,
      contactNum,
      contactEmail,
      // pick(vals, map.phoneNumber),
      // pick(vals, map.emailID),
      // pick(vals, map.serviceType2),
      // pick(vals, map.machineBrand2),
      // pick(vals, map.machineType2),
      // pick(vals, map.contactName2),
      // pick(vals, map.contactNumber2),
      // pick(vals, map.contactEmail2)
    ];
  });

  if (cleanedRows.length > 0) {
    clean.getRange(2, 1, cleanedRows.length, cleanedRows[0].length).setValues(cleanedRows);
    clean.getRange("A2:A").setNumberFormat("dd-mmm-yyyy hh:mm"); // ✅ enforce DateTime format
    Logger.log(`✅ ${cleanedRows.length} rows converted successfully.`);
  }
}

/**
 * Seeds robust ARRAYFORMULA columns (A–D) in "Cleaned Data".
 * - A: Timestamp (direct from Complaint Entry!A)
 * - B: Complaint ID (direct from Complaint Entry!B)
 * - C: Company Name (only if Customer Type contains "commercial")
 * - D: Name (always from Complaint Entry!D)
 *
 * This avoids per-row INDEX and prevents “carry-over” values.
 */
function seedCleanedDataHeaderFormulas() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const clean = ss.getSheetByName("Cleaned Data");
  if (!clean) throw new Error('Sheet "Cleaned Data" not found');

  // Clear existing formulas/values in A:D (except headers in row 1)
  const lastRow = Math.max(clean.getLastRow(), 2);
  clean.getRange(2, 1, lastRow - 1, 4).clearContent();

  // Set ARRAYFORMULA in A2:D2 so it auto-fills forever
  // A: Timestamp
  clean.getRange("A2").setFormula(
    `=ARRAYFORMULA(IF(LEN('Complaint Entry'!A2:A)=0, , 'Complaint Entry'!A2:A))`
  );

  // B: Complaint ID
  clean.getRange("B2").setFormula(
    `=ARRAYFORMULA(IF(LEN('Complaint Entry'!B2:B)=0, , 'Complaint Entry'!B2:B))`
  );

  // C: Company Name (only when customer type includes "commercial")
  // Uses REGEXMATCH(LOWER()) so it’s resilient to variations like "Commercial Client"
  clean.getRange("C2").setFormula(
    `=ARRAYFORMULA(IF(LEN('Complaint Entry'!H2:H)=0, ,
       IF(REGEXMATCH(LOWER('Complaint Entry'!H2:H), "commercial"),
          'Complaint Entry'!C2:C,
          ""
       )
     ))`
  );

  // D: Name (always from column D of Complaint Entry)
  clean.getRange("D2").setFormula(
    `=ARRAYFORMULA(IF(LEN('Complaint Entry'!D2:D)=0, , 'Complaint Entry'!D2:D))`
  );

  // Optional: format timestamps (entire column A)
  clean.getRange("A:A").setNumberFormat("dd-mmm-yyyy hh:mm");
}

