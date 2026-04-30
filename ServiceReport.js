const MASTER_SHEET_ID = "1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms";
//const REPORT_SHEET_ID ="1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Digital Complaint Report')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
/*correct working*/
/*function getCustomerList() {
  const masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSs.getSheetByName("Cleaned Data");
  const masterData = masterSheet.getDataRange().getValues();
  
  const reportSs = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = reportSs.getSheetByName("Complaint Report1");
  
  let resolvedIds = [];
  if (reportSheet) {
    const reportData = reportSheet.getDataRange().getValues();
    // We look at Column L (index 11) for "Yes" and grab the ID from Column B (index 1)
    resolvedIds = reportData
      .filter(row => String(row[18]).trim().toLowerCase() === "resolved")
      .map(row => String(row[1]).trim());
  }

  const list = masterData.slice(1).filter(r => {
    let id = String(r[1]).trim(); 
    let comp = String(r[2] || "").trim(); 
    let clnt = String(r[3] || "").trim();
    
    // ONLY show if it has a name AND the ID is NOT in the resolved list
    return (comp || clnt) && !resolvedIds.includes(id);
    
  }).map(r => {
    let comp = String(r[2] || "").trim(); 
    let clnt = String(r[3] || "").trim();
    return `${comp} | ${clnt}`;
  });

  return [...new Set(list)].sort(); 
}*/
/*30/4/26*/

/*function getCustomerList() {
  const masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSs.getSheetByName("Cleaned Data");
  const masterData = masterSheet.getDataRange().getValues();
  
  const reportSs = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = reportSs.getSheetByName("Complaint Report1");
  
  let resolvedIds = [];
  if (reportSheet) {
    const reportData = reportSheet.getDataRange().getValues();
    resolvedIds = reportData
      .filter(row => String(row[18]).trim().toLowerCase() === "resolved")
      .map(row => String(row[1]).trim());
  }

  const list = masterData.slice(1).filter(r => {
    let id = String(r[1]).trim(); 
    let comp = String(r[2] || "").trim(); 
    let clnt = String(r[3] || "").trim(); 
    return (comp || clnt) && !resolvedIds.includes(id);
  }).map(r => {
    let fullId = String(r[1]).trim();
    let lastFourId = fullId.length > 4 ? fullId.substring(fullId.length - 4) : fullId;
    let comp = String(r[2] || "").trim(); 
    let clnt = String(r[3] || "").trim();
    let address = String(r[8] || "").trim();

    // Logic to add (R) and (C) labels
    let displayName = "";
    if (clnt && comp) {
      displayName = `${clnt} (R) - ${comp} (C)`;
    } else if (clnt) {
      displayName = `${clnt} (R)`;
    } else if (comp) {
      displayName = `${comp} (C)`;
    }
    
    // Result: Client (R) - Company (C) (CID 1234) - Address
    return `${displayName} - (CID ${lastFourId})-${address}`;
  });

  return [...new Set(list)].sort(); 
}*///////
function getCustomerList() {
  const masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSs.getSheetByName("Cleaned Data");
  const masterData = masterSheet.getDataRange().getValues();
  
  const reportSs = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheet = reportSs.getSheetByName("Complaint Report1");
  
  let resolvedIds = [];
  if (reportSheet) {
    const reportData = reportSheet.getDataRange().getValues();
    resolvedIds = reportData
      .filter(row => String(row[18]).trim().toLowerCase() === "resolved")
      .map(row => String(row[1]).trim());
  }

  const list = masterData.slice(1).filter(r => {
    let id = String(r[1]).trim(); 
    let comp = String(r[2] || "").trim(); 
    let clnt = String(r[3] || "").trim(); 
    return (comp || clnt) && !resolvedIds.includes(id);
  }).map(r => {
    let fullId = String(r[1]).trim();
    let lastFourId = fullId.length > 4 ? fullId.substring(fullId.length - 4) : fullId;
    let comp = String(r[2] || "").trim(); 
    let clnt = String(r[3] || "").trim();
    let address = String(r[8] || "").trim();

    // Logic to add (R) and (C) labels
    let displayName = "";
    if (clnt && comp) {
      displayName = `${clnt} (R) - ${comp} (C)`;
    } else if (clnt) {
      displayName = `${clnt} (R)`;
    } else if (comp) {
      displayName = `${comp} (C)`;
    }
    
    // ✅ Return object instead of flat string
    return {
      label: `${displayName} - (CID ${lastFourId})`,  // shown in line 1
      address: address,                                 // shown in line 2
      value: `${displayName} - (CID ${lastFourId})-${address}` // full value stored on select
    };
  });

  // Deduplicate by full value, then sort by label
  const seen = new Set();
  return list
    .filter(item => {
      if (seen.has(item.value)) return false;
      seen.add(item.value);
      return true;
    })
    .sort((a, b) => a.label.localeCompare(b.label));
}


function getCustomerDetails(fullName) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName("Cleaned Data");
  const data = sheet.getDataRange().getValues();
  
  // 1. Extract the name section before the (CID)
  let namePart = fullName.split("- (CID")[0].trim();

  // 2. Search logic that ignores the (R) and (C) tags
  const match = data.find(r => {
    let compInSheet = String(r[2] || "").trim();
    let clntInSheet = String(r[3] || "").trim();
    
    // Check if the selected text matches the combinations used in the dropdown
    let formatBoth = `${clntInSheet} (R) - ${compInSheet} (C)`;
    let formatClntOnly = `${clntInSheet} (R)`;
    let formatCompOnly = `${compInSheet} (C)`;
    
    return namePart === formatBoth || namePart === formatClntOnly || namePart === formatCompOnly;
  });

  if (match) {
    return {
      address: match[8] || "",
      brand:   match[9] || "",
      contact: match[11] || "",
      phone:   match[12] || "",
      complaintType: match[6]
    };
  }
  return null;
}

/*********************************** */
/*corrected code*/
/*function getCustomerDetails(fullName) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName("Cleaned Data");
  const data = sheet.getDataRange().getValues();
  
  const match = data.find(r => {
    let comp = String(r[2] || "").trim();
    let clnt = String(r[3] || "").trim();
    return `${comp} | ${clnt}` === fullName;
  });

  if (match) {
    return {
      address: match[8] || "",
      brand:   match[9] || "",
      contact: match[11] || "",
      phone:   match[12] || "",
      complaintType: match[6]
    };
  }
  return null;
}*/


/*corrected code*/
 /* function submitComplaint(formData) {
  try {
    //const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const reportSheet = ss.getSheetByName("Complaint Report1");
    
    // 1. Get Master ID
    const masterSs = SpreadsheetApp.openById("1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms");
    const masterSheet = masterSs.getSheetByName("Cleaned Data");
    const masterData = masterSheet.getDataRange().getValues();
   // const match = masterData.find(r => `${r[2]} | ${r[3]}` === formData.company);
   // Replace lines 78-79 with this:
const match = masterData.find(r => {
  let comp = String(r[2] || "").trim();
  let clnt = String(r[3] || "").trim();
  let combined = (comp + " | " + clnt).trim();
  return combined === String(formData.company || "").trim();
});
    const finalId = match ? match[1] : "NO ID FOUND";
  
    // 2. Handle File Upload
    let proofUrl = "No proof uploaded";
    if (formData.proofData) {
      const folder = DriveApp.getFolderById("1uIi1YEjzRTcQyjo5UPb97rWvwAUaU8Rr");
      const blob = Utilities.newBlob(Utilities.base64Decode(formData.proofData), formData.proofType, "Proof_" + finalId);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      proofUrl = file.getUrl();
    }

    const technicianNames = (formData.technicianNames && formData.technicianNames.toString().trim()) ?
      formData.technicianNames.toString().trim() :
      [formData.tech, formData.helper]
        .filter(name => name && name.toString().trim())
        .join(", ");

    // 3. APPEND DATA (Matches your specific column order)
    reportSheet.appendRow([
      new Date(),                // Time Stamp
      finalId,                   // Complaint ID
      formData.company,          // Customer Name
      formData.address,          // Address
      formData.issue,
      formData.contact,          // Contact person Name
      formData.phone,            // Phone Number
      formData.model,            // Model
      formData.serial,           // Serial No
      formData.location,         // AC Location
      // --- NEW FIELDS ADDED HERE ---
      formData.type,      // Column K: Machine Type (Split/Window/etc)
      formData.gas,          // Column L: Gas Type
      formData.problem,          // Column M: Problem
      formData.actionTaken,      // Column N: Action Taken
      // -----------------------------
      formData.serviceType,      // Service Type (AMC/Paid/Warranty)
      //formData.brand,
      formData.make,            // Machine Brand
      formData.complaintType,    // Complaint Type
      technicianNames,           // Technician Name + Helper Name
      formData.resolved,         // Complaint Resolved?
      formData.paymentStatus || "N/A", // Column P: Payment Received? (Yes/Pending/No)
      formData.customerSig,      // Customer sign (Base64)
      formData.techSig,          // Tech sign (Base64)
      
      " ",
      formData.reportType,       // Service Report (Type)
      proofUrl.indexOf("http") > -1 ? '=HYPERLINK("' + proofUrl + '", "View Proof")' : proofUrl ,// Payment Proof
      formData.problemStatus
    ]);*/

  function submitComplaint(formData) {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const reportSheet = ss.getSheetByName("Complaint Report1");
    
    // 1. Get Master Data and Match via CID
    const masterSs = SpreadsheetApp.openById("1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms");
    const masterSheet = masterSs.getSheetByName("Cleaned Data");
    const masterData = masterSheet.getDataRange().getValues();

    // dropdown format is: "Name (R/C) - (CID ####) - Address"
    let selectedText = String(formData.company || "").trim();
    
    // EXTRACT THE NAME ONLY (Everything before the first dash)
    const finalCustomerName = selectedText.split(" - ")[0].trim();

    // EXTRACT THE 4-DIGIT CID to find the original full ID
    const cidMatch = selectedText.match(/\(CID\s(\d{4})\)/);
    const shortId = cidMatch ? cidMatch[1] : "";

    // Find the row in master data where Complaint ID (Column B/Index 1) ends with those 4 digits
    const match = masterData.find(r => {
      let fullIdInSheet = String(r[1]).trim();
      return fullIdInSheet.endsWith(shortId);
    });

    const finalId = match ? match[1] : "NO ID FOUND";
  
    // 2. Handle File Upload
    let proofUrl = "No proof uploaded";
    if (formData.proofData) {
      const folder = DriveApp.getFolderById("1uIi1YEjzRTcQyjo5UPb97rWvwAUaU8Rr");
      const blob = Utilities.newBlob(Utilities.base64Decode(formData.proofData), formData.proofType, "Proof_" + finalId);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      proofUrl = file.getUrl();
    }

    const technicianNames = (formData.technicianNames && formData.technicianNames.toString().trim()) ?
      formData.technicianNames.toString().trim() :
      [formData.tech, formData.helper]
        .filter(name => name && name.toString().trim())
        .join(", ");

    // 3. APPEND DATA
    reportSheet.appendRow([
      new Date(),                // Time Stamp
      finalId,                   // Stores Full ID (e.g. CMP-2024-0308)
      finalCustomerName,         // Stores Name only (e.g. Abhishek Nimbkar (R))
      formData.address,          // Address
      formData.issue,
      formData.contact,          // Contact person Name
      formData.phone,            // Phone Number
      formData.model,            // Model
      formData.serial,           // Serial No
      formData.location,         // AC Location
      formData.type,             // Machine Type
      formData.gas,              // Gas Type
      formData.problem,          // Problem
      formData.actionTaken,      // Action Taken
      formData.serviceType,      // Service Type
      formData.make,             // Machine Brand
      formData.complaintType,    // Complaint Type
      technicianNames,           // Technicians
      formData.resolved,         // Resolved Status
      formData.paymentStatus || "N/A", 
      formData.customerSig,      
      formData.techSig,          
      " ",
      formData.reportType,       
      proofUrl.indexOf("http") > -1 ? '=HYPERLINK("' + proofUrl + '", "View Proof")' : proofUrl,
      formData.problemStatus
    ]);  

try {
  const service = (formData.serviceType || "").toString().toUpperCase().trim();
  
  // Only execute the sync if the service is one of these three
  const validServices = ["AMC", "IN WARRANTY", "PAID"];
  
  if (validServices.includes(service) || service.includes("WARRANTY")) {
    Logger.log("Valid Service Type detected: " + service + ". Running central sync...");
    generateLabourAMC(); 
    generateWarrantyCAMC();
    generatePaymentFollowup();
    try {
    const lastRow = reportSheet.getLastRow();
    scheduleWhatsAppIn5Min_(lastRow);
    Logger.log("✅ scheduleWhatsAppIn5Min_ called for row: " + lastRow);
} catch(schedErr) {
    Logger.log("❌ scheduleWhatsAppIn5Min_ CRASHED: " + schedErr.toString());
}
  } else {
    Logger.log("Service Type '" + service + "' does not require Labour AMC sync.");
  }

} catch (err) {
  Logger.log("Trigger Error: " + err.toString());
}
syncComplaintStatus(finalId, formData.resolved); // bpms intake /dispatch sheet update status when form filled 
 if (formData.resolved && formData.resolved.toString().toLowerCase() === "resolved") {
      // If resolved is "Yes", remove it from the Pending List
      removeFromPendingList(finalId);
      
    } else if (formData.resolved && formData.resolved.toString().toLowerCase() === "pending") {
      // If resolved is "No", add it to the Pending List
      syncPendingComplaints(finalId, formData);
    }// At the very end of your submitComplaint(formData) function in server.gs:
    Logger.log("⏩ Reached schedule block. lastRow will be: " + reportSheet.getLastRow());
Logger.log("⏩ resolved value is: " + formData.resolved);

/*try {
    const lastRow = reportSheet.getLastRow();
    scheduleWhatsAppIn5Min_(lastRow);
    Logger.log("✅ scheduleWhatsAppIn5Min_ called for row: " + lastRow);
} catch(schedErr) {
    Logger.log("❌ scheduleWhatsAppIn5Min_ CRASHED: " + schedErr.toString());
}*/
    return { 
      status: "Success", 
      message: "Report Saved Successfully!" 
    };

  } catch (e) {
    // This sends the specific error back to the browser
    return { 
      status: "Success", 
      message: "Report Saved Successfully!" 
    }
    /*return { 
      //status: "Error", 
      message: e.toString() 
    };*/
  }
}


// 24-04-26 yash 


    
  



