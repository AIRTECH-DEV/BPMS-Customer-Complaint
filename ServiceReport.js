const MASTER_SHEET_ID = "1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms";
//const REPORT_SHEET_ID ="1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Digital Complaint Report')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getCustomerList() {
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
      .filter(row => String(row[18]).trim().toLowerCase() === "yes")
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
}

function getCustomerDetails(fullName) {
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
}



  function submitComplaint(formData) {
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


    
  



