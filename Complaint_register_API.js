// 1. Store your token in File > Project Settings > Script Properties
const scriptProperties = PropertiesService.getScriptProperties();
const META_ACCESS_TOKEN = scriptProperties.getProperty('META_TOKEN'); 
const PHONE_NUMBER_ID = "1002193126304358"; 
const TEMPLATE_NAME = "complaint_register"; 

// 1. DEFINE YOUR SHEET NAME HERE
const MY_SHEET_NAME = "Complaint Entry"; 

function Com_register(e) {
  if (!e) return;

  try {
    // 2. Explicitly select the sheet by name
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MY_SHEET_NAME);
    
    // Safety check: Make sure the sheet name is spelled correctly
    if (!sheet) {
      console.error("Sheet not found: " + MY_SHEET_NAME);
      return;
    }

    // 3. Get the row that was just submitted
    const row = e.range.getRow();
    
    // Give formulas time to calculate the Complaint ID in that row
    Utilities.sleep(6000); 

    // 4. Fetch the data from your specific sheet
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    let data = {};
    headers.forEach((h, i) => {
      if (h) data[h.toString().trim()] = rowData[i];
    });

    // 5. Logic for extraction
    const complaintId = data["Complaint ID"];
    const customerType = String(data["Customer Type"] || "").trim();
    
   
    
     let finalPhone;

// We use the 'rowData' array directly to avoid duplicate column name issues.
// index 16 = Column Q (Contact Person Number)
// index 29 = Column AD (Phone Number)

if (customerType.includes("Commercial")) {
  finalPhone = rowData[16]; 
} else if (customerType.includes("Residential")) {
  finalPhone = rowData[29];
} else {
  // Fallback: Default to Column Q if no match is found
  //finalPhone = rowData[16];
}
    if (!finalPhone || !complaintId) {
      throw new Error("Missing data in " + MY_SHEET_NAME);
    }

    // --- Format and Send ---
    let cleanPhone = String(finalPhone).replace(/\D/g, "");
    if (cleanPhone.length === 10) cleanPhone = "91" + cleanPhone;

    const payload = {
      "messaging_product": "whatsapp",
      "to": cleanPhone,
      "type": "template",
      "template": {
        "name": TEMPLATE_NAME,
        "language": { "code": "en" }, 
        "components": [{
          "type": "body",
          "parameters": [{"type": "text", "text": String(complaintId)}]
        }]
      }
    };
    

    const options = {
      "method": "post",
      "contentType": "application/json",
      "headers": { "Authorization": "Bearer " + META_ACCESS_TOKEN },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(`https://graph.facebook.com/v18.0/${PHONE_NUMBER_ID}/messages`, options);
    console.log("Response for row " + row + ": " + response.getContentText());

  } catch (err) {
    console.error("Fatal Error: " + err.message);
  }
}