// const TEMPLATE_NAME1 = "complaint_resolution_report1"
// function Complaint_resolution() {
//   // Wait for the sheet to stabilize after form submission
//   Utilities.sleep(60000); 
// try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const responseSheet = ss.getSheetByName("Complaint Report1"); 
//     const lastRow = responseSheet.getLastRow();
    
//     const headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
//     const rowData = responseSheet.getRange(lastRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];

//     let formData = {};
//     headers.forEach((header, index) => {
//       formData[header] = rowData[index];
//     });

//     let status = (formData["Complaint Resolved"] || formData["Complaint Resolved?"] || "").toString().trim().toLowerCase();
//     if (status !== "yes") {
//       Logger.log("Skipping: Status is '" + status + "', not 'yes'.");
//       return;
//     }

//     const complaintId = formData["Complaint ID"];
//     //const fileIdRaw = formData["Service Report"] || "";

//     // LOOKUP: Phone Number from "Cleaned Data"
//     const cleanedSheet = ss.getSheetByName("Cleaned Data");
//     const cleanedValues = cleanedSheet.getDataRange().getValues();
//     let finalPhone = "";
     
//     for (let i = 1; i < cleanedValues.length; i++) {
//       // Check Column B (index 1) for a match with the Complaint ID
//       if (String(cleanedValues[i][1]).trim() == String(complaintId).trim()) { 
//         finalPhone = cleanedValues[i][12]; // Get phone from Column M (index 12)
//         break;
//       }
//     }

//     if (!finalPhone) {
//       Logger.log("ERROR: No phone found in 'Cleaned Data' for ID: " + complaintId);
//       return;
//     }
//    /* const rowFormulas = responseSheet.getRange(lastRow, 1, 1, responseSheet.getLastColumn()).getFormulas()[0];
//     let finalViewLink = "";
//     let rawFormula = rowFormulas[18] || ""; // Index 18 is Column S

//     if (rawFormula.includes("HYPERLINK")) {
//       const urlMatch = rawFormula.match(/"([^"]+)"/);
//       if (urlMatch) finalViewLink = urlMatch[1];
//     } else {
//       // Fallback if it's a plain link and not a formula
//       finalViewLink = formData["Service Report"] || "";
//     }

//     // Check if link is empty before proceeding
//     if (!finalViewLink) {
//       Logger.log("ERROR: No link found in Column S for ID: " + complaintId);
//       return;
//     }
  
//     // 3. Ensure the file is accessible
//     if (finalViewLink.includes("drive.google.com")) {
//       try {
//         const fileIdMatch = finalViewLink.match(/[-\w]{25,}/);
//         if (fileIdMatch) {
//           DriveApp.getFileById(fileIdMatch[0]).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
//         }
//       } catch(err) {
//         Logger.log("Permission update failed: " + err.message);
//       }
//     }*/
//     // --- UPDATED RETRY LOGIC TO EXTRACT LINK FROM HYPERLINK FORMULA ---
//     let finalViewLink = "";
//     let attempts = 0;
//     const maxAttempts = 12; 

//     while (!finalViewLink && attempts < maxAttempts) {
//       // 1. Get the FORMULA and the VALUE for Column S (Index 18)
//       const range = responseSheet.getRange(lastRow, 18); // Column S is 19
//       const cellFormula = range.getFormula();
//       const cellValue = range.getValue();
      
//       // 2. Check if it's a HYPERLINK formula
//       if (cellFormula && cellFormula.includes("HYPERLINK")) {
//         // This regex looks for the first URL inside the quotes of the formula
//         const urlMatch = cellFormula.match(/"([^"]+)"/);
//         if (urlMatch && urlMatch[1]) {
//           finalViewLink = urlMatch[1];
//           Logger.log("Successfully extracted URL from formula: " + finalViewLink);
//         }
//       } 
//       // 3. Fallback if the cell just contains a raw URL string
//       else if (cellValue && cellValue.toString().startsWith("http")) {
//         finalViewLink = cellValue;
//       }

//       if (!finalViewLink) {
//         Logger.log("Link not found yet. Attempt " + (attempts + 1));
//         Utilities.sleep(5000); 
//         attempts++;
//       }
//     }
//     // --- END OF NEW RETRY LOGIC ---

//     // Format Phone to include country code
//     let cleanPhone = finalPhone.toString().replace(/\D/g, "");
//     if (cleanPhone.length === 10) cleanPhone = "91" + cleanPhone;
    
    
    
//     // Link Extraction and Permission Setting
//     /*let finalViewLink = "https://drive.google.com";
//     const fileIdMatch = fileIdRaw.match(/[-\w]{25,}/); 
//     //const fileIdMatch = fileIdRaw.match(/\/d\/([-\w]{25,})/);
//     if (fileIdMatch) {
//       const driveFileId = fileIdMatch[0];
//       try {
//         const file = DriveApp.getFileById(driveFileId);
//         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
//        // finalViewLink = "https://drive.google.com/open?id=" + driveFileId;
//        finalViewLink = "https://drive.google.com/file/d/" + driveFileId + "/view?usp=sharing";
//       } catch(err) {
//         // Fallback to the direct open link if specific file retrieval fails
//         //finalViewLink = "https://drive.google.com/open?id=" + driveFileId;
//         finalViewLink = "https://drive.google.com/uc?id=" + driveFileId;
//       }
//     }*/
    
//     // Send WhatsApp Message using NAMED PARAMETERS
    
//     // Send WhatsApp Message using NAMED PARAMETERS
//     const payload = {
//       "messaging_product": "whatsapp",
//       "to": cleanPhone,
//       "type": "template",
//       "template": {
//         "name": TEMPLATE_NAME1,
//         "language": { "code": "en_US" }, 
//         "components": [{
//           "type": "body",
//           "parameters": [
//             { 
//               "type": "text", 
//               "parameter_name": "complaint_id", // CHANGED to lowercase 'id' to match your template
//               "text": String(complaintId) 
//             },
//             { 
//               "type": "text", 
//               "parameter_name": "report_link", // Matches your template {{report_link}}
//               "text": String(finalViewLink) 
//             }
//           ]
//         }]
//       }
//     };

//     const options = {
//       "method": "post",
//       "contentType": "application/json",
//       "headers": { "Authorization": "Bearer " + META_ACCESS_TOKEN },
//       "payload": JSON.stringify(payload),
//       "muteHttpExceptions": true
//     };

//    // const response = UrlFetchApp.fetch(`https://graph.facebook.com/v18.0/${PHONE_NUMBER_ID}/messages`, options);
//     const response = UrlFetchApp.fetch(`https://graph.facebook.com/v21.0/${PHONE_NUMBER_ID}/messages`, options);
//     const result = JSON.parse(response.getContentText());

//     if (result.error) {
//       Logger.log("FAILED! Meta Error: " + result.error.message);
//     } else {
//       Logger.log("SUCCESS! Message ID: " + result.messages[0].id);
//     }

//   } catch (err) {
//     Logger.log("Critical Script Error: " + err.toString());
//   }
// }



// /*const TEMPLATE_NAME1 = "complaint_resolution_report1"
// function Complaint_resolution(e) {
//   // Wait for the sheet to stabilize after form submission
//   Utilities.sleep(2000); 
// try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const responseSheet = ss.getSheetByName("Complaint Report"); 
//     const lastRow = responseSheet.getLastRow();
    
//     const headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
//     const rowData = responseSheet.getRange(lastRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];

//     let formData = {};
//     headers.forEach((header, index) => {
//       formData[header] = rowData[index];
//     });

//     let status = (formData["Complaint Resolved"] || formData["Complaint Resolved?"] || "").toString().trim().toLowerCase();
//     if (status !== "yes") {
//       Logger.log("Skipping: Status is '" + status + "', not 'yes'.");
//       return;
//     }

//     const complaintId = formData["Complaint ID"];
//     const fileIdRaw = formData["Upload Report PDF or JPEG"] || "";

//     // LOOKUP: Phone Number from "Cleaned Data"
//     const cleanedSheet = ss.getSheetByName("Cleaned Data");
//     const cleanedValues = cleanedSheet.getDataRange().getValues();
//     let finalPhone = "";

//     for (let i = 1; i < cleanedValues.length; i++) {
//       // Check Column B (index 1) for a match with the Complaint ID
//       if (String(cleanedValues[i][1]).trim() == String(complaintId).trim()) { 
//         finalPhone = cleanedValues[i][12]; // Get phone from Column M (index 12)
//         break;
//       }
//     }

//     if (!finalPhone) {
//       Logger.log("ERROR: No phone found in 'Cleaned Data' for ID: " + complaintId);
//       return;
//     }

//     // Format Phone to include country code
//     let cleanPhone = finalPhone.toString().replace(/\D/g, "");
//     if (cleanPhone.length === 10) cleanPhone = "91" + cleanPhone;

//     // Link Extraction and Permission Setting
//     let finalViewLink = "https://drive.google.com";
//     const fileIdMatch = fileIdRaw.match(/[-\w]{25,}/); 
//     if (fileIdMatch) {
//       const driveFileId = fileIdMatch[0];
//       try {
//         const file = DriveApp.getFileById(driveFileId);
//         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
//         finalViewLink = "https://drive.google.com/open?id=" + driveFileId;
//       } catch(err) {
//         // Fallback to the direct open link if specific file retrieval fails
//         finalViewLink = "https://drive.google.com/open?id=" + driveFileId;
//       }
//     }

//     // Send WhatsApp Message using NAMED PARAMETERS
    
//     // Send WhatsApp Message using NAMED PARAMETERS
//     const payload = {
//       "messaging_product": "whatsapp",
//       "to": cleanPhone,
//       "type": "template",
//       "template": {
//         "name": TEMPLATE_NAME1,
//         "language": { "code": "en_US" }, 
//         "components": [{
//           "type": "body",
//           "parameters": [
//             { 
//               "type": "text", 
//               "parameter_name": "complaint_id", // CHANGED to lowercase 'id' to match your template
//               "text": String(complaintId) 
//             },
//             { 
//               "type": "text", 
//               "parameter_name": "report_link", // Matches your template {{report_link}}
//               "text": String(finalViewLink) 
//             }
//           ]
//         }]
//       }
//     };

//     const options = {
//       "method": "post",
//       "contentType": "application/json",
//       "headers": { "Authorization": "Bearer " + META_ACCESS_TOKEN },
//       "payload": JSON.stringify(payload),
//       "muteHttpExceptions": true
//     };

//    // const response = UrlFetchApp.fetch(`https://graph.facebook.com/v18.0/${PHONE_NUMBER_ID}/messages`, options);
//     const response = UrlFetchApp.fetch(`https://graph.facebook.com/v21.0/${PHONE_NUMBER_ID}/messages`, options);
//     const result = JSON.parse(response.getContentText());

//     if (result.error) {
//       Logger.log("FAILED! Meta Error: " + result.error.message);
//     } else {
//       Logger.log("SUCCESS! Message ID: " + result.messages[0].id);
//     }

//   } catch (err) {
//     Logger.log("Critical Script Error: " + err.toString());
//   }
// }*/