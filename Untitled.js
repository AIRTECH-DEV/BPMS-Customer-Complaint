const TEMPLATE_NAME1 = "complaint_resolution_report1";
const REPORT_SHEET_ID1 ="1oJMb4ZIbSOdSZZLD6DTIkl8-ItVb_AIZJlv4JPN6Jms";
//const Report_SHEET_ID =MASTER_SHEET_ID;


// ============================================================
//  STEP A — Called from submitComplaint() after saving the row
//  Schedules the WhatsApp to fire exactly 5 minutes later
// ============================================================
function scheduleWhatsAppIn5Min_(targetRow) {
  const fireAt = new Date(new Date().getTime() + 3 * 60 * 1000);
    Logger.log("  here fun comes ");
  // ✅ Capture ssId NOW while there's an active context
  const ssId = REPORT_SHEET_ID1;


  const trigger = ScriptApp.newTrigger('runScheduledWhatsApp_')
    .timeBased()
    .at(fireAt)
    .create();

  // Store BOTH row AND spreadsheet ID together
  PropertiesService.getScriptProperties().setProperty(
    'TRIGGER_ROW_' + trigger.getUniqueId(),
    JSON.stringify({ row: targetRow, ssId: ssId })
  );

  Logger.log('⏰ Scheduled WhatsApp for row ' + targetRow + ' | ssId: ' + ssId);
}
// ============================================================
//  STEP B — Runs automatically after 5 minutes
//  Apps Script passes the event object (e) which has triggerUid
// ============================================================
function runScheduledWhatsApp_(e) {
  Logger.log('🔔 Trigger fired. UID: ' + (e ? e.triggerUid : 'NO EVENT'));

  const props  = PropertiesService.getScriptProperties();
  const key    = 'TRIGGER_ROW_' + e.triggerUid;
  const stored = props.getProperty(key);

  if (!stored) {
    Logger.log('⚠️ No data found for key: ' + key);
    return;
  }

  const parsed   = JSON.parse(stored);
  const targetRow = parsed.row;
  const ssId      = parsed.ssId;

  Logger.log('Row: ' + targetRow + ' | ssId: ' + ssId);

  // Clean up
  props.deleteProperty(key);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getUniqueId() === e.triggerUid) ScriptApp.deleteTrigger(t);
  });

  // ✅ Pass ssId directly — no getActiveSpreadsheet() needed
  sendComplaintResolutionWhatsApp_(targetRow, ssId);
}




function sendComplaintResolutionWhatsApp_(targetRow, ssId) {
  try {
    if (!ssId) ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
 
    Logger.log("=== START (row " + targetRow + ") ===");
 
    // ── STEP 1: Check Column S status ──
    // Always open fresh by ID to avoid cached reads
    const ss1 = SpreadsheetApp.openById(ssId);
    const sheet1 = ss1.getSheetByName("Complaint Report1");
    const statusRaw = sheet1.getRange(targetRow, 19).getValue();
    const status = statusRaw.toString().trim().toLowerCase();
 
    Logger.log("Column S: '" + status + "'");
 
    if (status !== "yes") {
      Logger.log("⚠️ Status not 'yes'. Stopping.");
      return;
    }
    Logger.log("✅ Status = 'yes'.");
 
    // ── STEP 2: Read headers + row data ──
    const ss2 = SpreadsheetApp.openById(ssId);
    const sheet2 = ss2.getSheetByName("Complaint Report1");
 
    const headers = sheet2
      .getRange(1, 1, 1, sheet2.getLastColumn())
      .getValues()[0]
      .map(h => h.toString().trim());
 
    const rowData = sheet2
      .getRange(targetRow, 1, 1, sheet2.getLastColumn())
      .getValues()[0];
 
    let formData = {};
    headers.forEach((h, i) => { formData[h] = rowData[i]; });
 
    // ── STEP 3: Complaint ID ──
    const complaintId = formData["Complaint ID"];
    Logger.log("Complaint ID: " + complaintId);
    if (!complaintId) { Logger.log("❌ No Complaint ID. Stopping."); return; }
 
    // ── STEP 4: Phone from Cleaned Data ──
    const ss3 = SpreadsheetApp.openById(ssId);
    const cleanedSheet = ss3.getSheetByName("Cleaned Data");
    const cleanedValues = cleanedSheet.getDataRange().getValues();
    let finalPhone = "";
 
    for (let i = 1; i < cleanedValues.length; i++) {
      if (String(cleanedValues[i][1]).trim() === String(complaintId).trim()) {
        finalPhone = cleanedValues[i][12]; // Column M
        Logger.log("✅ Phone at row " + (i + 1) + ": " + finalPhone);
        break;
      }
    }
 
    if (!finalPhone) {
      Logger.log("❌ No phone for Complaint ID: " + complaintId + ". Stopping.");
      return;
    }
 
    // ── STEP 5: Format Phone ──
    let cleanPhone = finalPhone.toString().replace(/\D/g, "");
    if (cleanPhone.length === 10) cleanPhone = "91" + cleanPhone;
    Logger.log("Phone: " + cleanPhone);
 
    // ── STEP 6: POLL Column W for HTTPS link ──
    // Each attempt opens a fresh spreadsheet instance to defeat caching.
    // reportLink is set ONCE here and used directly — never re-read.
 
    let reportLink = "";
    let attempts = 0;
    const maxAttempts = 18;      // 18 × 10s = 3 minutes
    const pollInterval = 10000;
 
    Logger.log("--- Polling Column W ---");
 
    while (attempts < maxAttempts) {
      const freshSs = SpreadsheetApp.openById(ssId);
      const freshSheet = freshSs.getSheetByName("Complaint Report1");
 
      const range = freshSheet.getRange(targetRow, 23); // Column W
      const formula = range.getFormula();
      const value = range.getValue().toString();
 
      Logger.log("Attempt " + (attempts + 1) + "/" + maxAttempts +
                 " | Formula: '" + formula + "' | Value: '" + value + "'");
 
      if (formula && formula.includes("https")) {
        const urlMatch = formula.match(/"(https[^"]+)"/);
        if (urlMatch && urlMatch[1]) {
          reportLink = urlMatch[1];
          Logger.log("✅ Link found: " + reportLink);
          break;
        }
      }
 
      Logger.log("Not ready. Waiting 10s...");
      Utilities.sleep(pollInterval);
      attempts++;
    }
 
    // ── STEP 7: Validate link ──
    if (!reportLink || !reportLink.startsWith("https")) {
      Logger.log("❌ ABORT: No valid link after " + maxAttempts + " attempts. No WhatsApp sent.");
      return;
    }
    Logger.log("✅ Link confirmed: " + reportLink);
 
    // ── STEP 8: Build payload ──
    const payload = {
      "messaging_product": "whatsapp",
      "to": cleanPhone,
      "type": "template",
      "template": {
        "name": TEMPLATE_NAME1,
        "language": { "code": "en_US" },
        "components": [{
          "type": "body",
          "parameters": [
            { "type": "text", "parameter_name": "complaint_id", "text": String(complaintId) },
            { "type": "text", "parameter_name": "report_link",  "text": String(reportLink) }
          ]
        }]
      }
    };
 
    Logger.log("Payload:\n" + JSON.stringify(payload, null, 2));
 
    // ── STEP 9: Send WhatsApp ──
    const options = {
      "method": "post",
      "contentType": "application/json",
      "headers": { "Authorization": "Bearer " + META_ACCESS_TOKEN },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
 
    const response = UrlFetchApp.fetch(
      `https://graph.facebook.com/v21.0/${PHONE_NUMBER_ID}/messages`,
      options
    );
 
    const result = JSON.parse(response.getContentText());
 
    Logger.log("HTTP Code   : " + response.getResponseCode());
    Logger.log("Raw Response: " + response.getContentText());
 
    if (result.error) {
      Logger.log("❌ FAILED: " + result.error.code + " - " + result.error.message);
    } else {
      Logger.log("✅ SENT! ID: " + result.messages[0].id + " | Status: " + result.messages[0].message_status);
    }
 
    Logger.log("=== SUMMARY ===");
    Logger.log("Row         : " + targetRow);
    Logger.log("Complaint ID: " + complaintId);
    Logger.log("Phone       : " + cleanPhone);
    Logger.log("Link        : " + reportLink);
    Logger.log("Attempts    : " + (attempts + 1) + " / " + maxAttempts);
    Logger.log("Result      : " + (result.error ? "❌ " + result.error.message : "✅ " + result.messages[0].id));
    Logger.log("=== END ===");
 
  } catch (err) {
    Logger.log("❌ Critical Error: " + err.toString());
  }
}
 
// ============================================================
//  HELPER - Delete all triggers for a given function name
// ============================================================
function deleteTriggerByFunctionName_(funcName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(t);
      Logger.log("🗑️ Deleted trigger for: " + funcName);
    }
  }
}




function Test_Complaint_Resolution1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Complaint Report1");
  const lastRow = sheet.getLastRow();
  Logger.log("Manual test running on last row: " + lastRow);
  sendComplaintResolutionWhatsApp_(lastRow);
}


function TEST_ScheduleTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Complaint Report1");
  const lastRow = sheet.getLastRow();
  
  Logger.log("Testing schedule for row: " + lastRow);
  scheduleWhatsAppIn5Min_(lastRow);
  Logger.log("✅ Trigger created. Check back in 5 min under Executions tab.");
}


function getMySpreadsheetId() {
  Logger.log(SpreadsheetApp.getActiveSpreadsheet().getId());
}