// // ─────────────────────────────────────────────────────────────────────────────
// // CONSTANTS
// // ─────────────────────────────────────────────────────────────────────────────
// const COMPLAINT_TEMPLATE_DOC_ID = "1_dsXZdnwCajnmrfk4w3BgJI9-rz_vdDsCEJhHDisELo";
// const COMPLAINT_FOLDER_ID       = "1tI-b7z7TSlgKoJ3PMAUdihjn10XoytUH";
// const TEST_COMPLAINT_FOLDER_ID  = "1gS6OUjSOffA-9CslWReJlkthZ94dZUas";

// // ─────────────────────────────────────────────────────────────────────────────
// // MAIN — loops sheet rows, skips done/empty, calls builder
// // ─────────────────────────────────────────────────────────────────────────────
// /*function processPendingPDFs() {
//   const ss     = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet  = ss.getSheetByName("Complaint Report1");
//   const folder = DriveApp.getFolderById(COMPLAINT_FOLDER_ID);

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 2) return;

//   const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

//   for (let i = 0; i < data.length; i++) {
//     const row      = data[i];
//     const rowIndex = i + 2;
//     const id       = row[1]; // Column B — Complaint ID

//     const statusValue = row[18] ? row[18].toString().trim() : "";
//     if (statusValue !== "" || !id) continue;

//     const statusCell = sheet.getRange(rowIndex, 19); // Column S

//     try {
//       statusCell.setValue("GENERATING...");
//       SpreadsheetApp.flush();

//       const pdfFile = buildComplaintDocAndExportPDF(row, id, folder, false);

//       statusCell.setFormula('=HYPERLINK("' + pdfFile.getUrl() + '", "View Report")');
//       SpreadsheetApp.flush();

//     } catch (err) {
//       statusCell.setValue("ERROR");
//       Logger.log("Error on row " + rowIndex + ": " + err.message);
//     }
//   }
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // TEST — runs against the last row of the sheet
// // ─────────────────────────────────────────────────────────────────────────────
// function TEST_processPendingPDFs() {
//   const ss      = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet   = ss.getSheetByName("Complaint Report1");
//   const lastRow = sheet.getLastRow();

//   if (lastRow < 2) { Logger.log("No data rows found."); return; }

//   const row    = sheet.getRange(lastRow, 1, 1, 21).getValues()[0];
//   const id     = row[1];
//   const folder = DriveApp.getFolderById(TEST_COMPLAINT_FOLDER_ID);

//   Logger.log("=== TEST STARTED ===");
//   Logger.log("Row: " + lastRow + " | Complaint ID: " + id);

//   const pdfFile = buildComplaintDocAndExportPDF(row, id, folder, true);

//   Logger.log("=== PDF CREATED ===");
//   Logger.log("PDF Link: https://drive.google.com/file/d/" + pdfFile.getId() + "/view");
//   Logger.log("Folder:   https://drive.google.com/drive/folders/" + TEST_COMPLAINT_FOLDER_ID);
//   Logger.log("=== DONE ===");
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // CORE — copies your letterhead template, writes content, exports PDF
// // ─────────────────────────────────────────────────────────────────────────────
// function buildComplaintDocAndExportPDF(row, id, targetFolder, isTest) {

//   // ── Colors ──
//   const RED   = '#D0312D';
//   const DARK  = '#1A1A1A';
//   const GRAY  = '#777777';
//   const LGRAY = '#AAAAAA';
//   const AMBER = '#C8860A';
//   const GREEN = '#1A7A4A';
//   const WHITE = '#FFFFFF';
//   const BGRAY = '#F5F5F5';

//   // ── 1. Copy your letterhead template (header + footer + watermark preserved) ──
//   const tempName   = 'Temp_' + id + '_' + new Date().getTime();
//   const tempFolder = DriveApp.getFolderById(
//     isTest ? TEST_COMPLAINT_FOLDER_ID : COMPLAINT_FOLDER_ID
//   );
//   const copyFile = DriveApp.getFileById(COMPLAINT_TEMPLATE_DOC_ID).makeCopy(tempName, tempFolder);
//   const docId    = copyFile.getId();
//   const doc      = DocumentApp.openById(docId);
//   const body     = doc.getBody();

//   // Clear only the body — header and footer from template stay untouched
//   body.clear();

//   // ── 2. Format timestamp ──
//   const now    = new Date();
//   const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
//   const tsFormatted =
//     now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear() +
//     " | " +
//     ("0" + now.getHours()).slice(-2) + ":" + ("0" + now.getMinutes()).slice(-2);

//   // ── 3. TITLE ROW — "SERVICE REPORT" left, timestamp right ──
//   const titleTable = body.appendTable([['', '']]);
//   titleTable.setBorderWidth(0);
//   titleTable.setBorderColor(WHITE);

//   const leftCell = titleTable.getCell(0, 0);
//   leftCell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(0).setPaddingRight(4);
//   leftCell.setBackgroundColor(WHITE);
//   const titlePara = leftCell.getChild(0).asParagraph();
//   titlePara.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
//            .setSpacingBefore(0).setSpacingAfter(0).clear();
//   titlePara.appendText('SERVICE ')
//            .setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(DARK);
//   titlePara.appendText('REPORT')
//            .setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(RED);

//   const rightCell = titleTable.getCell(0, 1);
//   rightCell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(0);
//   rightCell.setBackgroundColor(WHITE);
//   const datePara = rightCell.getChild(0).asParagraph();
//   datePara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
//           .setSpacingBefore(0).setSpacingAfter(0).clear();
//   datePara.appendText(tsFormatted)
//           .setFontFamily('Courier New').setFontSize(11).setBold(true).setForegroundColor(AMBER);

//   // Red divider
//   const redLine = body.appendParagraph('');
//   redLine.setSpacingBefore(4).setSpacingAfter(10);
//   redLine.editAsText().setForegroundColor(RED);

//   // ── 4. ISSUE / REQUIREMENT ──
//   appendServiceSectionHeader(body, "ISSUE / REQUIREMENT", RED);

//   const issueValue = row[4] ? row[4].toString() : "N/A";
//   const issuePara  = body.appendParagraph(issueValue);
//   issuePara.setSpacingBefore(2).setSpacingAfter(12);
//   issuePara.editAsText()
//            .setFontFamily('Arial').setFontSize(13).setBold(true).setForegroundColor(DARK);

//   body.appendParagraph('').setSpacingAfter(8);

//   // ── 5. COMPLAINT DETAILS table ──
//   appendServiceSectionHeader(body, 'COMPLAINT DETAILS', RED);

//   const tableData = [
//     ["COMPLAINT ID",    id                                                            ],
//     ["CUSTOMER",        row[2]  ? row[2].toString()  : "N/A"                         ], // Col C
//     ["ADDRESS",         row[3]  ? row[3].toString()  : "N/A"                         ], // Col D
//     ["BRAND / MODEL",   (row[11] ? row[11].toString() : "N/A") + " / " +
//                         (row[7]  ? row[7].toString()  : "N/A")                       ], // Col L / H
//     ["SERIAL NO",       row[8]  ? row[8].toString()  : "N/A"                         ], // Col I
//     ["SERVICE TYPE",    row[10] ? row[10].toString() : "N/A"                         ], // Col K
//     ["TECHNICIAN",      row[13] ? row[13].toString() : "N/A"                         ], // Col N
//     ["RESOLVED STATUS", row[14] ? row[14].toString() : "N/A"                         ], // Col O
//     ["PAYMENT STATUS",  row[15] ? row[15].toString() : "N/A"                         ], // Col P
//     ["REPORT TYPE",     row[19] ? row[19].toString() : "N/A"                         ], // Col T
//   ];

//   const detailTable = body.appendTable(tableData);
//   detailTable.setBorderWidth(0);
//   styleServiceDetailTable(detailTable, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE);

//   body.appendParagraph('').setSpacingAfter(8);

//   // ── 6. PAYMENT PROOF image (Col U / index 20) — only if present ──
//   if (row[20] && row[20].toString().trim() !== "" && row[20].toString() !== "N/A") {
//     try {
//       const proofId = row[20].toString().match(/[-\w]{25,}/);
//       if (proofId) {
//         appendServiceSectionHeader(body, 'PAYMENT PROOF', RED);
//         body.appendParagraph('').setSpacingBefore(8).setSpacingAfter(4);
//         const proofBlob = DriveApp.getFileById(proofId[0]).getBlob();
//         body.appendImage(proofBlob).setWidth(250).setHeight(180);
//         body.appendParagraph('').setSpacingAfter(10);
//       }
//     } catch(e) { Logger.log("Payment proof error: " + e); }
//   }

//   // ── 7. SIGNATURES — page break then 2-column layout ──
//   body.appendPageBreak();
//   body.appendParagraph('').setSpacingBefore(8).setSpacingAfter(0);
//   appendServiceSectionHeader(body, 'SIGNATURES', RED);
//   body.appendParagraph('').setSpacingBefore(12).setSpacingAfter(6);

//   const sigTable = body.appendTable([['', '']]);
//   sigTable.setBorderWidth(0);

//   // Customer Signature — Col Q (index 16)
//   const custCell = sigTable.getCell(0, 0);
//   custCell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(0).setPaddingRight(6);
//   const custLabel = custCell.appendParagraph("Customer Signature:");
//   custLabel.editAsText()
//            .setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(GRAY);
//   if (row[16] && row[16].toString().includes("base64")) {
//     try {
//       const imgBlob = Utilities.newBlob(
//         Utilities.base64Decode(row[16].toString().split(",")[1]), "image/png"
//       );
//       custCell.appendImage(imgBlob).setWidth(150).setHeight(75);
//     } catch(e) {
//       custCell.appendParagraph("(signature unavailable)")
//               .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
//       Logger.log("Customer sig error: " + e);
//     }
//   } else {
//     custCell.appendParagraph("Not provided")
//             .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
//   }

//   // Technician Signature — Col R (index 17)
//   const techCell = sigTable.getCell(0, 1);
//   techCell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(0);
//   const techLabel = techCell.appendParagraph("Technician Signature:");
//   techLabel.editAsText()
//            .setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(GRAY);
//   if (row[17] && row[17].toString().includes("base64")) {
//     try {
//       const imgBlob = Utilities.newBlob(
//         Utilities.base64Decode(row[17].toString().split(",")[1]), "image/png"
//       );
//       techCell.appendImage(imgBlob).setWidth(150).setHeight(75);
//     } catch(e) {
//       techCell.appendParagraph("(signature unavailable)")
//               .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
//       Logger.log("Tech sig error: " + e);
//     }
//   } else {
//     techCell.appendParagraph("Not provided")
//             .editAsText().setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
//   }

//   // ── 8. Save & export as PDF via OAuth ──
//   // NOTE: No footer code here — template's Daikin footer is preserved automatically
//   doc.saveAndClose();
//   Logger.log("Doc written OK — Complaint ID: " + id);

//   const token   = ScriptApp.getOAuthToken();
//   const pdfUrl  = "https://docs.google.com/feeds/download/documents/export/Export?id="
//                   + docId + "&exportFormat=pdf";
//   const pdfResp = UrlFetchApp.fetch(pdfUrl, {
//     headers: { "Authorization": "Bearer " + token },
//     muteHttpExceptions: true
//   });

//   DriveApp.getFileById(docId).setTrashed(true);
//   Logger.log("Temp Doc deleted");

//   const suffix  = isTest ? "_ServiceReport_TEST.pdf" : "_ServiceReport.pdf";
//   const pdfName = Utilities.formatDate(
//     new Date(), "GMT+5:30", "dd_MMM_yyyy"
//   ) + "_" + id + suffix;

//   return targetFolder.createFile(pdfResp.getBlob()).setName(pdfName);
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // HELPERS
// // ─────────────────────────────────────────────────────────────────────────────

// function appendServiceSectionHeader(body, text, color) {
//   const para = body.appendParagraph(text);
//   para.setSpacingBefore(12).setSpacingAfter(6);
//   para.editAsText()
//       .setFontFamily('Arial').setFontSize(10).setBold(true)
//       .setForegroundColor(color || '#D0312D');
//   return para;
// }

// function styleServiceDetailTable(table, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE) {
//   for (let r = 0; r < table.getNumRows(); r++) {
//     const row = table.getRow(r);

//     const keyCell = row.getCell(0);
//     keyCell.setBackgroundColor(BGRAY)
//            .setPaddingTop(8).setPaddingBottom(8)
//            .setPaddingLeft(10).setPaddingRight(6);
//     keyCell.editAsText()
//            .setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(GRAY);

//     const valCell = row.getCell(1);
//     const rawVal  = valCell.getText().toLowerCase().trim();
//     valCell.setPaddingTop(8).setPaddingBottom(8)
//            .setPaddingLeft(10).setPaddingRight(6);
//     valCell.setBackgroundColor(r % 2 === 0 ? WHITE : '#FAFAFA');

//     const vt = valCell.editAsText();
//     vt.setFontFamily('Arial').setFontSize(11).setBold(false);

//     if (['resolved', 'paid', 'done', 'completed'].includes(rawVal)) {
//       vt.setForegroundColor(GREEN).setBold(true);
//     } else if (['pending', 'unpaid', 'in progress', 'yes'].includes(rawVal)) {
//       vt.setForegroundColor(AMBER).setBold(true);
//     } else if (rawVal === 'n/a' || rawVal === '') {
//       vt.setForegroundColor(LGRAY);
//     } else {
//       vt.setForegroundColor(DARK);
//     }
//   }
// }*/


// // ─────────────────────────────────────────────────────────────────────────────
// // MAIN — loops sheet rows, skips done/empty, calls builder
// // ─────────────────────────────────────────────────────────────────────────────
// function processPendingPDFs() {
//   const ss     = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet  = ss.getSheetByName("Complaint Report1");
//   const folder = DriveApp.getFolderById(COMPLAINT_FOLDER_ID);

//   const lastRow = sheet.getLastRow();
//   if (lastRow < 2) return;

//   // We fetch 25 columns (A through Y)
//   const data = sheet.getRange(2, 1, lastRow - 1, 25).getValues();

//   for (let i = 0; i < data.length; i++) {
//     const row      = data[i];
//     const rowIndex = i + 2;
//     const id       = row[1]; // Column B — Complaint ID

//     // Check Column W (index 22) for existing report link
//     const statusValue = row[22] ? row[22].toString().trim() : "";
//     if (statusValue !== "" || !id) continue;

//     const statusCell = sheet.getRange(rowIndex, 23); // Column W (23rd column)

//     try {
//       statusCell.setValue("GENERATING...");
//       SpreadsheetApp.flush();

//       const pdfFile = buildComplaintDocAndExportPDF(row, id, folder, false);

//       statusCell.setFormula('=HYPERLINK("' + pdfFile.getUrl() + '", "View Report")');
//       SpreadsheetApp.flush();

//     } catch (err) {
//       statusCell.setValue("ERROR");
//       Logger.log("Error on row " + rowIndex + ": " + err.message);
//     }
//   }
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // CORE — Writes content and exports PDF
// // ─────────────────────────────────────────────────────────────────────────────
// function buildComplaintDocAndExportPDF(row, id, targetFolder, isTest) {

//   const RED   = '#D0312D';
//   const DARK  = '#1A1A1A';
//   const GRAY  = '#777777';
//   const LGRAY = '#AAAAAA';
//   const AMBER = '#C8860A';
//   const GREEN = '#1A7A4A';
//   const WHITE = '#FFFFFF';
//   const BGRAY = '#F5F5F5';

//   const tempName   = 'Temp_' + id + '_' + new Date().getTime();
//   const tempFolder = DriveApp.getFolderById(isTest ? TEST_COMPLAINT_FOLDER_ID : COMPLAINT_FOLDER_ID);
//   const copyFile = DriveApp.getFileById(COMPLAINT_TEMPLATE_DOC_ID).makeCopy(tempName, tempFolder);
//   const docId    = copyFile.getId();
//   const doc      = DocumentApp.openById(docId);
//   const body     = doc.getBody();

//   body.clear();

//   // 1. Format timestamp
//   const now    = new Date();
//   const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
//   const tsFormatted = now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear() +
//                       " | " + ("0" + now.getHours()).slice(-2) + ":" + ("0" + now.getMinutes()).slice(-2);

//   // 2. Title Section
//   const titleTable = body.appendTable([['', '']]);
//   titleTable.setBorderWidth(0);
//   const titlePara = titleTable.getCell(0, 0).getChild(0).asParagraph().clear();
//   titlePara.appendText('SERVICE REPORT').setFontSize(22).setBold(true).setForegroundColor(RED);
//   titleTable.getCell(0, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT).clear()
//             .appendText(tsFormatted).setFontFamily('Courier New').setFontSize(11).setForegroundColor(AMBER);

//   // 3. Issue Section
//   appendServiceSectionHeader(body, "ISSUE / REQUIREMENT", RED);
//   body.appendParagraph(row[4] ? row[4].toString() : "N/A").setBold(true).setFontSize(13);

//   // 4. Detailed Table Mapping (ALL INDICES UPDATED)
//   appendServiceSectionHeader(body, 'COMPLAINT DETAILS', RED);

//   const tableData = [
//     ["COMPLAINT ID",    id                                      ], // Col B
//     ["CUSTOMER",        row[2]  ? row[2].toString()  : "N/A"    ], // Col C
//     ["ADDRESS",         row[3]  ? row[3].toString()  : "N/A"    ], // Col D
//     ["AC LOCATION",     row[9]  ? row[9].toString()  : "N/A"    ], // Col J
//     ["MACHINE TYPE",    row[10] ? row[10].toString() : "N/A"    ], // Col K (NEW)
//     ["GAS TYPE",        row[11] ? row[11].toString() : "N/A"    ], // Col L (NEW)
//     ["BRAND / MODEL",   (row[15] ? row[15].toString() : "N/A") + " / " + (row[7] ? row[7].toString() : "N/A")], // Col P / H
//     ["SERIAL NO",       row[8]  ? row[8].toString()  : "N/A"    ], // Col I
//     ["PROBLEM",         row[12] ? row[12].toString() : "N/A"    ], // Col M (NEW)
//     ["ACTION TAKEN",    row[13] ? row[13].toString() : "N/A"    ], // Col N (NEW)
//     ["SERVICE TYPE",    row[14] ? row[14].toString() : "N/A"    ], // Col O
//     ["TECHNICIAN",      row[17] ? row[17].toString() : "N/A"    ], // Col R
//     ["RESOLVED STATUS", row[18] ? row[18].toString() : "N/A"    ], // Col S
//     ["PAYMENT STATUS",  row[19] ? row[19].toString() : "N/A"    ], // Col T
//     ["REPORT TYPE",     row[23] ? row[23].toString() : "N/A"    ]  // Col X
//   ];

//   const detailTable = body.appendTable(tableData);
//   detailTable.setBorderWidth(0);
//   styleServiceDetailTable(detailTable, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE);

//   // 5. Payment Proof Image (Column Y / Index 24)
//   if (row[24] && row[24].toString().trim() !== "" && row[24].toString() !== "N/A") {
//     try {
//       const proofId = row[24].toString().match(/[-\w]{25,}/);
//       if (proofId) {
//         appendServiceSectionHeader(body, 'PAYMENT PROOF', RED);
//         const proofBlob = DriveApp.getFileById(proofId[0]).getBlob();
//         body.appendImage(proofBlob).setWidth(250).setHeight(180);
//       }
//     } catch(e) { Logger.log("Payment proof error: " + e); }
//   }

//   // 6. Signatures (Column U & V / Index 20 & 21)
//   body.appendPageBreak();
//   appendServiceSectionHeader(body, 'SIGNATURES', RED);
//   const sigTable = body.appendTable([['', '']]);
//   sigTable.setBorderWidth(0);

//   // Customer Signature
//   handleSignature(sigTable.getCell(0, 0), "Customer Signature:", row[20]);
//   // Tech Signature
//   handleSignature(sigTable.getCell(0, 1), "Technician Signature:", row[21]);

//   doc.saveAndClose();

//   // 7. Export PDF
//   const token = ScriptApp.getOAuthToken();
//   const pdfUrl = "https://docs.google.com/feeds/download/documents/export/Export?id=" + docId + "&exportFormat=pdf";
//   const pdfResp = UrlFetchApp.fetch(pdfUrl, { headers: { "Authorization": "Bearer " + token } });
  
//   DriveApp.getFileById(docId).setTrashed(true);
  
//   const pdfName = Utilities.formatDate(new Date(), "GMT+5:30", "dd_MMM_yyyy") + "_" + id + (isTest ? "_TEST.pdf" : ".pdf");
//   return targetFolder.createFile(pdfResp.getBlob()).setName(pdfName);
// }

// // ─────────────────────────────────────────────────────────────────────────────
// // HELPERS
// // ─────────────────────────────────────────────────────────────────────────────

// function handleSignature(cell, label, sigData) {
//   cell.appendParagraph(label).setBold(true).setFontSize(10).setForegroundColor('#777777');
//   if (sigData && sigData.toString().includes("base64")) {
//     try {
//       const imgBlob = Utilities.newBlob(Utilities.base64Decode(sigData.split(",")[1]), "image/png");
//       cell.appendImage(imgBlob).setWidth(150).setHeight(75);
//     } catch(e) { cell.appendParagraph("(error loading signature)"); }
//   } else {
//     cell.appendParagraph("Not provided").setItalic(true).setForegroundColor('#AAAAAA');
//   }
// }

// function appendServiceSectionHeader(body, text, color) {
//   const para = body.appendParagraph(text);
//   para.setSpacingBefore(12).setSpacingAfter(6);
//   para.editAsText().setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(color);
//   return para;
// }

// function styleServiceDetailTable(table, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE) {
//   for (let r = 0; r < table.getNumRows(); r++) {
//     const row = table.getRow(r);
//     row.getCell(0).setBackgroundColor(BGRAY).editAsText().setFontSize(9).setBold(true).setForegroundColor(GRAY);
//     const valCell = row.getCell(1);
//     const rawVal = valCell.getText().toLowerCase().trim();
//     valCell.setBackgroundColor(r % 2 === 0 ? WHITE : '#FAFAFA').editAsText().setFontSize(11);
    
//     if (['resolved', 'paid', 'yes'].includes(rawVal)) {
//       valCell.editAsText().setForegroundColor(GREEN).setBold(true);
//     } else if (['pending', 'no'].includes(rawVal)) {
//       valCell.editAsText().setForegroundColor(AMBER).setBold(true);
//     }
//   }
// }