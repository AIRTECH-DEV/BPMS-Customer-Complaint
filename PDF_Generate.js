// ─────────────────────────────────────────────────────────────────────────────
// MAIN

var COMPLAINT_TEMPLATE_DOC_ID = "1_dsXZdnwCajnmrfk4w3BgJI9-rz_vdDsCEJhHDisELo";
var COMPLAINT_FOLDER_ID       = "1G8jvcUvWixpq_6d5X8YGveIldb7pp_Il";
var TEST_COMPLAINT_FOLDER_ID  = "1gS6OUjSOffA-9CslWReJlkthZ94dZUas";
var MAX_AC_ROWS_PER_PAGE = 7;
// ─────────────────────────────────────────────────────────────────────────────
function processPendingPDFs() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName("Complaint Report1");
  const folder = DriveApp.getFolderById(COMPLAINT_FOLDER_ID);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 26).getValues(); // ← 25→26
  const proofRichTexts = sheet.getRange(2, 25, lastRow - 1, 1).getRichTextValues();

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const rowIndex = i + 2;
    const id       = row[1];
    const statusValue = row[22] ? row[22].toString().trim() : "";
    if (statusValue !== "" || !id) continue;

    const richProof   = proofRichTexts[i][0];
    const linkUrl     = richProof ? richProof.getLinkUrl() : null;
    row[24] = linkUrl || row[24];

    const statusCell = sheet.getRange(rowIndex, 23);
    try {
      statusCell.setValue("GENERATING...");
      SpreadsheetApp.flush();
      const pdfFile = buildComplaintDocAndExportPDF(row, id, folder, false);
      statusCell.setFormula('=HYPERLINK("' + pdfFile.getUrl() + '","View Report")');
      SpreadsheetApp.flush();
    } catch (err) {
      statusCell.setValue("ERROR");
      Logger.log("Error on row " + rowIndex + ": " + err.message);
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// TEST
// ─────────────────────────────────────────────────────────────────────────────
function TEST_processPendingPDFs() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName("Complaint Report1");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log("No data rows found."); return; }

  const row    = sheet.getRange(lastRow, 1, 1, 26).getValues()[0]; // ← 25→26
  const id     = row[1];
  const folder = DriveApp.getFolderById(TEST_COMPLAINT_FOLDER_ID);

  const richProof = sheet.getRange(lastRow, 25).getRichTextValue();
  const linkUrl   = richProof ? richProof.getLinkUrl() : null;
  row[24] = linkUrl || row[24];
  Logger.log("Payment proof URL resolved: " + row[24]);

  Logger.log("=== TEST STARTED === Row: " + lastRow + " | ID: " + id);
  const pdfFile = buildComplaintDocAndExportPDF(row, id, folder, true);
  Logger.log("PDF: https://drive.google.com/file/d/" + pdfFile.getId() + "/view");
  Logger.log("=== DONE ===");
}

// ─────────────────────────────────────────────────────────────────────────────
// CORE PDF GENERATION
// ─────────────────────────────────────────────────────────────────────────────
function buildComplaintDocAndExportPDF(row, id, targetFolder, isTest) {

  const RED   = '#D0312D';
  const DARK  = '#1A1A1A';
  const GRAY  = '#777777';
  const LGRAY = '#AAAAAA';
  const AMBER = '#C8860A';
  const GREEN = '#1A7A4A';
  const WHITE = '#FFFFFF';
  const BGRAY = '#F5F5F5';

  // 1. Copy template
  const tempName   = 'Temp_' + id + '_' + new Date().getTime();
  const tempFolder = DriveApp.getFolderById(
    isTest ? TEST_COMPLAINT_FOLDER_ID : COMPLAINT_FOLDER_ID
  );
  const copyFile = DriveApp.getFileById(COMPLAINT_TEMPLATE_DOC_ID).makeCopy(tempName, tempFolder);
  const docId    = copyFile.getId();
  const doc      = DocumentApp.openById(docId);
  const body     = doc.getBody();
  body.clear();

  // 2. Timestamp
  const now    = new Date();
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const tsFormatted =
    now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear() +
    " | " + ("0"+now.getHours()).slice(-2) + ":" + ("0"+now.getMinutes()).slice(-2);

  // 3. Title row
  const titleTable = body.appendTable([['','']]);
  titleTable.setBorderWidth(0).setBorderColor(WHITE);

  const lc = titleTable.getCell(0,0);
  lc.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(0).setPaddingRight(4);
  lc.setBackgroundColor(WHITE);
  const tp = lc.getChild(0).asParagraph();
  tp.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
    .setSpacingBefore(0).setSpacingAfter(0).clear();
  tp.appendText('COMPLAINT ');
  const titleText = tp.editAsText();
  titleText.setFontFamily('Arial').setFontSize(22).setBold(true).setForegroundColor(DARK);
  tp.appendText('REPORT');
  const fullTitle = tp.getText();
  tp.editAsText().setForegroundColor(fullTitle.length - 6, fullTitle.length - 1, RED);

  const rc = titleTable.getCell(0,1);
  rc.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(0);
  rc.setBackgroundColor(WHITE);
  const dp = rc.getChild(0).asParagraph();
  dp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setSpacingBefore(0).setSpacingAfter(0).clear();
  dp.appendText(tsFormatted)
    .setFontFamily('Courier New').setFontSize(11).setBold(true).setForegroundColor(AMBER);

  // Red divider
  const redLine = body.appendParagraph('');
  redLine.setSpacingBefore(4).setSpacingAfter(10);
  const redLineAttr = {};
  redLineAttr[DocumentApp.Attribute.BORDER_WIDTH]   = 2;
  redLineAttr[DocumentApp.Attribute.BORDER_COLOR]   = RED;
  redLineAttr[DocumentApp.Attribute.SPACING_BEFORE] = 4;
  redLineAttr[DocumentApp.Attribute.SPACING_AFTER]  = 10;
  redLine.setAttributes(redLineAttr);

  // 4. Complaint ID
  const idPara = body.appendParagraph(id.toString());
  idPara.setSpacingBefore(0).setSpacingAfter(8);
  idPara.editAsText()
        .setFontFamily('Arial').setFontSize(15).setBold(true).setForegroundColor(DARK);

  // 5. Customer Details
  appendSvcHeader(body, 'CUSTOMER DETAILS', RED);

  const custData = [
    ["CUSTOMER NAME",    row[2]  ? row[2].toString()  : "N/A"],
    ["ADDRESS",          row[3]  ? row[3].toString()  : "N/A"],
    ["SERVICE TYPE",     row[14] ? row[14].toString() : "N/A"],
    ["TECHNICIAN",       row[17] ? row[17].toString() : "N/A"],
    ["RESOLVED STATUS",  row[18] ? row[18].toString() : "N/A"],
    ["PAYMENT RECEIVED", row[19] ? row[19].toString() : "N/A"],
    ["REPORT TYPE",      row[23] ? row[23].toString() : "N/A"],
  ];
  const custTable = body.appendTable(custData);
  custTable.setBorderWidth(0);
  styleInfoTable(custTable, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE);
  body.appendParagraph('').setSpacingAfter(8);

  // 6. AC UNIT DETAILS – chunked with repeated header
  appendSvcHeader(body, 'AC UNIT DETAILS', RED);

  const models          = splitCSV(row[7]);
  const serials         = splitCSV(row[8]);
  // row[9] (locations) — REMOVED
  const machTypes       = splitCSV(row[10]);
  const gasTypes        = splitCSV(row[11]);
  const problems        = splitCSV(row[12]);
  const actions         = splitCSV(row[13]);
  const brands          = splitCSV(row[15]);
  const problemStatuses = splitCSV(row[25]); // ← NEW: col Z (index 25)

  const acCount = Math.max(
    models.length, serials.length,
    machTypes.length, gasTypes.length, problems.length, actions.length,
    brands.length, problemStatuses.length, 1
  );

  // ← LOCATION removed, PROBLEM STATUS added at end
  const acHeader   = ["S/N", "MACHINE BRAND", "MODEL", "SERIAL NO", "MACHINE TYPE", "GAS TYPE", "PROBLEM", "ACTION TAKEN", "PROBLEM STATUS"];
  const acDataRows = [];
  for (let r = 0; r < acCount; r++) {
    acDataRows.push([
      String(r + 1),
      brands[r]           || "—",
      models[r]           || "—",
      serials[r]          || "—",
      machTypes[r]        || "—",
      gasTypes[r]         || "—",
      problems[r]         || "—",
      actions[r]          || "—",
      problemStatuses[r]  || "—"  // ← NEW
    ]);
  }

  const acFontSize = acCount <= 4 ? 10 : acCount <= 8 ? 9 : 8;

  let chunkStart   = 0;
  let isFirstChunk = true;
  while (chunkStart < acDataRows.length) {
    if (!isFirstChunk) {
      body.appendPageBreak();
      appendPageTopSpacer(body);
      appendSvcHeader(body, 'AC UNIT DETAILS (continued)', RED);
      body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(10);
    }
    const chunkEnd  = Math.min(chunkStart + MAX_AC_ROWS_PER_PAGE, acDataRows.length);
    const chunkRows = [acHeader].concat(acDataRows.slice(chunkStart, chunkEnd));
    const acTable   = body.appendTable(chunkRows);
    acTable.setBorderWidth(1).setBorderColor('#AAAAAA');
    styleACTable(acTable, RED, DARK, GRAY, WHITE, LGRAY, acFontSize);

    chunkStart   = chunkEnd;
    isFirstChunk = false;
  }

  // 7. PAYMENT PROOF + SIGNATURES – always on a fresh page
  body.appendPageBreak();
  appendPageTopSpacer(body);

  const payProofVal = row[24] ? row[24].toString().trim() : "";
  const skipProof   = ["", "n/a", "no proof uploaded", "no proof", "none"];
  const hasProof    = payProofVal && !skipProof.includes(payProofVal.toLowerCase());

  if (hasProof) {
    const proofHdrPara = body.appendParagraph('PAYMENT PROOF');
    proofHdrPara.setSpacingBefore(0).setSpacingAfter(12);
    proofHdrPara.editAsText()
      .setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(RED);

    try {
      const proofIdMatch = payProofVal.match(/[-\w]{25,}/);
      if (!proofIdMatch) throw new Error("No file ID found");
      const proofBlob = DriveApp.getFileById(proofIdMatch[0]).getBlob();
      const imgEl = body.appendImage(proofBlob);
      const origW = imgEl.getWidth();
      const origH = imgEl.getHeight();
      const scale = Math.min(260 / origW, 200 / origH, 1);
      imgEl.setWidth(Math.round(origW * scale)).setHeight(Math.round(origH * scale));
    } catch(e) {
      Logger.log("Payment proof error: " + e);
      const proofLinkPara = body.appendParagraph('View Payment Proof');
      proofLinkPara.editAsText()
        .setFontFamily('Arial').setFontSize(10).setForegroundColor('#1155CC')
        .setUnderline(true).setLinkUrl(payProofVal);
    }

    body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(20);
  }

  const sigHdrPara = body.appendParagraph('SIGNATURES');
  sigHdrPara.setSpacingBefore(0).setSpacingAfter(16);
  sigHdrPara.editAsText()
    .setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(RED);

  const sigTable = body.appendTable([['', '']]);
  sigTable.setBorderWidth(0);

  const custSig = sigTable.getCell(0, 0);
  custSig.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(0).setPaddingRight(10);
  custSig.setBackgroundColor(WHITE);
  const custLbl = custSig.appendParagraph("Customer Signature");
  custLbl.setSpacingBefore(0).setSpacingAfter(10);
  custLbl.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(GRAY);

  if (row[20] && row[20].toString().includes("base64")) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(row[20].toString().split(",")[1]), "image/png"
      );
      custSig.appendImage(blob).setWidth(150).setHeight(75);
    } catch(e) {
      custSig.appendParagraph("(signature unavailable)").editAsText()
        .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
      Logger.log("Customer sig error: " + e);
    }
  } else {
    custSig.appendParagraph("Not provided").editAsText()
      .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
  }

  const techSig = sigTable.getCell(0, 1);
  techSig.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(10).setPaddingRight(0);
  techSig.setBackgroundColor(WHITE);
  const techLbl = techSig.appendParagraph("Technician Signature");
  techLbl.setSpacingBefore(0).setSpacingAfter(10);
  techLbl.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(GRAY);

  if (row[21] && row[21].toString().includes("base64")) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(row[21].toString().split(",")[1]), "image/png"
      );
      techSig.appendImage(blob).setWidth(150).setHeight(75);
    } catch(e) {
      techSig.appendParagraph("(signature unavailable)").editAsText()
        .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
      Logger.log("Tech sig error: " + e);
    }
  } else {
    techSig.appendParagraph("Not provided").editAsText()
      .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
  }

  body.appendParagraph('').setSpacingAfter(12);

  // 8. Export PDF
  doc.saveAndClose();
  Logger.log("Doc written OK — ID: " + id);

  const token   = ScriptApp.getOAuthToken();
  const pdfUrl  = "https://docs.google.com/feeds/download/documents/export/Export?id="
                  + docId + "&exportFormat=pdf";
  const pdfResp = UrlFetchApp.fetch(pdfUrl, {
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });

  DriveApp.getFileById(docId).setTrashed(true);
  Logger.log("Temp Doc deleted");

  const suffix  = isTest ? "_ServiceReport_TEST.pdf" : "_ServiceReport.pdf";
  const pdfName = Utilities.formatDate(new Date(), "GMT+5:30", "dd_MMM_yyyy")
                  + "_" + id + suffix;
  return targetFolder.createFile(pdfResp.getBlob()).setName(pdfName);
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function appendPageTopSpacer(body) {
  const spacer = body.appendParagraph('');
  spacer.setSpacingBefore(0).setSpacingAfter(28);
  return spacer;
}

function splitCSV(value) {
  if (!value || value.toString().trim() === "" || value.toString().trim().toLowerCase() === "n/a") {
    return ["N/A"];
  }
  return value.toString().split(",").map(s => s.trim()).filter(s => s !== "");
}

function appendSvcHeader(body, text, color) {
  const para = body.appendParagraph(text);
  para.setSpacingBefore(12).setSpacingAfter(6);
  para.editAsText()
      .setFontFamily('Arial').setFontSize(10).setBold(true)
      .setForegroundColor(color || '#D0312D');
  return para;
}

function styleInfoTable(table, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);

    const keyCell = row.getCell(0);
    keyCell.setBackgroundColor(BGRAY)
           .setPaddingTop(7).setPaddingBottom(7).setPaddingLeft(10).setPaddingRight(6);
    keyCell.editAsText()
           .setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(GRAY);

    const valCell = row.getCell(1);
    const rawVal  = valCell.getText().toLowerCase().trim();
    valCell.setPaddingTop(7).setPaddingBottom(7).setPaddingLeft(10).setPaddingRight(6);
    valCell.setBackgroundColor(r % 2 === 0 ? WHITE : '#FAFAFA');

    const vt = valCell.editAsText();
    vt.setFontFamily('Arial').setFontSize(11).setBold(false);

    if (['resolved','paid','done','completed','no'].includes(rawVal)) {
      vt.setForegroundColor(GREEN).setBold(true);
    } else if (['pending','unpaid','in progress','yes'].includes(rawVal)) {
      vt.setForegroundColor(AMBER).setBold(true);
    } else if (rawVal === 'n/a' || rawVal === '') {
      vt.setForegroundColor(LGRAY);
    } else {
      vt.setForegroundColor(DARK);
    }
  }
}

function styleACTable(table, RED, DARK, GRAY, WHITE, LGRAY, acFontSize) {
  const NUM_COLS = 9; // unchanged: removed LOCATION (+0), added PROBLEM STATUS (+0) = still 9

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const isHeader = (r === 0);
    const bgColor  = isHeader ? RED : WHITE;

    for (let c = 0; c < NUM_COLS; c++) {
      const cell = row.getCell(c);
      cell.setBackgroundColor(bgColor)
          .setPaddingTop(isHeader ? 6 : 5)
          .setPaddingBottom(isHeader ? 6 : 5)
          .setPaddingLeft(5)
          .setPaddingRight(5);

      const vt = cell.editAsText();
      if (isHeader) {
        vt.setFontFamily('Arial')
          .setFontSize(7.5)
          .setBold(true)
          .setForegroundColor(WHITE);
      } else {
        vt.setFontFamily('Arial')
          .setFontSize(acFontSize)
          .setBold(c === 0)
          .setForegroundColor(c === 0 ? GRAY : DARK);

        const txt = cell.getText().trim();
        if (txt === 'N/A' || txt === '—') {
          vt.setForegroundColor(LGRAY).setBold(false);
        }
      }
    }
  }
}  //BY YASH RANE 25-04-2026