// ─────────────────────────────────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────────────────────────────────

var COMPLAINT_TEMPLATE_DOC_ID = "1_dsXZdnwCajnmrfk4w3BgJI9-rz_vdDsCEJhHDisELo";
var COMPLAINT_FOLDER_ID       = "1G8jvcUvWixpq_6d5X8YGveIldb7pp_Il";
var TEST_COMPLAINT_FOLDER_ID  = "1gS6OUjSOffA-9CslWReJlkthZ94dZUas";

var MAX_AC_ROWS_PER_PAGE = 7;

// Layout tuning
var AC_FIRST_PAGE_BUDGET_PT   = 390; // first AC block page after customer section
var AC_OTHER_PAGE_BUDGET_PT   = 560; // continuation pages
var PAGE_TOP_SPACER_PT        = 28;
var SECTION_HEADER_BEFORE_PT   = 6;
var SECTION_HEADER_AFTER_PT    = 4;
var PAGE_BREAK_GAP_PT          = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PROCESS
// ─────────────────────────────────────────────────────────────────────────────
function processPendingPDFs() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName("Complaint Report1");
  const folder  = DriveApp.getFolderById(COMPLAINT_FOLDER_ID);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
  const proofRichTexts = sheet.getRange(2, 25, lastRow - 1, 1).getRichTextValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 2;
    const id = row[1];
    const statusValue = row[22] ? row[22].toString().trim() : "";
    if (statusValue !== "" || !id) continue;

    const richProof = proofRichTexts[i][0];
    const linkUrl   = richProof ? richProof.getLinkUrl() : null;
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

  const row    = sheet.getRange(lastRow, 1, 1, 26).getValues()[0];
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

  // Payment proof resolved early so we can paginate intelligently
  const payProofVal = row[24] ? row[24].toString().trim() : "";
  const skipProof   = ["", "n/a", "no proof uploaded", "no proof", "none"];
  const hasProof    = payProofVal && !skipProof.includes(payProofVal.toLowerCase());

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
  body.setMarginTop(20); 

  // 2. Timestamp
  const now = new Date();
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const tsFormatted =
    now.getDate() + " " + MONTHS[now.getMonth()] + " " + now.getFullYear() +
    " | " + ("0"+now.getHours()).slice(-2) + ":" + ("0"+now.getMinutes()).slice(-2);

  // 3. Title row
  // const titleTable = body.appendTable([['','']]);
  const firstPara = body.getParagraphs()[0];
if (firstPara) {
  firstPara.setSpacingBefore(0).setSpacingAfter(0);
}
  const titleTable = body.appendTable([['','']]);
titleTable.setBorderWidth(0);
titleTable.setAttributes({
  [DocumentApp.Attribute.SPACING_BEFORE]: 0,
  [DocumentApp.Attribute.SPACING_AFTER]: 0
});
  titleTable.setBorderWidth(0).setBorderColor(WHITE);

  const lc = titleTable.getCell(0,0);
  lc.setPaddingTop(4).setPaddingBottom(2).setPaddingLeft(0).setPaddingRight(4);
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
  rc.setPaddingTop(4).setPaddingBottom(0).setPaddingLeft(4).setPaddingRight(0);
  rc.setBackgroundColor(WHITE);
  const dp = rc.getChild(0).asParagraph();
  dp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setSpacingBefore(0).setSpacingAfter(0).clear();
  dp.appendText(tsFormatted)
    .setFontFamily('Courier New').setFontSize(11).setBold(true).setForegroundColor(AMBER);

  // Red divider
  const redLine = body.appendParagraph('');
  redLine.setSpacingBefore(0).setSpacingAfter(0);
  const redLineAttr = {};
  redLineAttr[DocumentApp.Attribute.BORDER_WIDTH]  = 1;
  redLineAttr[DocumentApp.Attribute.BORDER_COLOR]  = RED;
  redLineAttr[DocumentApp.Attribute.SPACING_BEFORE] = 1;
  redLineAttr[DocumentApp.Attribute.SPACING_AFTER]  = 1;
  redLine.setAttributes(redLineAttr);

  // 4. Complaint ID
  const idPara = body.appendParagraph(id.toString());
  idPara.setSpacingBefore(0)
  idPara.setSpacingBefore(0).setSpacingAfter(3);
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
    ["PAYMENT RECEIVED",  row[19] ? row[19].toString() : "N/A"],
    ["REPORT TYPE",      row[23] ? row[23].toString() : "N/A"],
  ];
  const custTable = body.appendTable(custData);
  custTable.setBorderWidth(0);
  styleInfoTable(custTable, BGRAY, GRAY, DARK, GREEN, AMBER, LGRAY, WHITE);
  body.appendParagraph('').setSpacingAfter(6);

  // 6. AC UNIT DETAILS
  appendSvcHeader(body, 'AC UNIT DETAILS', RED);

  const models          = splitCSV(row[7]);
  const serials         = splitCSV(row[8]);
  const locations       = splitCSV(row[9]);   // ✅ back
  const machTypes       = splitCSV(row[10]);
  const gasTypes        = splitCSV(row[11]);
  const problems        = splitCSV(row[12]);
  const actions         = splitCSV(row[13]);
  const brands          = splitCSV(row[15]);
  const problemStatuses = splitCSV(row[25]);

  const acHeader = [
    "S/N",
    "MACHINE BRAND",
    "MODEL",
    "SERIAL NO",
    "LOCATION",
    "MACHINE TYPE",
    "GAS TYPE",
    "PROBLEM",
    "ACTION TAKEN",
    "PROBLEM STATUS"
  ];

  const acCount = Math.max(
    models.length,
    serials.length,
    locations.length,
    machTypes.length,
    gasTypes.length,
    problems.length,
    actions.length,
    brands.length,
    problemStatuses.length,
    1
  );

  const acDataRows = [];
  for (let r = 0; r < acCount; r++) {
    acDataRows.push([
      String(r + 1),
      brands[r]          || "—",
      models[r]          || "—",
      serials[r]         || "—",
      locations[r]       || "—",
      machTypes[r]       || "—",
      gasTypes[r]        || "—",
      problems[r]        || "—",
      actions[r]         || "—",
      problemStatuses[r] || "—"
    ]);
  }

  const acFontSize = getAutoAcFontSize_(acDataRows);

  // Dynamic column widths
  const acColWidths = buildAcColumnWidths_(acHeader, acDataRows);

  // Smart pagination: estimate row heights and keep page-safe spacing
  const chunks = paginateAcRows_(acDataRows, acColWidths, hasProof);

  for (let i = 0; i < chunks.length; i++) {
    if (i > 0) {
      body.appendPageBreak();
      appendPageTopSpacer(body);
      appendSvcHeader(body, 'AC UNIT DETAILS (continued)', RED);
    }

    const chunkRows = [acHeader].concat(chunks[i].rows);
    const acTable   = body.appendTable(chunkRows);
    acTable.setBorderWidth(1).setBorderColor('#AAAAAA');

    styleACTable(acTable, RED, DARK, GRAY, WHITE, LGRAY, acFontSize, acColWidths);

    body.appendParagraph('').setSpacingAfter(4);
  }

  // 7. PAYMENT PROOF + SIGNATURES
  const paymentSectionHeight = estimatePaymentSectionHeight_(hasProof);
  const lastChunk            = chunks[chunks.length - 1] || { height: 0, limit: AC_OTHER_PAGE_BUDGET_PT };

  // If the last AC page cannot safely hold payment/signature section, move to next page.
  if ((lastChunk.height + paymentSectionHeight) > lastChunk.limit) {
    body.appendPageBreak();
    appendPageTopSpacer(body);
  } else {
    body.appendParagraph('').setSpacingAfter(6);
  }

  if (hasProof) {
    const proofHdrPara = body.appendParagraph('PAYMENT PROOF');
    proofHdrPara.setSpacingBefore(6).setSpacingAfter(8);
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
    } catch (e) {
      Logger.log("Payment proof error: " + e);
      const proofLinkPara = body.appendParagraph('View Payment Proof');
      proofLinkPara.editAsText()
        .setFontFamily('Arial')
        .setFontSize(10)
        .setForegroundColor('#1155CC')
        .setUnderline(true)
        .setLinkUrl(payProofVal);
    }

    body.appendParagraph('').setSpacingAfter(10);
  }

  const sigHdrPara = body.appendParagraph('SIGNATURES');
  sigHdrPara.setSpacingBefore(4).setSpacingAfter(10);
  sigHdrPara.editAsText()
    .setFontFamily('Arial').setFontSize(10).setBold(true).setForegroundColor(RED);

  const sigTable = body.appendTable([['', '']]);
  sigTable.setBorderWidth(0);

  // Customer signature
  const custSig = sigTable.getCell(0, 0);
  custSig.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(0).setPaddingRight(10);
  custSig.setBackgroundColor(WHITE);
  const custLbl = custSig.appendParagraph("Customer Signature");
  custLbl.setSpacingBefore(0).setSpacingAfter(8);
  custLbl.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(GRAY);

  if (row[20] && row[20].toString().includes("base64")) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(row[20].toString().split(",")[1]), "image/png"
      );
      custSig.appendImage(blob).setWidth(150).setHeight(75);
    } catch (e) {
      custSig.appendParagraph("(signature unavailable)").editAsText()
        .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
      Logger.log("Customer sig error: " + e);
    }
  } else {
    custSig.appendParagraph("Not provided").editAsText()
      .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
  }

  // Technician signature
  const techSig = sigTable.getCell(0, 1);
  techSig.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(10).setPaddingRight(0);
  techSig.setBackgroundColor(WHITE);
  const techLbl = techSig.appendParagraph("Technician Signature");
  techLbl.setSpacingBefore(0).setSpacingAfter(8);
  techLbl.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor(GRAY);

  if (row[21] && row[21].toString().includes("base64")) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(row[21].toString().split(",")[1]), "image/png"
      );
      techSig.appendImage(blob).setWidth(150).setHeight(75);
    } catch (e) {
      techSig.appendParagraph("(signature unavailable)").editAsText()
        .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
      Logger.log("Tech sig error: " + e);
    }
  } else {
    techSig.appendParagraph("Not provided").editAsText()
      .setFontFamily('Arial').setFontSize(10).setForegroundColor(LGRAY);
  }

  body.appendParagraph('').setSpacingAfter(8);

  // 8. Export PDF
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
  spacer.setSpacingBefore(0).setSpacingAfter(PAGE_TOP_SPACER_PT);
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
  para.setSpacingBefore(SECTION_HEADER_BEFORE_PT).setSpacingAfter(SECTION_HEADER_AFTER_PT);
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

function styleACTable(table, RED, DARK, GRAY, WHITE, LGRAY, acFontSize, colWidths) {
  const NUM_COLS = 10;

  for (let c = 0; c < NUM_COLS; c++) {
    try {
      table.setColumnWidth(c, colWidths[c]);
    } catch (e) {
      Logger.log("Column width set failed for col " + c + ": " + e);
    }
  }

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    const isHeader = (r === 0);

    for (let c = 0; c < NUM_COLS; c++) {
      const cell = row.getCell(c);

      cell.setBackgroundColor(isHeader ? RED : WHITE)
          .setPaddingTop(isHeader ? 5 : 4)
          .setPaddingBottom(isHeader ? 5 : 4)
          .setPaddingLeft(4)
          .setPaddingRight(4);

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
}

function getAutoAcFontSize_(rows) {
  let maxLen = 0;
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      maxLen = Math.max(maxLen, (rows[r][c] || '').toString().length);
    }
  }

  if (maxLen > 120) return 8;
  if (maxLen > 80)  return 9;
  if (maxLen > 50)  return 9.5;
  return 10;
}

function buildAcColumnWidths_(header, rows) {
  // Min / max widths in points
  const minW = [28, 72, 82, 78, 78, 62, 52, 108, 110, 70];
  const maxW = [36, 104, 125, 110, 118, 90, 72, 168, 168, 98];
  const weights = [0.6, 1.1, 1.3, 1.2, 1.1, 0.9, 0.7, 2.4, 2.4, 1.0];

  const maxLens = header.map(h => (h || '').toString().length);

  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      maxLens[c] = Math.max(maxLens[c], (rows[r][c] || '').toString().length);
    }
  }

  let raw = [];
  let total = 0;

  for (let c = 0; c < 10; c++) {
    const lenFactor = Math.min(maxLens[c], 80);
    let w = minW[c] + (lenFactor * weights[c] * 1.45);
    w = Math.max(minW[c], Math.min(maxW[c], w));
    raw.push(w);
    total += w;
  }

  const targetTotal = 520; // a safe printable width for an A4-style report body
  const scale = targetTotal / total;
  let widths = raw.map(w => Math.round(w * scale));

  // Fix rounding drift
  let diff = targetTotal - widths.reduce((a, b) => a + b, 0);
  let idx = 0;
  while (diff !== 0 && idx < 200) {
    for (let c = 0; c < widths.length && diff !== 0; c++) {
      const step = diff > 0 ? 1 : -1;
      const next = widths[c] + step;
      if (next >= minW[c] && next <= maxW[c]) {
        widths[c] = next;
        diff -= step;
      }
    }
    idx++;
  }

  return widths;
}

function estimateAcRowHeight_(row, colWidths) {
  let maxLines = 1;

  for (let c = 0; c < row.length; c++) {
    const txt = (row[c] || '').toString().trim();
    const charCapacity = Math.max(6, Math.floor(colWidths[c] / 5.8)); // rough chars per line
    const lines = Math.max(1, Math.ceil(txt.length / charCapacity));
    maxLines = Math.max(maxLines, lines);
  }

  // base row height + wrap height
  return 15 + ((maxLines - 1) * 9);
}

function paginateAcRows_(acDataRows, colWidths, hasProof) {
  const chunks = [];

  let currentRows = [];
  let currentHeight = 0;
  let currentLimit = AC_FIRST_PAGE_BUDGET_PT;
  let pageIndex = 0;

  for (let i = 0; i < acDataRows.length; i++) {
    const row = acDataRows[i];
    const rowHeight = estimateAcRowHeight_(row, colWidths);

    const tableHeaderHeight = (currentRows.length === 0) ? 18 : 0;
    const projectedHeight = currentHeight + tableHeaderHeight + rowHeight;

    if (currentRows.length > 0 && projectedHeight > currentLimit) {
      chunks.push({
        rows: currentRows,
        height: currentHeight,
        limit: currentLimit,
        pageIndex: pageIndex
      });

      pageIndex++;
      currentRows = [];
      currentHeight = 0;
      currentLimit = AC_OTHER_PAGE_BUDGET_PT;
    }

    if (currentRows.length === 0) {
      currentHeight += 18; // table header
    }

    currentRows.push(row);
    currentHeight += rowHeight;
  }

  if (currentRows.length > 0) {
    chunks.push({
      rows: currentRows,
      height: currentHeight,
      limit: currentLimit,
      pageIndex: pageIndex
    });
  }

  return chunks;
}

function estimatePaymentSectionHeight_(hasProof) {
  // If proof is an image, the section needs much more space.
  // These are safe layout estimates for deciding page breaks.
  if (hasProof) return 250;
  return 140;
}