/** CONFIG: set your responses sheet name & column header exactly as in your sheet */
const RESPONSES_SHEET_NAME = 'Complaint Entry';
const ID_HEADER = 'Complaint ID';     // the header text for the ID column
const ID_PREFIX = 'CMP';              // change if you like (e.g., 'TICKET')

/**
 * Installable trigger: runs on each form submission and writes a stable ID.
 * Set in Triggers: assignComplaintId → From spreadsheet → On form submit.
 */
function assignComplaintId(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Complaint Entry") return; // Skip other forms

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(RESPONSES_SHEET_NAME);
    if (!sh) throw new Error(`Sheet "${RESPONSES_SHEET_NAME}" not found`);

    // Find the Complaint ID column (by header name)
    const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    const headers = headerRange.getValues()[0];
    let idCol = headers.indexOf(ID_HEADER) + 1;
    if (idCol < 1) {
      // If header doesn't exist yet, create it as the last column
      idCol = sh.getLastColumn() + 1;
      sh.getRange(1, idCol).setValue(ID_HEADER);
    }

    // Row that was just submitted
    const row = e.range.getRow();

    // If an ID already exists (e.g., user edited response), do nothing
    const existing = sh.getRange(row, idCol).getValue();
    if (existing) return;

    // Get & bump a global counter (persists across submissions)
    const props = PropertiesService.getScriptProperties();
    let counter = Number(props.getProperty('complaint_counter') || '0');

    // If first time, initialize from current sheet values (max sequence found)
    if (counter === 0) {
      const idValues = sh.getRange(2, idCol, Math.max(0, sh.getLastRow() - 1)).getValues().flat();
      const re = new RegExp(`^${ID_PREFIX}-\\d{6}-(\\d{5})$`);
      let maxSeq = 0;
      idValues.forEach(v => {
        if (typeof v === 'string') {
          const m = v.match(re);
          if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
        }
      });
      counter = maxSeq; // start from current max
    }

    // Build ID: CMP-YYMMDD-##### (zero-padded sequence)
    counter += 1;
    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
    const seqPart = Utilities.formatString('%05d', counter);
    const id = `${ID_PREFIX}-${datePart}-${seqPart}`;

    // Write the ID to the submission row
    sh.getRange(row, idCol).setValue(id);

    // Persist the bumped counter
    props.setProperty('complaint_counter', String(counter));
  } finally {
    lock.releaseLock();
  }
}

function backfillComplaintIds() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RESPONSES_SHEET_NAME);
  const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
  const headers = headerRange.getValues()[0];
  const idCol = headers.indexOf(ID_HEADER) + 1;
  if (idCol < 1) throw new Error(`Column "${ID_HEADER}" not found`);

  const lastRow = sh.getLastRow();
  const idValues = sh.getRange(2, idCol, lastRow - 1).getValues().flat();

  for (let r = 2; r <= lastRow; r++) {
    if (!idValues[r - 2]) {
      // Simulate a submission event
      const e = { range: sh.getRange(r, 1) };
      assignComplaintId(e);
    }
  }
  SpreadsheetApp.getUi().alert('Backfilled Complaint IDs for existing rows.');
}