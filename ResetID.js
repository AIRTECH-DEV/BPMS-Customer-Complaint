 function fixAndResetFromRow3() {
   const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Complaint Entry');
   const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
   const idCol = headers.indexOf('Complaint ID') + 1;

// //   // Only clear if there are rows below row 2
  if (sh.getLastRow() > 2) {
     sh.getRange(4, idCol, sh.getLastRow() - 2).clearContent();
   }

   // Force counter to 0
  const props = PropertiesService.getScriptProperties();
   props.setProperty('complaint_counter', '0');

   const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
   const lastRow = sh.getLastRow();

// //   // Only loop if there are rows from 3 onwards
   if (lastRow >= 4) {
    for (let r = 4; r <= lastRow; r++) {
      let counter = Number(props.getProperty('complaint_counter'));
      counter += 1;
       const datePart = Utilities.formatDate(new Date(), tz, 'yyMMdd');
       const seqPart = Utilities.formatString('%05d', counter);
       const id = `CMP-${datePart}-${seqPart}`;
       sh.getRange(r, idCol).setValue(id);
       props.setProperty('complaint_counter', String(counter));
     }
   }

   SpreadsheetApp.getUi().alert('Done! Row 4 onwards reset. Next new entry continues from there.');
 }